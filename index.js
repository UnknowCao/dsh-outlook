/**
 * dsh-outlook — Outlook COM automation for DSH Harness (Windows).
 *
 * Bridges the local desktop Outlook client via a PowerShell COM dispatcher
 * (lib/olk.ps1). Read operations (account, search/read mail, calendar query,
 * people/room search, free/busy) run automatically.
 *
 * Write operations that leave the machine (sending mail, dispatching meeting
 * invites) use a TWO-CALL CONFIRMATION protocol: this host-plane plugin's
 * context cannot reach the session-scoped popup answerers, so an in-plugin
 * popup would fail closed every time. Instead the first call returns
 * `needs-confirmation` with the exact facts to confirm; the calling AGENT
 * asks the user through its own in-session popup (ask_user_question), and
 * only retries with `confirmed: true` after the user allows. That retry
 * mints the one-time OLK_CONFIRM token the bridge demands before sending.
 *
 * The bridge is invoked as
 *   powershell -NoProfile -Command "Invoke-Expression (Get-Content -Raw <olk.ps1>)"
 * because group policy commonly blocks running .ps1 files directly; arguments
 * travel via OLK_ACTION / OLK_ARGS environment variables and the result comes
 * back as one JSON object on stdout.
 * @module dsh-outlook
 */

import { spawn } from 'node:child_process'
import { createHash, randomBytes } from 'node:crypto'
import { statSync, appendFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { join, dirname } from 'node:path'

export const name = 'dsh-outlook'
export const inject = ['tools']

const HERE = dirname(fileURLToPath(import.meta.url))
const BRIDGE = join(HERE, 'lib', 'olk.ps1')
const TIMEOUT_MS = 120000

/** File-backed lifecycle log for the watcher (console.warn/error are not
 *  captured by the launcher log, which hid earlier failures). Truncates at
 *  1 MB so it cannot grow unbounded. */
const WATCH_LOG = join(HERE, '..', '.olk-watch-debug.log')
const WATCH_LOG_MAX = 1024 * 1024
function wlog(message) {
  try {
    let size = 0
    try { size = statSync(WATCH_LOG).size } catch { /* not created yet */ }
    if (size > WATCH_LOG_MAX) writeFileSync(WATCH_LOG, '')
    appendFileSync(WATCH_LOG, new Date().toISOString() + ' ' + message + '\n')
  } catch { /* diagnostics must never break the plugin */ }
}

/**
 * Serialize every bridge invocation: Outlook COM is single-threaded, and
 * concurrent PowerShell bridges racing the same COM apartment can throw
 * mid-operation. Each olk() call waits for the previous one to settle.
 */
const olkChain = { tail: Promise.resolve() }

/**
 * Run one bridge action and parse its JSON result.
 * Arguments travel over STDIN (not env vars) because Windows caps a single
 * environment variable near 32 KB — a large mail body silently overflowed
 * that cap in env transport, leaving the previous call's stale value behind.
 *
 * HUMAN-APPROVAL GATE (policy, enforced end to end):
 * Nothing that leaves the machine — sending mail, dispatching meeting
 * invites — may run without an explicit in-session human approval first.
 * `opts.confirmed` is set ONLY when the tool's caller passed confirmed:true
 * after the user allowed the two-call confirmation, and it mints a one-time
 * random OLK_CONFIRM token for the child process. The bridge independently
 * refuses every send path unless that token is present, so a future tool
 * (or a bug) that reaches the bridge without a human "allow" cannot send
 * anything.
 */
function olkOne(action, args = {}, opts = {}) {
  return new Promise((resolve) => {
    const env = {
      ...process.env,
      OLK_ACTION: action,
    }
    if (opts.confirmed) env.OLK_CONFIRM = randomBytes(24).toString('hex')
    const child = spawn('powershell', [
      '-NoProfile', '-Command',
      `Invoke-Expression (Get-Content -Raw '${BRIDGE.replace(/'/g, "''")}')`,
    ], {
      env,
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let settled = false
    const timer = setTimeout(() => {
      if (!settled) {
        settled = true
        child.kill()
        resolve({ status: 'error', error: 'bridge timed out after ' + TIMEOUT_MS + 'ms (is Outlook running?)' })
      }
    }, TIMEOUT_MS)
    child.stdin.on('error', () => { /* bridge died early; close handler reports */ })
    child.stdin.write(JSON.stringify(args))
    child.stdin.end()
    child.stdout.on('data', (d) => { stdout += d })
    child.stderr.on('data', (d) => { stderr += d })
    child.on('error', (e) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve({ status: 'error', error: 'failed to start powershell: ' + e.message })
    })
    child.on('close', (code) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      try {
        resolve(JSON.parse(stdout))
      } catch {
        resolve({
          status: 'error',
          error: 'bridge output not JSON: ' + stdout.slice(0, 400),
          exitCode: code,
          stderr: stderr.slice(0, 400),
        })
      }
    })
  })
}

/** Queued bridge entry point (see olkChain above). */
function olk(action, args = {}, opts = {}) {
  const run = () => olkOne(action, args, opts)
  const next = olkChain.tail.then(run, run)
  olkChain.tail = next.catch(() => { /* keep the queue alive after failures */ })
  return next
}

/** Stable-content SHA-256: keys sorted at every level so semantically equal
 *  argument objects hash identically regardless of property order. */
function contentHash(value) {
  const stable = (v) => {
    if (Array.isArray(v)) return v.map(stable)
    if (v !== null && typeof v === 'object') {
      const o = {}
      for (const k of Object.keys(v).sort()) o[k] = stable(v[k])
      return o
    }
    return v
  }
  return createHash('sha256').update(JSON.stringify(stable(value))).digest('hex')
}

/**
 * CONTENT-BINDING for the two-call confirmation (security hardening):
 * remembers the hash of the exact argument content carried by the LAST
 * needs-confirmation result per kind. A confirmed:true retry whose content
 * hash differs from what was last displayed is forced back through
 * needs-confirmation — a caller can never silently swap content between
 * what the user approved and what is dispatched.
 */
const confirmations = new Map()

/** Validate local attachment paths and render them (name + human size) for
 *  the confirmation popup. Returns { error } when a file is missing. */
function attachmentLines(paths) {
  if (!Array.isArray(paths) || paths.length === 0) return { lines: '' }
  const lines = []
  for (const p of paths) {
    let st
    try { st = statSync(p) } catch { return { error: 'attachment not found: ' + p } }
    const kb = Math.max(1, Math.round(st.size / 1024))
    const base = String(p).replace(/^.*[\\/]/, '')
    lines.push(base + ' (' + kb + ' KB)')
  }
  return { lines: lines.join(', ') }
}

/** The `needs-confirmation` first-call result for the two-call protocol.
 *  The agent must relay `detail` to the user verbatim via an in-session
 *  popup and retry with confirmed:true only on an explicit allow. The
 *  popup is EDITABLE: the user may return modified fields; the agent then
 *  retries the call with the edited arguments (which re-enters
 *  needs-confirmation showing the edited content for one final allow), and
 *  only confirmed:true on content the user has seen dispatches. */
function needsConfirmation(kind, detail, editable = null) {
  return {
    status: 'needs-confirmation',
    kind,
    detail,
    editable,
    instruction: 'Show this detail to the user via the in-session popup (ask_user_question) and offer three choices: (a) send as-is, (b) edit — for each editable field give a "keep as-is" option AND let the user type a replacement in the free-text answer, (c) reject. If the user edited any field, retry the call with the edited arguments (without confirmed) so the edited content gets one final confirmation popup; then retry with confirmed:true only after the user allows the content they saw. Never set confirmed:true for content the user has not explicitly seen and allowed.',
  }
}

/** NewMailEx watcher state shared by the badge route and check tool. */
const notify = {
  pending: [],   // new-mail events not yet fetched by the agent
  unread: null,  // last reported Inbox unread count
}

/**
 * Spawn (and baby-sit) the persistent PowerShell watcher (lib/watch.ps1).
 * Emits JSON lines on stdout; crashes restart with a 30s backoff. Killing
 * the child on fiber dispose is the only cleanup the process needs.
 */
function startWatcher(ctx) {
  if (process.env.OLK_WATCH === '0') return
  let child = null
  let stopped = false
  let restartTimer = null
  const WATCH = join(HERE, 'lib', 'watch.ps1')

  function spawnWatcher() {
    if (stopped) return
    child = spawn('powershell', [
      '-NoProfile', '-Command',
      `Invoke-Expression (Get-Content -Raw '${WATCH.replace(/'/g, "''")}')`,
    ], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
    wlog('spawned pid=' + child.pid)
    console.info('[dsh-outlook] watcher spawned (pid ' + child.pid + ')')
    child.on('error', (e) => wlog('spawn error: ' + e.message))
    let buf = ''
    child.stdout.on('data', (d) => {
      wlog('stdout chunk: ' + String(d).trim().slice(0, 200))
      buf += d
      let nl
      while ((nl = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, nl).trim()
        buf = buf.slice(nl + 1)
        if (line === '') continue
        try {
          const ev = JSON.parse(line)
          if (ev.type === 'hello' || ev.type === 'tick') {
            notify.unread = ev.unread
            if (ev.type === 'hello') console.info('[dsh-outlook] watcher online, unread=' + ev.unread)
          } else if (ev.type === 'newmail') {
            notify.pending.push({ entryId: ev.entryId, subject: ev.subject, sender: ev.sender, received: ev.received })
            if (notify.pending.length > 50) notify.pending.shift()
            console.info('[dsh-outlook] new mail: [' + ev.sender + '] ' + ev.subject)
          }
        } catch { /* partial line or non-JSON noise */ }
      }
    })
    child.stderr.on('data', (d) => { const s = String(d).trim(); if (s) { wlog('stderr: ' + s.slice(0, 200)); console.error('[dsh-outlook] watcher stderr: ' + s.slice(0, 200)) } })
    child.on('close', (code, signal) => {
      wlog('closed code=' + code + ' signal=' + signal + ' stopped=' + stopped)
      if (stopped) return
      console.warn('[dsh-outlook] watcher exited (code ' + code + '), restarting in 30s')
      restartTimer = setTimeout(spawnWatcher, 30000)
    })
  }

  spawnWatcher()
  ctx.effect(() => {
    // ctx.effect(fn): fn runs IMMEDIATELY (as setup) and its RETURN VALUE is
    // the disposer — the teardown must be returned, not written as the body
    // (writing it as the body killed the watcher 1ms after spawn).
    return () => {
      wlog('fiber dispose: stopping watcher')
      stopped = true
      if (restartTimer !== null) clearTimeout(restartTimer)
      if (child !== null) { try { child.kill() } catch { /* already gone */ } }
    }
  }, 'dsh-outlook: newmail watcher')
}

/**
 * Register the ten Outlook tools on `ctx.tools`.
 * @param {import('@deepseek-ai/cordis').Context} ctx - registrant context.
 */
export function apply(ctx) {
  if (process.platform !== 'win32') {
    console.info('[dsh-outlook] Windows only — skipping tool registration')
    return
  }

  // NewMailEx push watcher + badge state route (opt out with OLK_WATCH=0).
  startWatcher(ctx)
  ctx.inject(['webServer'], (webCtx) => {
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/api/state',
      handler: async (req, res) => {
        if (req.method !== 'GET') { res.statusCode = 405; res.end(); return }
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.end(JSON.stringify({
          ok: true,
          pendingCount: notify.pending.length,
          pending: notify.pending,
          unread: notify.unread,
        }))
      },
    }), 'dsh-outlook: badge state route')
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/api/client-log',
      handler: async (req, res) => {
        if (req.method !== 'POST') { res.statusCode = 405; res.end(); return }
        let body = ''
        let overflow = false
        for await (const chunk of req) {
          body += chunk
          if (body.length > 65536) { overflow = true; break } // beacon cap: 64 KB
        }
        if (overflow) { res.statusCode = 413; res.end(); return }
        try {
          const data = JSON.parse(body)
          wlog('client: ' + String(data.message || '').slice(0, 300))
        } catch { /* ignore malformed beacons */ }
        res.statusCode = 204
        res.end()
      },
    }), 'dsh-outlook: client beacon route')
  })

  ctx.tools.register({
    name: 'outlook_account',
    description: 'Get the currently configured primary Outlook account name and SMTP address. Read-only.',
    parameters: { type: 'object', additionalProperties: true, properties: {} },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute() { return olk('account') },
    presentCall: () => ({ card: 'generic', title: 'Outlook account', kind: 'other', rawInput: {} }),
  })

  ctx.tools.register({
    name: 'outlook_search_mail',
    description: 'Search emails in the local Outlook mailbox (read-only). Default: last 14 days of Inbox, newest first. Supports arbitrary folder paths ("Inbox/ProjectA", "已发送"), recipient-side matching (to), and exact date ranges (dateFrom/dateTo override days). Returns entryId values usable with outlook_read_mail.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        days: { type: 'number', description: 'How many days back to search. Default 14. Ignored when dateFrom/dateTo is given.' },
        folder: { type: 'string', description: 'Folder to search: "Inbox" (default), "已发送"/"Sent Items", "草稿"/"Drafts", or any path like "Inbox/ProjectA" (segments separated by / or \\). Paths are walked from the primary mailbox root.' },
        from: { type: 'string', description: 'Filter: sender name contains this text.' },
        to: { type: 'string', description: 'Filter: recipient display name contains this text (recipient-side search).' },
        subjectContains: { type: 'string', description: 'Filter: subject contains this text.' },
        dateFrom: { type: 'string', description: 'Exact range start, "yyyy-MM-dd" or "yyyy-MM-dd HH:mm". Overrides days.' },
        dateTo: { type: 'string', description: 'Exact range end, "yyyy-MM-dd" (extends to 23:59) or "yyyy-MM-dd HH:mm". Overrides days.' },
        limit: { type: 'number', description: 'Max results. Default 10.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('search_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Search Outlook mail', kind: 'other', rawInput: a }),
  })

  ctx.tools.register({
    name: 'outlook_read_mail',
    description: 'Read one full email by its entryId (obtained from outlook_search_mail). Returns subject, sender, recipients, body text and attachment names. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId from outlook_search_mail.' },
        maxChars: { type: 'number', description: 'Max body characters. Default 8000.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('read_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Read Outlook mail', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'outlook_draft_mail',
    description: 'Create an email DRAFT in the Outlook Drafts folder. Nothing is sent; the user can review and send it manually in Outlook, or ask for outlook_send_mail which confirms before sending.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        to: { type: 'string', description: 'To recipients, semicolon-separated.' },
        cc: { type: 'string', description: 'Cc recipients, semicolon-separated.' },
        subject: { type: 'string', description: 'Email subject.' },
        body: { type: 'string', description: 'Plain-text body.' },
        attachments: { type: 'array', items: { type: 'string' }, description: 'Optional local file paths to attach.' },
      },
      required: ['to', 'subject', 'body'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('draft_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Draft Outlook mail', kind: 'other', rawInput: { to: a.to, subject: a.subject } }),
  })

  ctx.tools.register({
    name: 'outlook_send_mail',
    description: 'Send an email through the local Outlook client. TWO-CALL CONFIRMATION with EDITABLE popup: without confirmed:true the call returns needs-confirmation with the exact content and the editable field list; show it to the user in-session — they may allow as-is, edit any field (to/cc/subject/body) via free-text answers, or reject. If edited, retry with the edited arguments (one more confirmation of the edited content), then confirmed:true only after the user allows. Nothing is sent on the first call.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        to: { type: 'string', description: 'To recipients, semicolon-separated.' },
        cc: { type: 'string', description: 'Cc recipients, semicolon-separated.' },
        subject: { type: 'string', description: 'Email subject.' },
        body: { type: 'string', description: 'Plain-text body.' },
        attachments: { type: 'array', items: { type: 'string' }, description: 'Optional local file paths to attach; shown in the confirmation popup.' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved this exact content via an in-session popup.' },
      },
      required: ['to', 'subject', 'body'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      // Content binding: the confirmed retry must carry EXACTLY the content
      // of the last needs-confirmation the user saw (see `confirmations`).
      const content = {
        to: a.to, cc: a.cc || '', subject: a.subject, body: String(a.body),
        attachments: Array.isArray(a.attachments) ? a.attachments : [],
      }
      const h = contentHash(content)
      if (a.confirmed !== true || confirmations.get('send-mail') !== h) {
        // Full content, verbatim: the user must see exactly what would be sent.
        const att = attachmentLines(a.attachments)
        if (att.error) return { status: 'error', error: att.error }
        confirmations.set('send-mail', h)
        return needsConfirmation('send-mail',
          '收件人: ' + a.to + (a.cc ? '\n抄送: ' + a.cc : '') + '\n主题: ' + a.subject
          + (att.lines ? '\n附件: ' + att.lines : '')
          + '\n\n正文:\n' + String(a.body),
          ['to', 'cc', 'subject', 'body'])
      }
      const draft = await olk('draft_mail', a)
      if (draft.status !== 'ok') return draft
      return olk('send_draft', { entryId: draft.draftEntryId }, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Send Outlook mail (asks first)', kind: 'other', rawInput: { to: a.to, subject: a.subject } }),
  })

  ctx.tools.register({
    name: 'outlook_calendar_query',
    description: 'List calendar events for the next N days (default 7) from the local Outlook client. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        days: { type: 'number', description: 'How many days ahead to query. Default 7.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('calendar_query', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Query Outlook calendar', kind: 'other', rawInput: a }),
  })

  ctx.tools.register({
    name: 'outlook_calendar_create',
    description: 'Create a calendar entry in Outlook. Personal entries (no recipients) are created directly. With attendees it becomes a meeting invite and uses TWO-CALL CONFIRMATION with EDITABLE popup: the first call returns needs-confirmation; the user may allow as-is, edit fields (subject/location/body), or reject. Edited content gets one more confirmation before confirmed:true dispatches the invite.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        subject: { type: 'string', description: 'Event subject.' },
        start: { type: 'string', description: 'Start time, "yyyy-MM-dd HH:mm" (24h).' },
        end: { type: 'string', description: 'End time, "yyyy-MM-dd HH:mm" (24h).' },
        location: { type: 'string', description: 'Location text.' },
        body: { type: 'string', description: 'Plain-text notes/body.' },
        recipients: { type: 'array', items: { type: 'string' }, description: 'Optional attendee SMTP addresses; adding any turns this into a meeting invite (requires user confirmation).' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved the invite via an in-session popup. Ignored for personal entries.' },
      },
      required: ['subject', 'start', 'end'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const hasRecipients = Array.isArray(a.recipients) && a.recipients.length > 0
      const content = {
        subject: a.subject, start: a.start, end: a.end,
        location: a.location || '', body: String(a.body || ''),
        recipients: Array.isArray(a.recipients) ? a.recipients : [],
      }
      const h = contentHash(content)
      if ((hasRecipients && a.confirmed !== true) || (hasRecipients && confirmations.get('calendar-invite') !== h)) {
        // Full content, verbatim: the user must see exactly what would be sent.
        confirmations.set('calendar-invite', h)
        return needsConfirmation('calendar-invite',
          '主题: ' + a.subject + '\n时间: ' + a.start + ' ~ ' + a.end
            + (a.location ? '\n地点: ' + a.location : '')
            + '\n参会人: ' + a.recipients.join('; ') + '（将发出会议邀请）'
            + (a.body ? '\n\n正文:\n' + String(a.body) : ''),
          ['subject', 'location', 'body'])
      }
      return olk('calendar_create', a, { confirmed: hasRecipients })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Create Outlook calendar entry (asks first)', kind: 'other', rawInput: { subject: a.subject, start: a.start, end: a.end } }),
  })

  ctx.tools.register({
    name: 'outlook_search_people',
    description: 'Search the Outlook address book (GAL) by display name, SMTP address or alias. Returns matching people with name and smtp. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        query: { type: 'string', description: 'Search text: partial display name (e.g. 刘文明), email, or alias.' },
        limit: { type: 'number', description: 'Max results. Default 20.' },
      },
      required: ['query'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('search_people', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Search Outlook address book', kind: 'other', rawInput: { query: a.query } }),
  })

  ctx.tools.register({
    name: 'outlook_freebusy',
    description: 'Query a colleague busy/free status from Outlook. Returns busy slots (status: tentative/busy/out-of-office) per day for the next N days starting today. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        email: { type: 'string', description: 'SMTP address of the person (from outlook_search_people).' },
        days: { type: 'number', description: 'How many days starting today. Default 3.' },
      },
      required: ['email'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('freebusy', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Query Outlook free/busy', kind: 'other', rawInput: { email: a.email, days: a.days } }),
  })

  ctx.tools.register({
    name: 'outlook_search_rooms',
    description: 'Search Outlook meeting rooms by name and optionally check their availability for a time window (date + start/end time). Returns rooms with name, smtp, availability and conflicting slots. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        query: { type: 'string', description: 'Optional name filter, e.g. "Ningbo" or "09.3".' },
        date: { type: 'string', description: 'Check availability on this date, "yyyy-MM-dd". Skip to list rooms only.' },
        start: { type: 'string', description: 'Window start "HH:mm". Default 09:00. Used with date.' },
        end: { type: 'string', description: 'Window end "HH:mm". Default 18:00. Used with date.' },
        limit: { type: 'number', description: 'Max rooms. Default 15.' },
      },
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('search_rooms', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Search Outlook meeting rooms', kind: 'other', rawInput: { query: a.query, date: a.date } }),
  })

  ctx.tools.register({
    name: 'outlook_reply_mail',
    description: 'Reply to, reply-all to, or forward an existing email (entryId from outlook_search_mail). Auto-quotes the original message below your new text; forward carries the original attachments and needs a "to" recipient. TWO-CALL CONFIRMATION with EDITABLE popup: the first call returns needs-confirmation showing the original mail plus your reply; the user may allow as-is, edit (body/to), or reject. Only confirmed:true dispatches.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the mail to reply to / forward (from outlook_search_mail).' },
        mode: { type: 'string', description: '"reply" (sender only), "reply-all" (sender + all recipients), or "forward" (new recipient, keeps attachments).' },
        body: { type: 'string', description: 'New text to write above the auto-quoted original.' },
        to: { type: 'string', description: 'Forward recipient(s), semicolon-separated. Required for forward; ignored for reply/reply-all.' },
        attachments: { type: 'array', items: { type: 'string' }, description: 'Optional extra local file paths to attach (forward already carries the original attachments).' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved this exact reply via an in-session popup.' },
      },
      required: ['entryId', 'mode'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const content = {
        entryId: a.entryId, mode: a.mode, body: String(a.body || ''),
        to: a.to || '', attachments: Array.isArray(a.attachments) ? a.attachments : [],
      }
      const h = contentHash(content)
      if (a.confirmed !== true || confirmations.get('reply-mail') !== h) {
        // Fetch the original mail so the confirmation shows exactly which
        // mail the reply goes to — fail closed if it cannot be read.
        const prev = await olk('reply_preview', { entryId: a.entryId })
        if (prev.status !== 'ok') return prev
        const att = attachmentLines(a.attachments)
        if (att.error) return { status: 'error', error: att.error }
        const modeText = a.mode === 'forward' ? '转发给: ' + (a.to || '(未填)')
          : a.mode === 'reply-all' ? '全部回复' : '回复'
        confirmations.set('reply-mail', h)
        return needsConfirmation('reply-mail',
          modeText + '\n原邮件: [' + prev.sender + '] ' + prev.subject
          + (a.to && a.mode === 'forward' ? '\n转发收件人: ' + a.to : '')
          + (att.lines ? '\n附加附件: ' + att.lines : '')
          + (a.mode === 'forward' ? '\n（转发自动携带原邮件附件）' : '')
          + '\n\n你的内容:\n' + String(a.body || '')
          + '\n\n--- 以下自动引用原邮件 ---\nFrom: ' + prev.sender
          + '\nSent: ' + (prev.received || '')
          + '\nTo: ' + prev.to
          + '\nSubject: ' + prev.subject,
          a.mode === 'forward' ? ['to', 'body'] : ['body'])
      }
      return olk('reply_mail', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Reply/forward Outlook mail (asks user first)', kind: 'other', rawInput: { entryId: a.entryId, mode: a.mode, to: a.to } }),
  })

  ctx.tools.register({
    name: 'outlook_save_attachment',
    description: 'Save one attachment of an email to a local directory, by mail entryId and attachment display name (names come from outlook_read_mail). Returns the saved file path and size. Writes a local file only; nothing leaves the machine.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the mail holding the attachment (from outlook_search_mail).' },
        name: { type: 'string', description: 'Attachment display name exactly as listed by outlook_read_mail.' },
        dir: { type: 'string', description: 'Target directory for the saved file (must exist).' },
      },
      required: ['entryId', 'name', 'dir'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('save_attachment', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Save Outlook attachment', kind: 'other', rawInput: { entryId: a.entryId, name: a.name } }),
  })

  ctx.tools.register({
    name: 'outlook_check_new',
    description: 'Fetch new-mail arrivals pushed by the Outlook NewMailEx watcher since the last check, as a list (entryId, subject, sender, time). Fetching clears the pending queue and resets the sidebar badge. Returns unread totals too. Read-only.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {},
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute() {
      const pending = notify.pending.splice(0, notify.pending.length)
      return { status: 'ok', count: pending.length, newMails: pending, unread: notify.unread }
    },
    presentCall: () => ({ card: 'generic', title: 'Check new Outlook mail', kind: 'other', rawInput: {} }),
  })

  console.info('[dsh-outlook] registered 13 Outlook COM tools')
}
