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
 * because group policy commonly blocks running .ps1 files directly; the
 * action name travels in the OLK_ACTION environment variable, arguments
 * travel as ONE JSON object on STDIN (env vars cap near 32 KB), and the
 * result comes back as one JSON object on stdout.
 * @module dsh-outlook
 */

import { spawn } from 'node:child_process'
import { createHash, randomBytes } from 'node:crypto'
import { statSync, appendFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { fileURLToPath } from 'node:url'
import { join, dirname } from 'node:path'

export const name = 'dsh-outlook'
export const inject = ['tools']

const HERE = dirname(fileURLToPath(import.meta.url))
const BRIDGE = join(HERE, 'lib', 'olk.ps1')
const TIMEOUT_MS = 120000

/** File-backed lifecycle log for the watcher (console.warn/error are not
 *  captured by the launcher log, which hid earlier failures). Lives in the
 *  OS temp dir, NOT the plugin install dir — a global/DSH install location
 *  can be read-only (then the log silently no-ops, which is fine for
 *  diagnostics). Truncates at 1 MB so it cannot grow unbounded. */
const WATCH_LOG = join(tmpdir(), 'dsh-outlook-watch-debug.log')
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
      `Invoke-Expression (Get-Content -Raw -Encoding UTF8 '${BRIDGE.replace(/'/g, "''")}')`,
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
 * what the user approved and what is dispatched. A successful dispatch
 * consumes the hash (single-use), so the same approval cannot be replayed
 * into a second send.
 */
const confirmations = new Map()

/** Run the confirmed dispatch and consume the approval on success, so a
 *  replayed confirmed:true call must pass through a fresh popup. */
async function dispatchConfirmed(kind, ...olkArgs) {
  const result = await olk(...olkArgs)
  if (result && result.status === 'ok') confirmations.delete(kind)
  return result
}

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
      `Invoke-Expression (Get-Content -Raw -Encoding UTF8 '${WATCH.replace(/'/g, "''")}')`,
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
          } else if (ev.type === 'no-outlook') {
            // Classic Outlook not installed on this machine: the watcher
            // exits with code 3 right after this; do not restart it.
            console.info('[dsh-outlook] ' + (ev.message || 'classic Outlook not installed') + ' — badge watcher disabled (tools still available if Outlook appears later via a profile restart)')
          }
        } catch { /* partial line or non-JSON noise */ }
      }
    })
    child.stderr.on('data', (d) => { const s = String(d).trim(); if (s) { wlog('stderr: ' + s.slice(0, 200)); console.error('[dsh-outlook] watcher stderr: ' + s.slice(0, 200)) } })
    child.on('close', (code, signal) => {
      wlog('closed code=' + code + ' signal=' + signal + ' stopped=' + stopped)
      if (stopped) return
      if (code === 3) {
        // Watcher's "classic Outlook not installed" exit: no restart loop
        // on machines without desktop Outlook (public-distribution safety).
        console.warn('[dsh-outlook] watcher disabled: classic desktop Outlook not installed')
        return
      }
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
    // Same-origin guard for the two badge routes (M3): mail metadata must
    // not be readable by random pages probing localhost. Non-browser
    // clients (no Origin header) stay allowed — same trust as before.
    const sameOrigin = (req) => {
      const origin = req.headers.origin
      if (origin === undefined) return true
      try {
        return new URL(origin).host === req.headers.host
      } catch {
        return false
      }
    }
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/api/state',
      handler: async (req, res) => {
        if (req.method !== 'GET') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.end(JSON.stringify({
          ok: true,
          pendingCount: notify.pending.length,
          pending: notify.pending,
          unread: notify.unread,
        }))
      },
    }), 'dsh-outlook: badge state route')
    // Click-to-open route: GET serves the chat-surface markdown link
    // (/olk-notify/open?entryId=... — browser navigation carries no Origin
    // header, same trust as above); POST serves the sidebar popover. Both
    // open the mail in a desktop Outlook inspector via the open_mail
    // bridge action (local window, nothing leaves the machine).
    const ENTRY_ID_RE = /^[A-Za-z0-9+/=_-]{20,512}$/
    const readEntryId = (q) => {
      const id = String(q || '')
      return ENTRY_ID_RE.test(id) ? id : null
    }
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
    const openPage = (res, ok, message, subject) => {
      res.statusCode = ok ? 200 : 500
      res.setHeader('Content-Type', 'text/html; charset=utf-8')
      res.end('<!doctype html><meta charset="utf-8"><title>Outlook</title>' +
        '<body style="font-family:system-ui;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;color:#222">' +
        '<div style="text-align:center"><div style="font-size:40px">' + (ok ? '📨' : '⚠️') + '</div>' +
        '<p>' + esc(message) + (subject ? '<br><b>' + esc(subject) + '</b>' : '') + '</p>' +
        (ok ? '<p style="opacity:.6">此页面可关闭</p>' : '') + '</div></body>')
    }
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/open',
      handler: async (req, res) => {
        if (req.method !== 'GET' && req.method !== 'POST') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
        let entryId = null
        let json = false
        if (req.method === 'GET') {
          const u = new URL(req.url, 'http://localhost')
          entryId = readEntryId(u.searchParams.get('entryId'))
        } else {
          json = true
          let body = ''
          let overflow = false
          for await (const chunk of req) {
            body += chunk
            if (body.length > 8192) { overflow = true; break }
          }
          if (overflow) { res.statusCode = 413; res.end(); return }
          try { entryId = readEntryId(JSON.parse(body).entryId) } catch { entryId = null }
        }
        if (entryId === null) {
          if (json) { res.statusCode = 400; res.setHeader('Content-Type', 'application/json; charset=utf-8'); res.end(JSON.stringify({ ok: false, error: 'invalid entryId' })) }
          else openPage(res, false, 'entryId 无效')
          return
        }
        const r = await olk('open_mail', { entryId })
        if (r.status === 'ok') {
          // Reading by opening: drop the mail from the pending queue (badge
          // count = pending length, so the sidebar chip decrements on the
          // next poll) and keep the unread total in step when the mail was
          // previously unread.
          const i = notify.pending.findIndex((m) => m.entryId === entryId)
          if (i >= 0) notify.pending.splice(i, 1)
          if (r.wasUnread && typeof notify.unread === 'number' && notify.unread > 0) notify.unread--
        }
        if (json) {
          res.setHeader('Content-Type', 'application/json; charset=utf-8')
          res.statusCode = r.status === 'ok' ? 200 : 500
          res.end(JSON.stringify({ ok: r.status === 'ok', subject: r.subject, sender: r.sender, error: r.error }))
        } else if (r.status === 'ok') openPage(res, true, '已在 Outlook 中打开', r.subject)
        else openPage(res, false, '打开失败：' + (r.error || 'unknown'))
      },
    }), 'dsh-outlook: open-mail route')
    ctx.effect(() => webCtx.webServer.register({
      kind: 'exact',
      path: '/olk-notify/api/client-log',
      handler: async (req, res) => {
        if (req.method !== 'POST') { res.statusCode = 405; res.end(); return }
        if (!sameOrigin(req)) { res.statusCode = 403; res.end(); return }
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
    name: 'outlook_open_mail',
    description: 'Open one email in a desktop Outlook inspector window by entryId (from outlook_search_mail / outlook_check_new). Local action: nothing leaves the machine, no confirmation needed. When presenting search results in chat, prefer composing a clickable markdown link per mail, e.g. [在Outlook打开](http://127.0.0.1:3080/olk-notify/open?entryId=<entryId>) — the user can click it to open the mail directly.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId from outlook_search_mail or outlook_check_new.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) { return olk('open_mail', a) },
    presentCall: (a) => ({ card: 'generic', title: 'Open Outlook mail', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'outlook_search_mail',
    description: 'Search emails in the local Outlook mailbox (read-only). Default: last 14 days of Inbox, newest first. Supports arbitrary folder paths ("Inbox/ProjectA", "已发送"), recipient-side matching (to), and exact date ranges (dateFrom/dateTo override days). Returns entryId values usable with outlook_read_mail. Tip: when listing results to the user, append a clickable open link per mail: [打开](http://127.0.0.1:3080/olk-notify/open?entryId=<entryId>).',
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
        bcc: { type: 'string', description: 'Bcc (blind carbon copy) recipients, semicolon-separated.' },
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
        bcc: { type: 'string', description: 'Bcc (blind carbon copy) recipients, semicolon-separated; hidden from other recipients.' },
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
        to: a.to, cc: a.cc || '', bcc: a.bcc || '', subject: a.subject, body: String(a.body),
        attachments: Array.isArray(a.attachments) ? a.attachments : [],
      }
      const h = contentHash(content)
      if (a.confirmed !== true || confirmations.get('send-mail') !== h) {
        // Full content, verbatim: the user must see exactly what would be sent.
        const att = attachmentLines(a.attachments)
        if (att.error) return { status: 'error', error: att.error }
        confirmations.set('send-mail', h)
        return needsConfirmation('send-mail',
          '收件人: ' + a.to + (a.cc ? '\n抄送: ' + a.cc : '') + (a.bcc ? '\n密送: ' + a.bcc : '') + '\n主题: ' + a.subject
          + (att.lines ? '\n附件: ' + att.lines : '')
          + '\n\n正文:\n' + String(a.body),
          ['to', 'cc', 'bcc', 'subject', 'body'])
      }
      const draft = await olk('draft_mail', a)
      if (draft.status !== 'ok') return draft
      return dispatchConfirmed('send-mail', 'send_draft', { entryId: draft.draftEntryId }, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Send Outlook mail (asks first)', kind: 'other', rawInput: { to: a.to, subject: a.subject } }),
  })

  ctx.tools.register({
    name: 'outlook_calendar_query',
    description: 'List calendar events for the next N days (default 7) from the local Outlook client. Each event carries an entryId usable with outlook_meeting_update / outlook_meeting_cancel, plus an isMeeting flag. Read-only.',
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
    description: 'Create a calendar entry in Outlook. Personal entries (no attendees) are created directly. With attendees (required recipients or optionalRecipients) it becomes a meeting invite and uses TWO-CALL CONFIRMATION with EDITABLE popup: the first call returns needs-confirmation; the user may allow as-is, edit fields (subject/location/body), or reject. Supports recurring meetings via the recurrence object (daily/weekly/monthly).',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        subject: { type: 'string', description: 'Event subject.' },
        start: { type: 'string', description: 'Start time, "yyyy-MM-dd HH:mm" (24h).' },
        end: { type: 'string', description: 'End time, "yyyy-MM-dd HH:mm" (24h).' },
        location: { type: 'string', description: 'Location text.' },
        body: { type: 'string', description: 'Plain-text notes/body.' },
        recipients: { type: 'array', items: { type: 'string' }, description: 'Required attendee SMTP addresses; adding any (or optionalRecipients) turns this into a meeting invite (requires user confirmation).' },
        optionalRecipients: { type: 'array', items: { type: 'string' }, description: 'Optional attendees (marked optional in the invite).' },
        recurrence: { type: 'object', additionalProperties: true, description: 'Optional recurrence: { type: "daily"|"weekly"|"monthly", interval: n, daysOfWeek: ["Mon","Wed"...] (weekly), count: n | until: "yyyy-MM-dd" }. Defaults to 10 occurrences if neither count nor until is given.' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved the invite via an in-session popup. Ignored for personal entries.' },
      },
      required: ['subject', 'start', 'end'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const required = Array.isArray(a.recipients) ? a.recipients : []
      const optional = Array.isArray(a.optionalRecipients) ? a.optionalRecipients : []
      const hasRecipients = required.length > 0 || optional.length > 0
      const rec = a.recurrence || null
      const content = {
        subject: a.subject, start: a.start, end: a.end,
        location: a.location || '', body: String(a.body || ''),
        recipients: required, optionalRecipients: optional,
        recurrence: rec ? JSON.stringify(rec) : '',
      }
      const h = contentHash(content)
      if ((hasRecipients && a.confirmed !== true) || (hasRecipients && confirmations.get('calendar-invite') !== h)) {
        // Full content, verbatim: the user must see exactly what would be sent.
        confirmations.set('calendar-invite', h)
        let recText = ''
        if (rec) {
          recText = '\n周期: ' + rec.type + (rec.interval ? ' x' + rec.interval : '')
          if (rec.daysOfWeek && rec.daysOfWeek.length) recText += ' on ' + rec.daysOfWeek.join(',')
          if (rec.count) recText += ', ' + rec.count + '次'
          else if (rec.until) recText += ', until ' + rec.until
        }
        return needsConfirmation('calendar-invite',
          '主题: ' + a.subject + '\n时间: ' + a.start + ' ~ ' + a.end
            + (a.location ? '\n地点: ' + a.location : '')
            + (required.length ? '\n必选参会人: ' + required.join('; ') : '')
            + (optional.length ? '\n可选参会人: ' + optional.join('; ') : '')
            + (hasRecipients ? '（将发出会议邀请）' : '')
            + recText
            + (a.body ? '\n\n正文:\n' + String(a.body) : ''),
          ['subject', 'location', 'body'])
      }
      return dispatchConfirmed('calendar-invite', 'calendar_create', a, { confirmed: hasRecipients })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Create Outlook calendar entry (asks first)', kind: 'other', rawInput: { subject: a.subject, start: a.start, end: a.end } }),
  })

  ctx.tools.register({
    name: 'outlook_meeting_requests',
    description: 'List meeting requests waiting for your response (invites others sent you, last 30 days), with organizer, proposed time and location. Read-only; respond with outlook_respond_meeting.',
    parameters: { type: 'object', additionalProperties: true, properties: {} },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute() { return olk('meeting_requests') },
    presentCall: () => ({ card: 'generic', title: 'List Outlook meeting requests', kind: 'other', rawInput: {} }),
  })

  ctx.tools.register({
    name: 'outlook_respond_meeting',
    description: 'Accept, tentatively accept, or decline a meeting request (entryId from outlook_meeting_requests). TWO-CALL CONFIRMATION: the first call returns needs-confirmation; retry with confirmed:true after the user allows. Accepting/declining sends a response to the organizer.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the meeting request (from outlook_meeting_requests).' },
        response: { type: 'string', description: '"accept", "tentative", or "decline".' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved via an in-session popup.' },
      },
      required: ['entryId', 'response'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const h = contentHash({ entryId: a.entryId, response: a.response })
      if (a.confirmed !== true || confirmations.get('meeting-respond') !== h) {
        confirmations.set('meeting-respond', h)
        const label = a.response === 'accept' ? '接受' : a.response === 'tentative' ? '暂定接受' : a.response === 'decline' ? '拒绝' : String(a.response)
        return needsConfirmation('meeting-respond',
          '会议响应: ' + label + '\n请求 entryId: ' + a.entryId + '\n（响应将发送给组织者；先用 outlook_meeting_requests 查看详情）')
      }
      return dispatchConfirmed('meeting-respond', 'respond_meeting', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Respond to Outlook meeting (asks first)', kind: 'other', rawInput: { entryId: a.entryId, response: a.response } }),
  })

  ctx.tools.register({
    name: 'outlook_meeting_update',
    description: 'Update an existing appointment/meeting on the calendar (entryId from outlook_calendar_query — use the outlook entryId if available). Changed fields are applied; if the meeting has attendees, Outlook re-notifies them. TWO-CALL CONFIRMATION: first call returns needs-confirmation; retry with confirmed:true after the user allows. Applies to personal appointments too (uniform gate).',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the appointment/meeting to update.' },
        subject: { type: 'string', description: 'New subject (optional).' },
        start: { type: 'string', description: 'New start, "yyyy-MM-dd HH:mm" (optional).' },
        end: { type: 'string', description: 'New end, "yyyy-MM-dd HH:mm" (optional).' },
        location: { type: 'string', description: 'New location (optional).' },
        body: { type: 'string', description: 'New body (optional).' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved; required only when attendees will be notified.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const h = contentHash({ entryId: a.entryId, subject: a.subject || '', start: a.start || '', end: a.end || '', location: a.location || '', body: a.body || '' })
      if (a.confirmed !== true || confirmations.get('meeting-update') !== h) {
        confirmations.set('meeting-update', h)
        const changes = []
        if (a.subject) changes.push('主题→ ' + a.subject)
        if (a.start) changes.push('开始→ ' + a.start)
        if (a.end) changes.push('结束→ ' + a.end)
        if (a.location) changes.push('地点→ ' + a.location)
        if (a.body) changes.push('正文→ ' + String(a.body).slice(0, 300))
        if (changes.length === 0) return { status: 'error', error: 'nothing to update: provide at least one changed field.' }
        return needsConfirmation('meeting-update',
          '更新会议 entryId: ' + a.entryId + '\n变更:\n- ' + changes.join('\n- ')
          + '\n（若为多人会议，将自动通知所有参会人）')
      }
      return dispatchConfirmed('meeting-update', 'meeting_update', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Update Outlook meeting (asks first)', kind: 'other', rawInput: { entryId: a.entryId } }),
  })

  ctx.tools.register({
    name: 'outlook_meeting_cancel',
    description: 'Cancel a meeting (notifies all attendees with a cancellation) or delete a personal appointment. TWO-CALL CONFIRMATION for both paths: the first call returns needs-confirmation; retry with confirmed:true after the user allows.',
    parameters: {
      type: 'object',
      additionalProperties: true,
      properties: {
        entryId: { type: 'string', description: 'entryId of the appointment/meeting to cancel.' },
        confirmed: { type: 'boolean', description: 'Set true ONLY after the user approved via an in-session popup.' },
      },
      required: ['entryId'],
    },
    output: {
      schema: { type: 'object', additionalProperties: true },
      render: (_a, v) => [{ type: 'text', text: JSON.stringify(v) }],
    },
    async execute(a) {
      const h = contentHash({ entryId: a.entryId })
      if (a.confirmed !== true || confirmations.get('meeting-cancel') !== h) {
        confirmations.set('meeting-cancel', h)
        return needsConfirmation('meeting-cancel',
          '取消会议 entryId: ' + a.entryId + '\n（若为多人会议，将向所有参会人发出取消通知）')
      }
      return dispatchConfirmed('meeting-cancel', 'meeting_cancel', a, { confirmed: true })
    },
    presentCall: (a) => ({ card: 'generic', title: 'Cancel Outlook meeting (asks first)', kind: 'other', rawInput: { entryId: a.entryId } }),
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
        cc: { type: 'string', description: 'Optional Cc recipients, semicolon-separated. Overwrites the reply-all preset when given.' },
        bcc: { type: 'string', description: 'Optional Bcc recipients, semicolon-separated; hidden from other recipients.' },
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
        to: a.to || '', cc: a.cc || '', bcc: a.bcc || '',
        attachments: Array.isArray(a.attachments) ? a.attachments : [],
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
          + (a.cc ? '\n抄送: ' + a.cc : '')
          + (a.bcc ? '\n密送: ' + a.bcc : '')
          + (att.lines ? '\n附加附件: ' + att.lines : '')
          + (a.mode === 'forward' ? '\n（转发自动携带原邮件附件）' : '')
          + '\n\n你的内容:\n' + String(a.body || '')
          + '\n\n--- 以下自动引用原邮件 ---\nFrom: ' + prev.sender
          + '\nSent: ' + (prev.received || '')
          + '\nTo: ' + prev.to
          + '\nSubject: ' + prev.subject,
          a.mode === 'forward' ? ['to', 'cc', 'bcc', 'body'] : ['cc', 'bcc', 'body'])
      }
      return dispatchConfirmed('reply-mail', 'reply_mail', a, { confirmed: true })
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
    description: 'Fetch new-mail arrivals pushed by the Outlook NewMailEx watcher since the last check, as a list (entryId, subject, sender, time). Fetching clears the pending queue and resets the sidebar badge. Returns unread totals too. Read-only. Tip: when presenting arrivals to the user, append a clickable open link per mail: [打开](http://127.0.0.1:3080/olk-notify/open?entryId=<entryId>).',
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

  console.info('[dsh-outlook] registered 18 Outlook COM tools')
}
