/* dsh-outlook client face v2.6.1 (hand-written, zero build step).
 *
 * v2.6.1 — rail-mode badge: the count chip sits INSIDE the 36px circle's
 * top right (14px dot) in the collapsed sidebar, so it can never be
 * clipped by the collapse animation or overlapped by the row above;
 * chip gains z-index:1 everywhere.
 *
 * v2.6.0 — slot-based rewrite. The sidebar button no longer injects raw
 * DOM into the shell's footArea (first child, MutationObserver
 * self-healing, hashed-classname sniffing for wide/collapsed). It now
 * registers through the official `sidebar.footer.action` slot exactly
 * like dsh-dock's Whale Bay entry:
 *
 *   - React owns the row lifecycle end to end (no observers, no
 *     relocation, no freeze risk — the v2.5.1 microtask-storm class of
 *     bugs is structurally impossible now);
 *   - wide/rail geometry comes from the slot's `wide` prop instead of
 *     `[class*="collapsed"]` matching;
 *   - the arrivals popover is a React portal (react-dom is a
 *     kernel-provided module), so its lifecycle is React-owned too;
 *   - order:10 keeps the row above Whale Bay (order:90) — the same
 *     visual position the footArea-first-child hack used to buy.
 *
 * Hover/base colors follow the native settings trigger tokens
 * (--dsw-alias-interactive-bg-hover), as aligned in v2.5.2/v2.5.3.
 *
 * The host half (state route /olk-notify/api/state, watcher, tools) is
 * unchanged; this file only swaps the presentation substrate.
 *
 * Loader format matches the modules node half: a __ModuleLoader__ bundle
 * whose factory receives `require` for kernel-provided modules.
 */
window.__ModuleLoader__.load({
  id: 'dsh-outlook',
  factory: (require) => {
    const React = require('react')
    const ReactDOM = require('react-dom')

    const STATE_URL = '/olk-notify/api/state'
    const POLL_MS = 20000

    /** Lifecycle beacon → host debug file (diagnostics; silent on failure). */
    const beacon = (message) => {
      try {
        fetch('/olk-notify/api/client-log', {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ message }),
        }).catch(() => {})
      } catch { /* never break the UI over a beacon */ }
    }
    beacon('factory loaded (v2.11.0 chat-link interception + slot)')

    /**
     * Styles for the row. Includes the same slot-anchor stacking rule
     * dsh-dock installs ([data-slot="sidebar.footer.action"] laid out as a
     * vertical column) so full-row entries coexist even if dsh-dock is
     * absent or changes its mind; duplicated rule, id-guarded tag.
     */
    function ensureStyles() {
      if (document.head === null) return null
      if (document.getElementById('dsh-outlook-styles') !== null) return null
      const style = document.createElement('style')
      style.id = 'dsh-outlook-styles'
      style.textContent = [
        '[data-slot="sidebar.footer.action"]{',
        '  display: flex !important; /* override the inline display:contents */',
        '  flex-direction: column;',
        '  width: 100%;',
        '  align-items: stretch;',
        '  min-width: 0;',
        '}',
        // Trigger chrome mirrors the native settings trigger. NOTE flex:none —
        // the settings trigger's flex:1 is HORIZONTAL (row context); this slot
        // anchor is stacked as a vertical column (dsh-dock layout), where
        // flex:1 would stretch the button's HEIGHT. Width is set explicitly
        // instead (calc(100% + 4px), same as the settings triggerRow).
        '.olk-trigger{box-sizing:border-box;cursor:pointer;flex:none;width:calc(100% + 4px);min-width:0;height:42px;color:var(--dsw-alias-label-primary);background:0 0;border:none;border-radius:12px;align-items:center;gap:8px;margin:4px -2px;padding:0 10px 0 8px;font-family:inherit;font-size:14px;line-height:22px;display:flex;transition:background .12s ease;position:relative}',
        '.olk-trigger:hover,.olk-trigger:focus-visible{background:var(--dsw-alias-interactive-bg-hover);outline:none}',
        '.olk-trigger.olk-rail{border-radius:50%;flex:none;justify-content:center;gap:0;width:36px;height:36px;margin:8px 0 10px;padding:0}',
        '.olk-label{white-space:nowrap;overflow:hidden;min-width:0;flex:1 1 auto;text-align:left}',
        '.olk-chip{position:absolute;top:-4px;right:-4px;min-width:16px;height:16px;padding:0 4px;border-radius:8px;font-size:10px;color:#fff;background:var(--dsw-alias-label-danger,#e5484d);box-shadow:0 0 0 2px var(--dsw-specific-sidebar-fill,#fff);display:flex;align-items:center;justify-content:center;z-index:1}',
        '.olk-chip[hidden]{display:none}',
        // Rail (collapsed sidebar): badge sits INSIDE the 36px circle's top
        // right — cannot be clipped by the column's collapse animation or
        // overlapped by the row above; smaller dot geometry.
        '.olk-trigger.olk-rail .olk-chip{top:4px;right:4px;min-width:14px;height:14px;padding:0 3px;border-radius:7px;font-size:9px;box-shadow:none}',
        '.olk-popover{position:fixed;z-index:60;min-width:260px;max-width:360px;padding:10px 12px;border-radius:10px;background-color:var(--dsh-bg-elevated,#fff);color:inherit;box-shadow:0 8px 24px rgba(0,0,0,.18);font-size:12.5px;line-height:1.6}',
        '.olk-popover-title{font-weight:600;margin-bottom:6px}',
        '.olk-popover-hint{opacity:.65;margin-top:6px}',
        '.olk-mailrow{display:block;width:100%;text-align:left;background:0 0;border:none;border-radius:6px;padding:4px 6px;margin:0 -6px;font:inherit;color:inherit;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;transition:background .12s ease}',
        '.olk-mailrow:hover:not(:disabled),.olk-mailrow:focus-visible:not(:disabled){background:var(--dsw-alias-interactive-bg-hover);outline:none}',
        '.olk-mailrow:disabled{cursor:default;opacity:.55}',
        // Toast shown when a chat markdown link to /olk-notify/open is
        // intercepted (v2.11.0): the mail opens via same-page POST, so the
        // user needs some in-page feedback instead of a result tab.
        '.olk-toast{position:fixed;z-index:70;left:50%;bottom:28px;transform:translateX(-50%);max-width:min(70vw,480px);padding:8px 14px;border-radius:10px;background:var(--dsh-bg-elevated,#fff);color:var(--dsw-alias-label-primary,#222);box-shadow:0 8px 24px rgba(0,0,0,.18);font-size:12.5px;line-height:1.5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
      ].join('\n')
      document.head.appendChild(style)
      return style
    }

    /** Envelope icon (16 viewBox, Feather-compatible strokes). */
    function MailIcon({ size }) {
      return React.createElement('svg', {
        width: size, height: size, viewBox: '0 0 16 16',
        fill: 'none', stroke: 'currentColor', 'stroke-width': 1.3,
        'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': true,
      },
        React.createElement('rect', { x: 2, y: 3.5, width: 12, height: 9.5, rx: 1.5 }),
        React.createElement('path', { d: 'm2.5 4.5 5.5 4.7 5.5-4.7' }))
    }

    /**
     * The new-mail trigger row. `wide` comes from the slot owner props.
     * Polls the host state route; the count chip appears with arrivals.
     */
    function NewMailRow(props) {
      const wide = props.wide !== false
      const [pending, setPending] = React.useState([])
      const [open, setOpen] = React.useState(false)
      const btnRef = React.useRef(null)

      React.useEffect(() => {
        let stopped = false
        const poll = async () => {
          if (stopped) return
          try {
            const res = await fetch(STATE_URL, { method: 'GET' })
            if (!res.ok) return
            const data = await res.json()
            if (!stopped && data && data.ok) setPending((data && data.pending) || [])
          } catch { /* server restarting; next tick retries */ }
        }
        poll()
        const timer = setInterval(poll, POLL_MS)
        return () => { stopped = true; clearInterval(timer) }
      }, [])

      // Clicking anywhere outside the popover dismisses it.
      React.useEffect(() => {
        if (!open) return undefined
        const dismiss = (ev) => {
          if (btnRef.current !== null && ev.target === btnRef.current) return
          const pop = document.querySelector('.olk-popover')
          if (pop !== null && pop.contains(ev.target)) return
          setOpen(false)
        }
        document.addEventListener('click', dismiss, true)
        return () => document.removeEventListener('click', dismiss, true)
      }, [open])

      const n = pending.length
      const lines = pending.slice(0, 8).map((m) => '[' + (m.sender || '?') + '] ' + (m.subject || '(无主题)'))

      // Click-to-open: POST the entryId to the host open-mail route, which
      // opens the mail in a desktop Outlook inspector AND marks it read.
      // On success the item is removed from the local list immediately —
      // the badge count (pending length) decrements right away, matching
      // the host-side pending-queue removal (next poll would agree).
      const [busy, setBusy] = React.useState(false)
      const openMail = async (entryId) => {
        if (busy || entryId === undefined || entryId === null) return
        setBusy(true)
        try {
          const res = await fetch('/olk-notify/open', {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ entryId: entryId }),
          })
          const data = await res.json().catch(() => null)
          if (data && data.ok) setPending((prev) => prev.filter((m) => m.entryId !== entryId))
          else beacon('open failed: ' + ((data && data.error) || ('HTTP ' + res.status)))
        } catch (e) {
          beacon('open error: ' + (e && e.message))
        } finally { setBusy(false) }
      }

      // Popover: React portal to body, fixed position beside the button.
      let popover = null
      if (open) {
        let left = 60
        let top = 60
        if (btnRef.current !== null) {
          const r = btnRef.current.getBoundingClientRect()
          left = r.right + 8
          top = Math.max(8, r.top - 8)
        }
        popover = ReactDOM.createPortal(
          React.createElement('div', {
            className: 'olk-popover',
            role: 'dialog',
            'aria-label': 'Outlook 新邮件列表',
            style: { left: left + 'px', top: top + 'px' },
            onClick: (ev) => ev.stopPropagation(),
          },
            React.createElement('div', { className: 'olk-popover-title', role: 'heading', 'aria-level': 2 }, '📬 新邮件（' + n + '）'),
            pending.slice(0, 10).map((m, i) => React.createElement('button', {
              key: i,
              type: 'button',
              className: 'olk-mailrow',
              disabled: busy,
              onClick: () => { openMail(m.entryId) },
              title: '点击在 Outlook 中打开并标记已读',
            },
              '[' + (m.sender || '?') + '] ' + (m.subject || '(无主题)') + ' · ' + (m.received || ''))),
            React.createElement('div', { className: 'olk-popover-hint' }, '点击邮件可在 Outlook 中打开；在会话里问「有什么新邮件」可读取并清空')),
          document.body)
      }

      return React.createElement(React.Fragment, null,
        React.createElement('button', {
          ref: btnRef,
          type: 'button',
          className: 'olk-trigger' + (wide ? '' : ' olk-rail'),
          'data-dsh-plugin': 'dsh-outlook',
          'data-dsh-part': 'sidebar-entry',
          'aria-haspopup': 'dialog',
          'aria-label': n > 0 ? 'Outlook 新邮件 ' + n + ' 封' : 'Outlook 邮件',
          title: lines.length > 0 ? lines.join('\n') : 'Outlook 邮件（暂无新邮件）',
          onClick: () => setOpen((v) => !v),
        },
          React.createElement(MailIcon, { size: wide ? 16 : 18 }),
          wide ? React.createElement('span', { className: 'olk-label' }, 'Outlook') : null,
          React.createElement('span', { className: 'olk-chip', hidden: n <= 0 }, String(n))),
        popover)
    }

    /**
     * v2.11.0 — chat-link click interception. A plain left-click on any
     * anchor whose URL path is /olk-notify/open?entryId=… (regardless of
     * the port baked into the markdown link — matched by pathname, posted
     * same-origin) is answered with the very same POST the sidebar popover
     * uses, so the mail opens in Outlook with NO extra browser tab. Modifier
     * clicks / middle clicks keep the navigation fallback (the GET result
     * page). This also repairs the known boundary where a custom-port DSH
     * deployment made the hardcoded :3080 chat links dead: interception
     * rewrites them to the page's own origin.
     */
    const toastTimers = new Set()
    const showToast = (text, isError) => {
      const div = document.createElement('div')
      div.className = 'olk-toast'
      div.setAttribute('role', 'status')
      if (isError) div.style.color = 'var(--dsw-alias-label-danger,#e5484d)'
      div.textContent = text
      document.body.appendChild(div)
      const t = setTimeout(() => {
        toastTimers.delete(t)
        if (div.parentNode !== null) div.parentNode.removeChild(div)
      }, 2600)
      toastTimers.add(t)
    }

    const openMailByEntryId = async (entryId) => {
      try {
        const res = await fetch('/olk-notify/open', {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ entryId: entryId }),
        })
        const data = await res.json().catch(() => null)
        if (data && data.ok) showToast('📨 已在 Outlook 中打开：' + (data.subject || '(无主题)'))
        else {
          showToast('⚠️ 打开失败：' + ((data && data.error) || ('HTTP ' + res.status)), true)
          beacon('chat-link open failed: ' + ((data && data.error) || ('HTTP ' + res.status)))
        }
      } catch (e) {
        showToast('⚠️ 打开失败：' + (e && e.message ? e.message : 'network error'), true)
        beacon('chat-link open error: ' + (e && e.message))
      }
    }

    const interceptChatLink = (ev) => {
      if (ev.defaultPrevented || ev.button !== 0 || ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey) return
      const target = ev.target
      if (target === null || typeof target.closest !== 'function') return
      const a = target.closest('a[href]')
      if (a === null) return
      let u
      try { u = new URL(a.href, location.href) } catch { return }
      if (u.pathname !== '/olk-notify/open') return
      const entryId = u.searchParams.get('entryId')
      if (entryId === null || entryId === '') return
      ev.preventDefault()
      ev.stopPropagation()
      openMailByEntryId(entryId)
    }

    return {
      name: 'dsh-outlook',
      inject: ['slots'],
      apply(ctx) {
        beacon('apply called (slot registration)')
        const stylesTag = ensureStyles()
        if (stylesTag !== null) {
          ctx.effect(() => () => {
            if (stylesTag.parentNode !== null) stylesTag.parentNode.removeChild(stylesTag)
          }, 'dsh-outlook: styles')
        }
        // Capture-phase delegation survives React re-renders of the chat
        // surface; the listener itself lives in this fiber's effect.
        document.addEventListener('click', interceptChatLink, true)
        ctx.effect(() => () => {
          document.removeEventListener('click', interceptChatLink, true)
          for (const t of toastTimers) clearTimeout(t)
          toastTimers.clear()
          for (const div of document.querySelectorAll('.olk-toast')) {
            if (div.parentNode !== null) div.parentNode.removeChild(div)
          }
        }, 'dsh-outlook: chat link interception')
        ctx.slots.inject('sidebar.footer.action', () => ctx.slots.register(
          {
            name: 'sidebar.footer.action',
            id: 'newmail-badge',
            order: 10, // above Whale Bay (order 90); matches the old footArea-first position
            label: 'Outlook 新邮件',
          },
          (props) => React.createElement(NewMailRow, props),
        ))
      },
    }
  },
})
