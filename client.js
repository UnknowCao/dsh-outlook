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
    beacon('factory loaded (v2.6.0 slot-based)')

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
            pending.slice(0, 10).map((m, i) => React.createElement('div', { key: i },
              '[' + (m.sender || '?') + '] ' + (m.subject || '(无主题)') + ' · ' + (m.received || ''))),
            React.createElement('div', { className: 'olk-popover-hint' }, '在会话里问「有什么新邮件」可读取并清空')),
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
