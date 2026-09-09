# Changelog

本项目的版本叙事遵循"为什么改"优先于"改了什么"。每条都先给动机。

## 2.6.1 — 让收起侧边栏时角标永远可见

**为什么**：收起态角标探出按钮外沿（top:-4px），会被侧边栏收起动画的裁剪容器切掉或被上方行遮挡——用户看不到新邮件提醒，通知功能在最需要安静的 rail 形态下失效。

**改了什么**：收起态改用 14px 内置小红点（圆钮右上角内侧），展开态保持外置描边角标不变；角标加 `z-index:1`。

## 2.6.0 — 从 DOM 注入迁移到官方槽位

**为什么**：旧版把按钮直接插进 shell 的 footArea 容器并用双 MutationObserver 自愈。这换来了位置自由，但代价是三起真实事故：v2.5.1 之前的 observer 微任务风暴冻结整个页面（"Loading plugins…"）；对哈希 CSS 类名的子串匹配随 shell 升级静默失效；悬停反馈因内联样式优先级永远不渲染。

**改了什么**：整个按钮改为通过 `sidebar.footer.action` 槽位注册（与 dsh-dock 鲸湾菜单同一正规通道），React 全权托管生命周期——observer、自愈、类名探测全部删除（约 150 行防御代码归零）。`wide`/rail 几何来自槽位 props；弹层改为 React Portal（点击外部关闭保留）。order:10 保持原视觉位置（鲸湾之上）。

## 2.5.3 — 悬停反馈真的能渲染了

**为什么**：v2.5.2 修正了悬停色变量（`--dsh-` → `--dsw-`）但用户仍看不到底色——内联 `background:'transparent'` 的优先级永远高于样式表的 `:hover` 规则。

**改了什么**：透明底色从内联样式移入注入样式表，让 `:hover` 可覆盖。

## 2.5.2 — 与原生设置按钮逐项对齐

**为什么**：悬停用错了 token 前缀（`--dsh-`，变量不存在），收起态边距与原生设置按钮不一致（4px vs 8px/10px），并排时肉眼可见地不齐。

**改了什么**：hover token 对齐 `--dsw-alias-interactive-bg-hover`；rail 边距对齐 `8px 0 10px`；其余几何参数逐项核对确认本来就一致。

## 2.6.0 同批 — 深度 review 修复 14 项

安全审计发现的确定性 bug 与加固（全部有测试用例，见 TESTCASES.md）：

- **内容哈希绑定**（安全加固）：`confirmed:true` 重试的内容哈希必须与最后一次展示给用户的确认一致，否则强制重新确认——"展示 A、发送 B"的偷换在 host 层即被拦截（两种攻击场景实测拦截）。
- `search_mail` 的 `unread` 字段语义反转（`-not $it.UnRead` → `[bool]$it.UnRead`）。
- 未知 action 返回 `status:'ok'` + error 的矛盾组合 → 改为 throw。
- 附件保存：`GetFileName` 防路径穿越 + 同名自动 `(n)` 编号不覆盖。
- COM 调用串行化（Outlook 单线程，并发 spawn 互踩）。
- `client-log` 路由 64KB 上限（413）；调试日志 1MB 截断轮转。
- watcher `received` 改用邮件真实 `ReceivedTime`。
- `ResolveAll()` 失败列出未解析参会人并中止；会议室查询优先 SMTP 唯一解析。
- `account` 按默认收件箱 DeliveryStore 匹配。
- DASL LIKE 剥离 `%`/`_` 通配符；发送失败清理孤儿草稿；弹层 ARIA 语义（dialog）。

## 2.5.1 及更早

（历史版本未留存 changelog；quirks 与实战教训编码在 lib/olk.ps1 头部注释与 README "Quirks" 一节。）
