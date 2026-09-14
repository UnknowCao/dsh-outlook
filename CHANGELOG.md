# Changelog

本项目的版本叙事遵循"为什么改"优先于"改了什么"。每条都先给动机。

## 2.11.0 — 聊天里的"打开"链接点击不再留死标签页

**为什么**：agent 在会话里贴的 `[打开](…/olk-notify/open?entryId=…)` markdown 链接走 GET 导航——每点一次就多出一个"已在 Outlook 中打开 / 此页面可关闭"的标签页，最常用的"看一眼这封邮件"动作反而在浏览器里留下一排垃圾标签。同时该链接硬编码 `127.0.0.1:3080`，自定义端口部署时点击直接连接失败（README 已列为已知边界）。

**改了什么**：
- client 半新增捕获阶段点击拦截：普通左键点击路径名为 `/olk-notify/open` 且带 `entryId` 的锚点时，按**匹配 pathname 而非完整 URL** 识别（任意端口都命中），改为同源 POST——与侧边栏弹层完全相同的已验证通道——Outlook 弹窗打开邮件，**零新标签页**，页面底部给出 2.6s toast 反馈（成功含主题，失败含错误并 beacon 上报）。这同时修复了自定义端口部署的链接失效边界：拦截后总是打向页面自身 origin。
- **GET 导航降级整体移除**（按需求决策）：`/olk-notify/open` 改为 POST-only JSON 路由，结果页（"此页面可关闭"）、关闭按钮全部删除——正常路径不再产生标签页，就没有必要维护第二形态的降级面。修饰键/中键点击该链接会得到 405（有意行为）。GET 非法 entryId 曾误返 500 的问题随 GET 分支一并消失。
- 拦截监听器与 toast 均注册在 fiber effect 中，插件停用/更新时自动移除（监听器、定时器、DOM 三者都清理）。
- pending 队列清理/角标减 1 逻辑不变：侧边栏与聊天链接现在走完全相同的 POST 处理器。

## 2.10.1 — 公共发布前的移植性加固

**为什么**：插件将面向任意用户的电脑分发，不能假设安装目录可写、Outlook 一定装过、README 永远是最新的。

**改了什么**：
- watcher 诊断日志从插件安装目录（全局安装可能只读）移到 `%TEMP%\dsh-outlook-watch-debug.log`。
- 桥在 COM 对象创建前做 HKCR 预检：未装经典 Outlook（或只有 UWP "New Outlook"）时返回可操作的错误信息，而非天书般的 ActiveX 报错。
- 桥/watcher 调用改为 `Get-Content -Raw -Encoding UTF8`，不再依赖文件 BOM（编辑工具会剥 BOM 的坑已两次复现）。
- README 刷新到 18 工具/43 用例的现状，补会议四件套与 `outlook_open_mail`，已知边界一节更新（端口假设、OLK_WATCH 开关、日志位置）。
- `meeting_update`/`meeting_cancel` 工具描述与实际行为对齐（统一两段式确认）。

## 2.10.0 — 会话里点击即打开邮件 + 桥的编码隐患修复

**为什么**：搜索/新邮件结果只给了 entryId，用户想看原文必须自己在 Outlook 里翻——最常用的"看一眼这封邮件"动作反而最绕。同时在实测中暴露：`olk.ps1`/`watch.ps1` 是无 BOM 的 UTF-8，PowerShell 5.1 的 `Get-Content -Raw` 按系统代码页（GBK）解码时中文全乱、整文件解析失败——生产环境只是因系统代码页恰好非 GBK 而侥幸未爆。

**改了什么**：
- 桥新增 `open_mail` 动作：`GetItemFromID().Display()` 在桌面 Outlook 检查器窗口打开邮件，并显式置 `UnRead=$false`（打开即视为已读，返回 `wasUnread`）。纯本地动作，不离开机器，不走确认门。
- host 新增 `/olk-notify/open` 路由：GET（`?entryId=`，返回极简 HTML 结果页，供会话聊天里的 markdown 链接点击）+ POST（JSON，供侧边栏弹层）。沿用 sameOrigin 守卫（浏览器导航无 Origin 头，与既有路由同信任级）；entryId 白名单正则（`[A-Za-z0-9+/=_-]{20,512}`）。打开成功后 host 把该 entryId 移出 `notify.pending` 并 `unread--`，角标计数随之减 1。
- 新工具 `outlook_open_mail`（共 18 个工具）；`outlook_search_mail`/`outlook_check_new` 描述追加"给每封邮件贴可点击打开链接"的提示，agent 展示结果时直接给 `[打开](…/olk-notify/open?entryId=…)` 链接。
- 侧边栏新邮件弹层每封邮件变为可点击按钮：POST 打开路由，成功后条目立即从列表移除（角标即时减 1，不等 20s 轮询），失败 beacon 上报。
- `olk.ps1`/`watch.ps1` 前置 UTF-8 BOM，根除 PS5.1 ANSI 误读。

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
