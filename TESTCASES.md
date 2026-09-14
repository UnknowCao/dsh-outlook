# dsh-outlook 侧边栏角标功能 — 测试用例集

覆盖链路：**NewMailEx watcher → index.js 状态路由 → 客户端 20s 轮询 → 角标渲染 → 弹层 → outlook_check_new 清空**。

标注说明：`[自动]` 可用脚本/工具验证；`[手动]` 需要浏览器观察；`[回归]` 对应历史上真实发生过的缺陷。

---

## A. 数据链路（watcher → 状态 API）

| # | 用例 | 步骤 | 预期 | 类型 |
|---|---|---|---|---|
| A1 | 单封新邮件推送 | 向自己发送一封邮件（经 send_mail 确认流程） | ≤20s 后 `GET /olk-notify/api/state` 返回 `pendingCount:1`，pending 含正确的 entryId/subject/sender/received | [自动] |
| A2 | 批量到达 | 一封邮件同时抄送自己或触发多 EntryId 的 NewMailEx | 每个 EntryId 各生成一条 pending，无丢失 | [自动] |
| A3 | received 用真实时间 | 观察 A1 返回的 received | 等于邮件 `ReceivedTime`，非事件处理时刻（回归 #7：曾用 Get-Date） | [自动][回归] |
| A4 | watcher 崩溃自愈 | `taskkill` 掉 watcher 的 powershell 进程 | `.olk-watch-debug.log` 记录 exited，30s 后 respawn，`hello` 重新上线，期间 state 路由仍可响应（unread 为旧值） | [自动] |
| A5 | Outlook 未启动 | 退出 Outlook 后重启 DSH | watcher 无限等待不崩溃；Outlook 启动后发出 hello（unread 正确） | [手动] |
| A6 | pending 上限 | 模拟 >50 封未取走 | 队列保留最新 50 封（`shift()` 丢最旧），计数 = 50 | [自动] |
| A7 | check_new 清空 | 队列非空时调用 `outlook_check_new` | 返回全部 pending 且 `count` 一致；再次调用返回 `count:0`；角标在下一轮询后消失 | [自动] |
| A8 | watcher 禁用开关 | 设 `OLK_WATCH=0` 后重启 DSH | watcher 不启动，`state` 返回 `pendingCount:0`、`unread:null`，工具其余功能正常 | [手动] |

## B. 角标渲染（client.js NewMailRow）

| # | 用例 | 步骤 | 预期 | 类型 |
|---|---|---|---|---|
| B1 | 零状态隐藏 | `pendingCount:0` | 角标 `hidden`，按钮仅显示图标+"Outlook" | [手动] |
| B2 | 数字居中 | pending=1 | "1" 在 16px 红圆内水平垂直居中（回归：曾因 text-align/line-height 偏移） | [手动][回归] |
| B3 | 双位数扩展 | pending≥10 | `min-width:16px` + `padding:0 4px` 使圆角胶囊变宽，数字仍居中 | [手动] |
| B4 | 角标不被裁剪 | 展开态观察角标 | `top:-4px;right:-4px` 探出按钮的部分完整显示（回归：曾被按钮 overflow:hidden 裁角） | [手动][回归] |
| B5 | 轮询周期 | 发信后计时 | ≤20s 内角标出现（POLL_MS=20000） | [手动] |
| B6 | 收起态 | 收起侧边栏 | 36px 圆钮上角标仍显示于右上角，不遮挡图标 | [手动] |
| B7 | tooltip | 悬停按钮 | title 列出最新 8 封 `[发件人] 主题`，无新邮件时显示"暂无新邮件" | [手动] |
| B8 | aria | 检查按钮属性 | `aria-label` = "Outlook 新邮件 N 封"（0 时"Outlook 邮件"）；`aria-haspopup:dialog` | [手动] |
| B9 | 服务器重启期轮询 | 重启 DSH 期间观察 | fetch 失败被静默吞掉，恢复后下一轮询正常（无报错刷屏） | [手动] |

## C. 弹层（popover）

| # | 用例 | 步骤 | 预期 | 类型 |
|---|---|---|---|---|
| C1 | 打开 | 点击按钮 | 弹层出现在按钮右侧 8px，标题"📬 新邮件（N）"，列出最新 10 封（发件人/主题/时间） | [手动] |
| C2 | 关闭-外部点击 | 点击弹层外任意处 | 弹层立即移除，document 监听器解绑 | [手动] |
| C3 | 关闭-toggle | 再次点击按钮 | 弹层关闭（点击按钮不触发 C2 的外部关闭） | [手动] |
| C4 | 空列表 | pending=0 时点击 | 显示"新邮件（0）"+ 底部提示"在会话里问『有什么新邮件』可读取并清空" | [手动] |
| C5 | 卸载清理 | 禁用插件 / 刷新页面 | 弹层、按钮、`dsh-outlook-styles` 标签全部移除，无 DOM 残留 | [手动] |

## D. 共存与样式回归（槽位化 v2.6.0）

| # | 用例 | 步骤 | 预期 | 类型 |
|---|---|---|---|---|
| D1 | 排序 | 展开侧边栏 | Outlook 在鲸湾之上、两者在设置区上方（order 10 < 90） | [手动] |
| D2 | 高度一致 | 与设置按钮并排比较 | 展开态均 42px 高、12px 圆角、同内边距/字体（回归：曾在纵向 flex 列里被 flex:1 拉伸） | [手动][回归] |
| D3 | 悬停一致 | 分别悬停三按钮 | 同样的 `--dsw-alias-interactive-bg-hover` 底色（回归：曾用不存在的 --dsh- 变量 + 内联背景覆盖） | [手动][回归] |
| D4 | 无 dsh-dock 场景 | 停用 dsh-dock 后刷新 | 自带的 `[data-slot]` 堆叠 CSS 使条目仍垂直排列，不塌成横排 | [手动] |
| D5 | rail 间距 | 收起态对比 | 圆钮边距 `8px 0 10px` 与设置按钮一致 | [手动] |

## E. 服务端路由

| # | 用例 | 步骤 | 预期 | 类型 |
|---|---|---|---|---|
| E1 | 方法限制 | POST `/olk-notify/api/state` | 405 | [自动] |
| E2 | 鉴权 | 无 cookie 请求两个路由 | 401（web 层统一拦截） | [自动] |
| E3 | beacon 上限 | POST client-log 200KB body | 413，连接关闭，日志不写入 | [自动] |
| E4 | beacon 噪音 | POST 非 JSON body | 204 静默忽略 | [自动] |
| E5 | 日志轮转 | 写入 >1MB 后观察 | `.olk-watch-debug.log` 截断重写，不无限增长 | [自动] |

## F. 已知边界（当前设计不处理，避免误报）

- pending ≥ 100 时角标显示实际数字（无 "99+" 截断）——如需可作后续改进；
- 多浏览器标签页各自轮询，共享同一服务端队列：`check_new` 被任一会话调用后所有页面角标归零；
- 自发邮件（Send 到自己）NewMailEx 有 ~10-20s 投递延迟，属 Exchange 行为，非插件缺陷。

## G. 点击打开邮件（v2.10.0）

| # | 用例 | 步骤 | 预期 | 类型 |
|---|---|---|---|---|
| G1 | 桥 open_mail | search_mail 取 entryId 后直接跑桥 `OLK_ACTION=open_mail` | 桌面 Outlook 弹出检查器窗口，返回 `opened:true` + subject/sender（已实测通过） | [自动] |
| G2 | 工具 | 调用 `outlook_open_mail` | 同 G1，纯本地无确认门 | [自动] |
| G3 | 会话链接 | agent 在聊天里贴 `[打开](http://127.0.0.1:3080/olk-notify/open?entryId=…)` 并点击 | 新标签页显示"已在 Outlook 中打开 + 主题"，桌面弹出邮件窗口 | [手动] |
| G4 | 弹层点击 | 侧边栏弹层中点击一封邮件 | POST 成功，条目立即从列表移除、角标即时减 1；桌面弹出邮件窗口 | [手动] |
| G5 | entryId 白名单 | GET `/olk-notify/open?entryId=<script>` | 400 / "entryId 无效"，不触桥 | [自动] |
| G6 | POST body 上限 | POST >8KB body | 413，连接关闭 | [自动] |
| G7 | 跨域 Origin | 携带异源 Origin 头请求三个 olk-notify 路由 | 403 | [自动] |
| G8 | GBK 代码页回归 | 在系统代码页为 GBK 的 shell 里跑桥 | 解析成功无乱码（回归：无 BOM UTF-8 曾被 PS5.1 按 ANSI 误读致整文件解析失败，已加 BOM 修复并实测通过） | [自动][回归] |
| G9 | 打开即已读 | 打开一封 unread:true 的邮件后用 search_mail 复查 | 桥返回 `wasUnread:true`；复查该邮件 `unread:false`（已实测通过）。注意：edit 工具重写 ps1 会剥掉 BOM，改后必须重新补 BOM | [自动] |
| G10 | 角标同步减 1 | 打开 pending 中的一封（弹层点击或会话链接） | host 从 notify.pending 移除该 entryId 且 unread--；弹层本地列表即时移除；其它标签页 ≤20s 轮询后一致 | [自动][手动] |

---

## 自动化脚本（配合 A 组用例）

```powershell
# 状态断言：期望 pendingCount = N
$v = '<cookie>'; (Invoke-RestMethod -Headers @{Cookie=$v} 'http://127.0.0.1:3080/olk-notify/api/state').pendingCount

# watcher 存活断言
(Get-Process powershell | Measure-Object).Count -ge 2  # 桥接/watcher 至少可存在

# 触发一封真实新邮件：在会话中调用 outlook_send_mail（需用户确认）后执行 A1 断言
```
