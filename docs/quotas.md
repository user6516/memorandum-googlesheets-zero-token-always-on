# Google 配额与运行成本

## 不消耗 AI token

程序不调用 OpenAI、Gemini 或其他大语言模型 API，不需要 AI API Key、个人 AI 订阅或持续打开聊天。AI 可以辅助开发代码，但日常运行仅由 Google 表格、Apps Script、持久化属性和 MailApp 完成。

**提醒执行的 AI token 消耗为零。** 本项目也没有 AI API 账单。开发时使用 AI 聊天服务的消耗，不属于备忘录程序的运行成本。

这不代表所有 Google 服务无限使用，也不保证账号自身的付费方案或将来服务政策不变。

## 常用配额

下表根据 **2026-11-13** 查阅的 [Google 官方配额文档](https://developers.google.com/apps-script/guides/services/quotas) 整理。实际以官方当前标准及账号条件为准，试用账号可能另有限制。

| 项目 | 普通个人账号 | Google Workspace |
| --- | --- | --- |
| 邮件收件人人次 | 100 / 天 | 1,500 / 天 |
| 自动触发器累计运行时间 | 90 分钟 / 天 | 6 小时 / 天 |
| Properties 属性读写 | 50,000 / 天 | 500,000 / 天 |
| 单次脚本最长运行时间 | 6 分钟 | 6 分钟 |
| 每个用户、每个脚本的触发器数量 | 20 | 20 |
| 单条属性大小 | 9KB | 9KB |
| 单个属性存储总大小 | 500KB | 500KB |

邮件额度按收件人人次计算：一封发给一个地址，计一次；给同一个地址连续发十封，计十次，不是十个不同地址才计十次。配额按用户计算，同账号其他脚本也可能共享。每日配额按首次请求起算的24小时重置，不应当作北京时间零点刷新。

## 每十分钟检查是否会耗尽运行额度？

每十分钟约144次/天，每五分钟约288次/天。运行额度扣的是实际执行时长，不是两个触发点之间的等待时间。

例如平均每次15秒，十分钟检查的主体耗时约为36分钟/天；五分钟检查约为72分钟/天。这里只是预算示例，编辑同步、服务耗时和统计自身开销也需考虑。不能据此承诺不会超额。

程序没有到点就快速退出；定时检查不读任务表；元数据未变化时不重复写回；后台遇到同步锁立即退出。即使如此，服务变慢、频繁编辑和其他脚本仍可能使账号超额。

## 如何查看

在 Apps Script 函数下拉框选择 `checkRuntime`，运行后查看执行日志：

- 最近约24～25小时，本项目已记录调用的主体耗时、次数、平均与最长耗时、失败次数。
- 当前剩余邮件收件人人次，由 [MailApp.getRemainingDailyQuota()](https://developers.google.com/apps-script/reference/mail/mail-app#getRemainingDailyQuota()) 查询。

主体耗时不含统计逻辑自身、其他项目和未记录操作；手动运行tick也计入，统计失败可能缺样。这是估算，不是 Google 官方剩余运行分钟数。

实际执行时长和失败信息应结合左侧“执行记录”查看；也可使用 [Apps Script Dashboard](https://developers.google.com/apps-script/guides/dashboard)。达到配额通常会抛出错误并停止相关执行，不会自动获得额外资源。

## Google 官方资料

| 内容 | 链接 |
| --- | --- |
| 服务额度、限制与超额错误 | [Quotas for Google Services](https://developers.google.com/apps-script/guides/services/quotas) |
| 云端运行与平台概述 | [Apps Script overview](https://developers.google.com/apps-script/overview) |
| 从表格创建脚本 | [Container-bound Scripts](https://developers.google.com/apps-script/guides/bound) |
| 编辑、更改和定时触发器 | [Installable Triggers](https://developers.google.com/apps-script/guides/triggers/installable) |
| 定时接口与时间精度 | [ClockTriggerBuilder](https://developers.google.com/apps-script/reference/script/clock-trigger-builder) |
| 发信与剩余额度接口 | [MailApp](https://developers.google.com/apps-script/reference/mail/mail-app) |
| 持久化属性 | [Properties Service](https://developers.google.com/apps-script/guides/properties) |
| 授权与权限 | [Authorization](https://developers.google.com/apps-script/guides/services/authorization) |
| 执行与健康监控 | [Apps Script Dashboard](https://developers.google.com/apps-script/guides/dashboard) |
