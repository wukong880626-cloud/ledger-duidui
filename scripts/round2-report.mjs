import {readFile,writeFile} from 'node:fs/promises';
const perf=JSON.parse(await readFile('output/round2/release-performance.json','utf8')),numeric=JSON.parse(await readFile('output/round2/release-numeric-performance.json','utf8')),checks=JSON.parse(await readFile('output/round2/browser-checks.json','utf8'));
const tests=await readFile('output/round2/unit-integration-tests.txt','utf8');if(!tests.includes('pass 44')||!tests.includes('fail 0')||checks.passed!==10)throw Error('验收尚未全部通过');
const type={csv:'CSV',xlsx:'XLSX 文本金额',xlsxnum:'XLSX 数值金额'};
const table=[...perf,...numeric].map(x=>{const stages=[x.import,x.compare,x.export];return `| ${type[x.format]} | ${x.rowsPerSide} | ${(x.import.ms/1000).toFixed(3)} | ${(x.compare.ms/1000).toFixed(3)} | ${(x.export.ms/1000).toFixed(3)} | ${Math.max(...stages.map(s=>s.maxGapMs)).toFixed(1)} | ${Math.max(...stages.map(s=>s.sampledRendererRSSPeakMB)).toFixed(0)} | ${x.exportedDifferences} |`;}).join('\n');
await writeFile('第二轮验收_v0.1.1.md',`# 台账对对看 · 第二轮验收 0.1.1

2026-10-22。44 项单元/集成测试、10 项第二轮浏览器检查及原有离线回归全部通过。全部使用新造数据，未读取真实客户资料。无公网部署、公开仓库、项目根许可证或系统安全设置变更。

## 已复现和修复

1. XLSX 在行循环内反复计算列范围，导致平方级耗时。旧版每侧 10,000 条双表导入 10.464 秒，主线程长任务 5.211 秒；缓存范围并移入随包本机 Worker 后，文本金额同规模导入 0.425 秒。解析、核对和导出均在本机计算线程执行，无外部服务。
2. 重复键说明重复拼接全部组员行号。1,000 条重复数据旧版说明总计 3,912,020 字符；改成同组缓存及前 20 个行号预览，每个源行仍逐条保留。5,000 条重复回归已完整导出 5,001 条待处理明细，说明每条不足 250 字符。
3. 原始数值 1.00000000000000001 曾被读取为 1 并可能误判一致。新增 XLSX 原始 XML 预检：超过 15 位有效数字或有读取舍入风险的关键数值隔离，保留 XML 数字原值。正常数值和高精度文本金额有回归覆盖，已丢失的源精度不自动恢复。
4. 底层表格库会静默删控制字符/DEL，并把 _x0041_ 或重叠转义文本解释成字符。现在控制字符、超长原值或安全转义后超长明确拒绝；字面转义文本保真导出，公式注入文本仍防护。
5. 异步任务增加修订校验；核对中清空不会回流旧结果，报告生成中清空或改设置会取消旧下载。两项均在 25,000 条差异的真实浏览器操作中通过。

## 实际规模和性能

同一台 macOS、Node.js 23.11.0、Chrome 155、本机 HTTP、固定依赖。每档为每侧行数、5 列、3 分量主键；B 表行列重排，A 元/B 万元，全部状态不同。每份报告均重新读取并验证全部差异条数。数值金额档位使用真正数值单元格及金额格式。

| 输入 | 每侧条数 | 双表导入 秒 | 核对 秒 | 导出 秒 | 最大心跳间隔 ms | 渲染 RSS 采样峰值 MiB | 核验导出条数 |
|---|---:|---:|---:|---:|---:|---:|---:|
${table}

以上是本机单次端到端观察值，包含 Playwright 动作及 60 ms 收尾等待，不是纯算法、统计分位或跨机器保证。主页面心跳周期 20 ms，间隔值包含正常周期；最终九档各阶段均未观察到超过 50 ms 的主线程 Long Task。Worker 本身耗时不计入主线程指标。

RSS 为每个 CLI 阶段调用期间每 50 ms 采样的隔离 Chrome 渲染进程合计，计入 Worker 和浏览器保留分配；不是整机内存、完整 Chrome 全进程内存、数据独占内存或精确瞬时峰值。GC/分配保留会使较小档位出现较高值。原始 JSON 同时保存主线程 JS 堆，未以它冒充全部内存。

旧 50,000 条 CSV 上限已做探索：移入 Worker 后导出约 1.439 秒、最大心跳间隔约 22.6 ms，但渲染 RSS 采样约 1,183 MiB。因此当前不承诺 50,000 条输入，保守收紧。

## 0.1.1 可用输入限制

建议日常先按每侧 10,000 条以内使用。硬上限为：每表 25,000 条、200 列；每文件 20 MB、150,000 单元格（含标题、所有表及读取矩形）；标题/说明/空行等额外记录最多 100 条；两侧异常合计 25,000 条。多字段异常可能比原行数多；超限明确拒绝完整核对，不返回部分结果或声称一致。

最多 8 个主键分量、20 个核对字段，键分量 256 字符；输入单格 32,766 字符，安全转义后不超过 32,767。XLSX 声明解压体积 50 MB、内部项 1,000 个、20 张表；不接受 ZIP64、分卷或宏。声明体积检查不是压缩炸弹安全扫描，不处理不可信来源。

25,000 条/侧、5 列虽已通过，导出仍可能使隔离渲染进程接近 1 GB；低内存或宽表应拆小。没有任意宽表、极限组合或低内存设备性能保证。

## 第二轮浏览器检查

${checks.checks.map((c,i)=>`${i+1}. ${c}。`).join('\n')}

另重跑 npm run test:browser，已启动页面在浏览器离线状态下导入、核对、下载及安全防护通过。冷启动指阻断全部外部网络但允许 127.0.0.1；不声称在开发者工具连本机请求也阻断后还能冷加载。日常断开互联网不阻止本机回环地址。

本机 Node 只提供静态文件；台账字节仅在浏览器和本机 Worker 内存处理，不 POST 给 HTTP 服务。运行无 CDN、遥测、AI API 或外部上传。

## 证据与复跑

- output/round2/unit-integration-tests.txt：44 项测试，新增回归在 test/round2.test.mjs。
- release-performance.json、release-numeric-performance.json 及日志：最终九档数据；reproduced-baseline-issues.json：缺陷复现；baseline-50k-performance.json、worker-exploratory-performance.json：旧上限探索，均在 output/round2。
- output/round2/browser-checks.json / .log、text-fidelity-browser-export.xlsx：浏览器检查和真实下载；samples/第二轮 为小型合成样本。
- scripts/round2-fixtures.mjs、round2-functional-fixtures.mjs、round2-browser-check.mjs、round2-browser-perf.mjs 可复跑。ZIP 排除可重建的大表/大份性能导出，保留生成代码、测量 JSON、日志、小样本和测试。
- ZIP CRC、解压后无 node_modules 启动及 Worker 文件校验见 output/package-validation.json；校验值见 release/SHA256SUMS.txt。

## 仍未测试

桌面 Excel/WPS 人工逐表打开、其他浏览器/Windows、真机手机、Finder 双击启动、低内存设备、真实复杂或加密工作簿、恶意 ZIP64/压缩炸弹和所有宽度/文本极限组合未验收。部分边界由集成测试验证，未全部在浏览器重做；真实客户资料未用于测试。

本版要求本机 HTTP；直接双击 dist/index.html 不属于支持方式。首轮工具拒绝 file:，本轮未重试或绕过。无剩余开发阻塞，但以上未测范围保留。
`);
await writeFile('测试结果.md','# 当前验收结果 · 0.1.1\n\n44 项单元/集成测试、10 项第二轮浏览器检查及原有离线回归通过。CSV、文本金额 XLSX、数值金额 XLSX 各完成每侧 1,000 / 10,000 / 25,000 条实测并核验完整导出。\n\n详细数据、修复和未测范围见 [第二轮验收_v0.1.1.md](第二轮验收_v0.1.1.md)，启动见 [README.md](README.md)。每表上限 25,000 条、每文件 150,000 单元格、异常合计 25,000 条；建议每侧 10,000 条以内。本版须使用本机 HTTP。\n');
await writeFile('更新说明_v0.1.1.md','# 0.1.1 更新说明\n\n- 修复 XLSX 导入平方级耗时，计算移到随包本机 Worker。\n- 重复键说明改为有界预览，逐条来源完整保留。\n- 识别原始 XML 数值精度风险，隔离并保留原值。\n- 修复控制字符静默删改及字面转义文本导出变化，完善长度防护。\n- 加入结构、单元格、解压体积和异常总量上限；每表收紧至 25,000 条。\n- 防止清空/改设置后旧结果回流或下载旧报告。\n- 44 项测试、10 项第二轮浏览器检查和九档压力验收通过，详见第二轮验收报告。\n\n须使用本机 HTTP；未改原台账，无公网部署或运行数据上传。保留 v0.1.0 ZIP，新增本版 ZIP。\n');
console.log('已生成 0.1.1 验收文档与更新说明。');
