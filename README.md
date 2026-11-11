# REI Location Print Assistant / REI 库位打印助手

A local-first browser tool that turns WMS location-stock Excel exports into one printable worksheet per virtual REI location. No login, API, backend, external fonts, or file upload service.

## 使用方法

1. 打开工具，选择一个或多个 WMS `.xlsx` 导出文件，或直接拖入。
2. 两种方式都支持：**多个文件各包含一个库位**；**一个文件包含多个库位**。也可以选择多个不重叠的多库位文件。
3. 默认转换文件中的全部 REI 库位；可指定 REI-1～REI-100 内的连续范围。范围含首尾，例如 REI-10～REI-20 是 **11 张工作表**。
4. 用“选日期”打开日历，或用“今天”恢复当天日期；打印日期统一用于所有标签页和下载文件名。可填写作业人，查看每个库位的预览，再点击“生成并下载 Excel”。
5. 用 Excel/WPS 打开下载文件，打印时选择“整个工作簿”，先检查打印预览，再打印。

输入取自 WMS 的“库位库存数据”页（支持多个具有相同表头的数据页），按字段名匹配，不按文件名或固定列序号猜测：

| 输入表头 | 输出列 | 处理 |
| --- | --- | --- |
| SKU编码 / SKU | SKU | 保持编码原文，写为文本 |
| Barcode / 条码 | Barcode | 保留文本前导零，缺失时留空并提示 |
| 库位 / 库位编码 / Location | Rei | REI 大小写统一，按数字排序 |
| 库位库存 / 库存数量 / System quantity | 展开行数 | 数量为 2 就生成两行；使用账面库存，不使用可用库存 |
| 无 | Location | 留空，供现场填写实际物理库位 |

输出沿用 RT 作业单结构：STARTZEIT / ENDZEIT、Name / Datum、SKU / Barcode / Rei / Location。每个库位一张工作表，A4 纵向，固定打印区域，宽 1 页、高 1 页，至少保留 18 行货物填写区域。日期按本机当天默认，支持调整；不会带入模板中旧日期或其他员工姓名。

### 可用库存预警

导入时读取可选的「可用库存」列。可用库存小于库位库存（包括 0）时，页面按库位与 SKU 给出预警；等于或大于库位库存时不预警。只对当前所选范围提示。空白不当作 0，缺少列或无效值明确提示无法判断。

所有这些预警均不阻止下载、不改变打印数量：库位库存为 3、可用库存为 0，仍打印 3 行。打印表继续沿用原四列，不插入可用库存列或预警文字，不改变单页设置。

### 数据检查

- 同一个库位出现在两个文件中时，**阻止下载**并提示移除重叠文件。不会把两次库存快照相加，也不会自动挑“最新”的文件。
- 同一工作表中，同一库位的多条记录均按原始库存展开，不擅自去重。同一库位跨多个数据页重复出现时，阻止合并并提示，避免把复制的快照累加。
- 同一个 REI 编码出现在多个库区时阻止合并，需先在 WMS 筛选唯一库区。
- 非 REI-1～REI-100 的记录不进入作业单，界面显示排除数量。
- 范围内缺少库位记录时，不判断为无货。默认阻止下载；可明确选择为缺失库位生成“未提供该库位数据”提示页。
- 源表明确有记录且数量为 0 时，输出“源表库存为 0”，与缺少数据区分；不代表现场已核实为空。
- 负数、小数、非法库存、含公式的标识/数量单元格、重复必需表头会报错，不静默跳过。
- 数量较多（超过 35 行）会提示单页缩放可能影响可读性，仍保留全部行。网页预览最多展示前 80 行，Excel 不截断。
- 编码以文本为准；纯 `000000` 数字格式可还原前导零，其他特殊数字编码格式或不安全整数会拒绝。Excel 已损失的数字精度无法恢复。
- 每文件最多 10 MB、50,000 数据行；每批最多 100 个文件、20,000 件展开行。大批次请先在 WMS 缩小导出范围。

## 本地运行 / Development

Requires Node.js 22 or later.

```sh
npm ci
npm run dev
```

Open the local URL printed by Vite. The preview includes a **clearly labelled synthetic demo** of REI-10 through REI-20; no user stock data is committed to this repository.

```sh
npm test
npm run build
npm run preview
```

`dist/` is a static website. It can be served by any static web server. This repository does not deploy automatically or connect to WMS.

## 可双击打开的离线版 / Portable offline HTML

```sh
npm run portable
```

Open `dist/REI-Print-Assistant.html` directly in Chrome or Edge. Its CSS, UI and Excel conversion module are embedded in one HTML file. It works without a server or network once downloaded. The same file is suitable for handing to a colleague; it contains the application and synthetic demo only, never your uploaded inventory.

Source files and generated workbooks are not persisted in browser storage; reopening the page requires importing files again. Existing WMS data and the separate `WMS_Inventory_check` project are not modified. This release automates **local conversion only**; downloading from the signed-in WMS website remains manual.

## Validation

```sh
npm test
npm run portable
node scripts/smoke-portable.mjs
```

The portable smoke test runs the actual bundled application in a local DOM environment and reopens its downloaded demo workbook. An optional `WMS_SAMPLE` environment variable validates the supplied REI-22 example without committing source inventory to the repository.

Vitest covers multi-file and consolidated imports, quantity expansion, numerical ordering, inclusive ranges, overlap rejection, cross-area conflicts, missing-versus-zero stock, text identifiers, invalid inputs, generated XLSX print settings, and UI behavior. Generated XLSX files are re-opened to verify values and print setup.

Excel/WPS and printer drivers can affect physical print layout. A4 portrait and fit-to-one-page settings are embedded; verify the downloaded workbook's print preview on the machine that will print it. No rows are deleted to meet a page target.

## English summary

Import one or many WMS `.xlsx` exports, choose all detected REI locations or an inclusive numeric range, and preview/download one workbook with numerically ordered location tabs. Each unit of stock becomes one row. Destination `Location` stays blank for field work. Overlapping snapshots are blocked, missing locations are explicitly marked rather than treated as empty, and every sheet has A4 portrait, one-page print settings. All processing stays in the browser; there is no WMS login or API automation in this release.
