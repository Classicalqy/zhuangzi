# Repo 总览

## 一句话定位
这是一个围绕《庄子》文本注疏数据的“Excel 清洗与结构化 + 前端可视化”仓库。

## 架构分层
1. 数据源层：`zz.xlsx`
2. 结构化处理层：`rebuild_zhuangzi_tables.py`
3. 前端数据层：`generate_web_data.py` -> `data.js`
4. 展示层：`index.html`（首页） + `reader.html`（阅读器）

## 主流程细节
1. `rebuild_zhuangzi_tables.py`
- 自动识别源 Sheet 的表头行与数据起始行。
- 抽取分句列，写入 `small_sentences`。
- 对每个单元格按句切分后做“注释/阐释”分类：
  - 注释：偏音义、异文、训诂、人名地名解释等。
  - 阐释：偏哲理阐发或跨句义理说明。
- 将连续阐释句段合并并做区间扩展（词面相关度 + 邻近边界）。
- 输出 `annotations` 与 `interpretations`。
- 可选择同步结果回写到 `zz.xlsx`。

2. `generate_web_data.py`
- 读取 `zz_structured.xlsx` 的三张标准化表。
- 建立：
  - `texts`: 篇章与分句（含 `has_note`）
  - `annotations_by_key`: 以 `text_id-sentence_id` 聚合的单句注释
  - `interpretations`: 跨句阐释数组
- 生成 `window.ZZ_DATA = ...` 到 `data.js`。

3. `index.html` + `reader.html`
- 单页原生 JS，无框架依赖。
- `index.html`：首页（项目统计、篇章入口）。
- `reader.html`：阅读器（原文、注释、阐释、搜索）。
- 交互：
  - 篇章切换
  - 原文句点击弹窗（注释 + 覆盖该句的阐释）
  - 跨句阐释点击后句段高亮 + 定位 + 弹窗
  - 当前篇章范围检索（原文/注释/阐释）

## 关键文件关系
- 输入：`zz.xlsx`, `template.xlsx`
- 中间产物：`zz_structured.xlsx`
- 前端产物：`data.js`
- 展示页：`index.html`（首页）, `reader.html`（阅读器）

## 现状与风险点
- 规则驱动分类（启发式）可解释性强，但边界样本可能误分。
- 大体量 `data.js` 适合静态演示，不适合频繁网络传输场景。
- `split.py` 为实验脚本，和主流程解耦。
