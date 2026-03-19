# 挑战杯项目 README

## 项目简介
本仓库用于整理《庄子》文本中的分句、字义注释与跨句阐释，并生成可在浏览器中交互查看的数据与页面。

当前数据流如下：

`zz.xlsx` -> `zz_structured.xlsx` -> `data.js` -> `index.html`（首页） / `reader.html`（阅读器）

## 当前数据规模（来自 `data.js`）
- 篇章数：3
- 分句数：956
- 有注释分句键数：716
- 跨句阐释条数：1967
- 篇章：人间世、逍遥游、养生主

## 目录说明
- `rebuild_zhuangzi_tables.py`：核心重建脚本。从原始工作簿抽取并分类为标准化三张表（分句、注释、阐释）。
- `generate_web_data.py`：将标准化 Excel 转换为前端可直接消费的 `data.js`。
- `index.html`：项目首页（统计总览 + 篇章入口）。
- `reader.html`：阅读页（搜索、句段高亮、弹窗查看注释/阐释）。
- `template.xlsx`：标准化输出模板（三张目标表结构）。
- `zz.xlsx`：原始/工作数据源（含源 Sheet 与同步后的目标表）。
- `zz_structured.xlsx`：标准化结果文件。
- `data.js`：前端数据文件（`window.ZZ_DATA`）。
- `split.py`：文本断句与 CSV 导出实验脚本（不在主流程中）。
- `sentences_output.csv`：`split.py` 输出示例。

## 环境与依赖
如果运行 Python 脚本，先激活指定环境：

```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh
conda activate daily
```

主要依赖：
- `openpyxl`

## 使用流程
### 1) 重建结构化表
```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh && conda activate daily && python rebuild_zhuangzi_tables.py
```

默认会：
- 读取 `zz.xlsx`
- 基于 `template.xlsx` 生成 `zz_structured.xlsx`
- 将 `small_sentences` / `annotations` / `interpretations` 回写并同步到 `zz.xlsx`

### 2) 生成前端数据
```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh && conda activate daily && python generate_web_data.py
```

输出 `data.js`（全量 JSON 载荷挂载到 `window.ZZ_DATA`）。

### 3) 浏览页面
直接打开 `index.html`（或用本地静态服务）：
- 首页可查看统计信息并跳转到各篇章
- 点击入口进入 `reader.html` 后可使用左侧篇章切换
- 原文分句展示（可点句查看字义注）
- 跨句阐释列表（点击后高亮对应句段）
- 当前篇章范围内全文检索（原文/注释/阐释）

## 注意事项
- `data.js` 是大文件，脚本每次生成会整体覆盖。
- `rebuild_zhuangzi_tables.py` 内含较多启发式规则（注释/阐释分类与区间扩展），修改前建议先备份数据文件。
