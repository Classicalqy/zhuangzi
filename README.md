# 庄子注疏数字平台

本仓库用于生成并维护一个静态网页平台，包含三个并行模块：

- 庄子注疏可视化（分句阅读、单句注释、跨句阐释、检索）
- 后人对前人的评价（评价关系可视化）
- 哲学词典（概念、内涵外延、注释条目）

平台目前聚焦《庄子》内七篇中的三篇：`逍遥游`、`养生主`、`人间世`。

---

## 仓库结构

### 页面文件
- `index.html`：首页（平台入口、统计、篇章入口、概览）
- `reader.html`：庄子注疏可视化页面
- `criticism.html`：后人对前人的评价页面
- `philosophy.html`：哲学词典页面

### 数据源（Excel）
- `zz_structured.xlsx`：主注疏数据（`small_sentences` / `annotations` / `interpretations`）
- `criticism.xlsx`：评价关系数据（`eval_edges` / `ref_notes` 等）
- `philosophy.xlsx`：哲学词典数据（3 个 sheet）

### 生成脚本
- `generate_web_data.py`：`zz_structured.xlsx -> data.js`
- `generate_criticism_data.py`：`criticism.xlsx + zz_structured.xlsx -> criticism_data.js`
- `generate_philosophy_data.py`：`philosophy.xlsx + zz_structured.xlsx -> philosophy_data.js`

### 前端数据文件（脚本生成）
- `data.js`
- `criticism_data.js`
- `philosophy_data.js`

---

## 环境与依赖

按项目约定，运行 Python 脚本前请先激活环境：

```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh
conda activate daily
```

主要依赖：
- `openpyxl`

---

## 使用流程

### 1) 更新主注疏阅读数据
当你修改了 `zz_structured.xlsx` 后，执行：

```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh && conda activate daily && python generate_web_data.py
```

会生成/覆盖：`data.js`。

### 2) 更新“后人对前人的评价”数据
当你修改了 `criticism.xlsx` 后，执行：

```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh && conda activate daily && python generate_criticism_data.py
```

会生成/覆盖：`criticism_data.js`。

### 3) 更新“哲学词典”数据
当你修改了 `philosophy.xlsx` 后，执行：

```bash
source /opt/homebrew/Caskroom/miniforge/base/etc/profile.d/conda.sh && conda activate daily && python generate_philosophy_data.py
```

会生成/覆盖：`philosophy_data.js`。

### 4) 本地查看
可直接双击 `index.html`，或使用本地静态服务（推荐）：

```bash
python -m http.server 8000
```

然后访问：
- `http://localhost:8000/index.html`

---

## 数据约定（简要）

### `zz_structured.xlsx`
- `small_sentences`：分句原文
- `annotations`：单句字义注释
- `interpretations`：跨句阐释（支持 `tendency` 字段：儒/释/道）

### `criticism.xlsx`
- `eval_edges`：有态度评价关系（`stance=Y/N`，可含 `highlight`）
- `ref_notes`：参考条目（不画关系线）

### `philosophy.xlsx`
每个 sheet 采用同类结构：
- 前三列：`text_id`、`概念`、`内涵外延`
- 后续列：`注1`、`注2`、…（可空）

---

## 页面能力概览

### `reader.html`
- 按篇章阅读分句
- 点击分句查看“字义注释 + 涉及该句的跨句阐释”
- 跨句阐释支持按儒/释/道筛选
- 检索（原文/注释/阐释）

### `criticism.html`
- 按分句展示评价关系
- 关系颜色区分：支持/反对
- 参考条目单独展示
- `highlight` 可高亮分句中的关键词

### `philosophy.html`
- 按篇章浏览概念词条
- 展示“内涵外延 + 注释列表”
- 页面内关键词筛选

---

## 注意事项

- `data.js`、`criticism_data.js`、`philosophy_data.js` 均为生成产物，重新生成会整体覆盖。
- Excel 打开时可能产生临时锁文件（如 `~$*.xlsx`），提交或处理前请忽略/删除。
- 若浏览器未显示最新改动，请强制刷新缓存。
