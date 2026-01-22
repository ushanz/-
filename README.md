# 政策文本词频统计

## 项目简介

这是一个用于 **批量处理政策文本** 的脚本，支持从一个根目录递归扫描 `.pdf / .doc / .docx` 文件，完成：

1. **文本抽取**

   * PDF：优先抽取文本层（PyMuPDF），若疑似扫描件（文本过短），自动启用 **Tesseract OCR** 转文本
   * DOC：优先用 **LibreOffice headless** 转成 docx，再用 python-docx 抽取
   * DOCX：直接用 python-docx 抽取（含段落 + 表格）

2. **关键词词频统计**

   * 支持按“类别（category）- 关键词（keyword）”的词典结构统计
   * 输出两类统计结果：

     * `keyword_counts_long.csv`：每个文件每个关键词的词频
     * `category_counts_long.csv / category_counts_wide.csv`：每个文件每个类别的总词频（类别内关键词词频求和）

3. **导出 Gephi 可用的网络表**

   * 类别网络（category-level co-occurrence network）
   * 关键词网络（keyword-level co-occurrence network）
   * 同时提供节点表（nodes）与边表（edges），便于画“政策工具组合/共现”网络图

4. **完整处理清单 manifest**

   * 对每一个文件记录：是否成功、抽取后的文本长度、输出 txt 路径、失败原因
   * 用于保证“总文件数 = 成功 + 失败”，避免遗漏

---

## 适用场景

* 政策文本、公告、条例、规划、通知等文件批量清洗与文本抽取
* 构建政策工具词典、统计词频、做工具组合分析
* 输出网络数据，用 Gephi 做共现网络可视化（类似政策工具组合网络图）

---

## 依赖与环境要求（Windows 推荐）

### Python 依赖（建议在虚拟环境里）

```bash
pip install pymupdf python-docx pywin32
```

> `pywin32` 用于 Word COM 兜底（脚本以 LibreOffice 转换为主，Word 是备选）。

### 外部软件（必须）

1. **Tesseract OCR**

   * 安装后需要知道：

     * `tesseract.exe` 的路径（例如：`D:\Tesseract-OCR\tesseract.exe`）
     * `tessdata` 目录（例如：`D:\Tesseract-OCR\tessdata`）
   * 扫描版 PDF（图片 PDF）依赖它做 OCR

2. **LibreOffice**

   * 需要 `program` 目录，例如：`D:\LibreOffice\program`
   * 用于 `.doc -> .docx` 转换（headless 模式）

---

## 输入数据目录结构要求

脚本会递归扫描你指定的 `POLICY_ROOT` 下所有文件：

* 支持扩展名：`.pdf`, `.doc`, `.docx`
* 自动跳过临时文件：以 `~$` 开头的 Word 临时文件

目录可以任意层级。脚本会从相对路径中提取：

* `root_type`：相对路径第 1 级目录名
* `level1`：相对路径第 2 级目录名
* `level2`：相对路径第 3 级目录名

这些字段会写入统计表，便于你按“国家/省/市”等层级做汇总分析。

---

## 关键词词典结构

脚本内置一个 `KEYWORDS` 字典：

```python
KEYWORDS: Dict[str, List[str]] = {
  "Regulations(法规/规章)": [...],
  "Tax incentives(税收激励)": [...],
  ...
}
```

* 每个 key 是一个类别（category）
* value 是该类别下的一组关键词（keyword list）
* **类别总词频 = 该类别下所有关键词词频之和**

### 清洗关键词解释


```python
KEYWORDS[cat] = [k.strip() for k in kws if k and k.strip()]
```

作用：

* 去除关键词前后空格，避免匹配失败（比如 `" 规范 "`）
* 删除空字符串，避免出现“空关键词”导致异常统计

---

## 处理流程说明（核心逻辑）

### 1) 文本抽取逻辑

* `.docx`：python-docx 抽取段落 + 表格
* `.doc`：

  1. 优先 LibreOffice headless 转 docx
  2. 失败再尝试 Word COM（并在 RPC 失败时杀掉残留 WINWORD 重试）
* `.pdf`：

  1. PyMuPDF 抽取文本层
  2. 若文本太短（`MIN_TEXT_LEN_FOR_PDF`），判定为扫描件 → OCR
  3. OCR 会把每页渲染为图片，再用 `tesseract input stdout` 得到文本

### 2) 词频统计逻辑

* 将所有关键词预编译为正则（`re.escape` 子串匹配）
* 输出：

  * `keyword_counts_long.csv`：每条记录 = 文件 + 类别 + 关键词 + count
  * `category_counts_long.csv`：每条记录 = 文件 + 类别 + category_total
  * `category_counts_wide.csv`：每行一个文件，每列一个类别（便于直接透视/回归）

### 3) 网络构建逻辑（用于 Gephi）

**类别网络**

* 节点：category
* 边：同一文件中同时出现的两个 category 形成一条边
* 边权重提供两种：

  * `binary`：共现次数（出现于多少个文件）
  * `weighted`：每个文件对边贡献 `min(totalA, totalB)`，再累加（更像“强度共现”）

**关键词网络**

* 节点：`category::keyword`（唯一ID）
* 边：同一文件中同时出现的两个关键词节点形成一条边
* 同样提供 `binary` 与 `weighted(min)` 两种边权重

---

## 输出文件说明

输出目录：`OUT_DIR`

### 1) 文本输出

* `txt/`
  每个源文件对应一个 txt（文件名含 hash，避免重名覆盖）

### 2) 可追溯清单（建议优先看这个）

* `manifest.csv`
  每个文件一行，字段包括：

  * `status`：OK / FAIL
  * `text_len`：抽取文本长度
  * `txt_file`：成功时 txt 路径
  * `error`：失败原因

> 用它可以确保：**总文件数 = OK + FAIL**，并快速定位问题文件。

### 3) 词频统计

* `keyword_counts_long.csv`
  文件粒度 × 关键词粒度（适合做关键词层面的分析/网络构建）

* `category_counts_long.csv`
  文件粒度 × 类别粒度（适合做类别层面的统计与回归）

* `category_counts_wide.csv`
  宽表：每行一个文件，每列一个类别（最方便用 Excel/Pandas 做后续分析）

### 4) Gephi 网络文件

**类别网络：**

* `gephi_nodes_categories.csv`
* `gephi_edges_categories_binary.csv`
* `gephi_edges_categories_weighted.csv`

**关键词网络：**

* `gephi_nodes_keywords.csv`
* `gephi_edges_keywords_binary.csv`
* `gephi_edges_keywords_weighted.csv`

---

## 使用方法

1. 修改脚本顶部配置：

* `POLICY_ROOT`
* `OUT_DIR`
* `TESSERACT_EXE`, `TESSDATA_DIR`
* `LIBREOFFICE_PROGRAM_DIR`

2. 运行：

```bash
D:/Anaconda/envs/yourenv/python.exe path/to/run_all_with_ocr.py
```

3. 检查是否完整处理：

* 看 `manifest.csv`
  确认：`总文件数 = 成功 + 失败`

4. Gephi 作图：

* Data Laboratory → Import Spreadsheet
* 导入 nodes，再导入 edges
* Layout（ForceAtlas2）
* 节点大小可用 `TotalCount` 或 `DocCount`

---

## 常见问题与排错

### 1) “Tesseract couldn’t load any languages”

通常是 `TESSDATA_PREFIX` 或 `--tessdata-dir` 指向错误，或路径里被错误加了引号。
解决：

* 确认 `TESSDATA_DIR` 指向 **tessdata 文件夹本身**
* 确认里面有 `chi_sim.traineddata`

### 2) Word 报 “远程过程调用失败（RPC）”

这通常是 Word COM 的稳定性问题或残留进程导致。
脚本已做了：RPC 失败 → `taskkill WINWORD.EXE` → 重试一次。
如果仍失败，建议优先依赖 LibreOffice 转换。

### 3) LibreOffice “用户配置锁 / soffice 不可用”

脚本使用了独立 profile（`-env:UserInstallation=...`）避免锁。
如果仍失败，确认 `LIBREOFFICE_PROGRAM_DIR` 是否正确（必须是 `program` 目录）。

---

## 其他说明

* **不依赖 pytesseract**：直接调用 `tesseract.exe` 更可控，避免“路径被错误加引号”等问题
* **先 LibreOffice 后 Word**：LibreOffice headless 更适合批量处理；Word COM 作为备选
* **manifest.csv**：保证每个文件都有结果记录，防止漏文件、漏统计
* **两种边权重**：输出的有binary和weighted的表格，binary 适合看“共现频率”，weighted(min) 更适合看“共现强度”

---


