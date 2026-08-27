# Download Papers From List（从清单下载论文）

这是一个面向 Codex 的论文下载 Skill。它接收已有的 Word、Markdown 或 Excel 论文清单，识别其中的论文题目和下载地址，并按照用户指定的格式标记筛选论文进行下载。下载结束后会生成成功、失败和跳过记录，便于继续补充或人工处理。

本 Skill 只负责下载清单中已经存在的论文，不负责根据题目重新搜索论文。

## 主要功能

- 读取 Word（`.docx`）、Markdown（`.md`）和 Excel（`.xlsx`）论文清单。
- 下载全部论文，或者只下载标题标红、加粗、标红或加粗、同时标红且加粗的论文。
- 支持在格式筛选的基础上继续使用题目正则表达式筛选。
- 优先识别 PDF 下载地址，并可从清单提供的论文详情页中继续识别 `[pdf]`、`Download PDF` 等链接或按钮。
- 支持 CVF/CVPR Open Access 详情页、`citation_pdf_url` 元数据、嵌入式 PDF 查看器和带 URL 属性的下载按钮。
- 最多跟随 3 层详情页，并在后续请求中保留 Cookie 和来源页信息。
- 排除常见的 GitHub、GitLab、Bitbucket 代码仓库地址，避免把代码链接当成论文。
- 默认不覆盖已经下载的同名 PDF。
- 下载后验证 PDF 文件签名，避免把登录页面或错误页面保存成 PDF。
- 输出 Markdown 和 CSV 两份报告，记录成功、失败、跳过及失败原因。

## 安装方法

将仓库克隆到 Codex Skills 目录：

```powershell
git clone https://github.com/hfnjupt/download-papers-from-list.git `
  "$env:USERPROFILE\.codex\skills\download-papers-from-list"
```

如果 Codex 没有立即识别到该 Skill，请重启 Codex。

## 推荐的输入结构

### Word

推荐使用表格，每行对应一篇论文：

| Title | Download URL | Code |
|---|---|---|
| 论文英文题目 | PDF 或论文详情页地址 | 可选代码地址 |

程序读取 `Title` 列的文字颜色和粗体状态。也支持标题位于一个段落、下载地址位于紧随其后的段落。

### Markdown

推荐使用 Markdown 表格：

```markdown
| Title | Download URL | Code |
|---|---|---|
| **需要下载的论文** | https://example.org/paper.pdf | https://github.com/example/code |
| 普通论文 | https://example.org/another-paper.pdf | |
```

`**题目**` 或 `__题目__` 会被识别为加粗题目。使用带有红色样式的 HTML `span` 或 `font` 标签时，也可以识别红色题目。

### Excel

推荐设置以下列：

- `Title`、`论文题目`、`论文标题`、`题目` 或 `标题`
- `Download URL`、`PDF URL`、`下载地址`、`论文链接` 或 `PDF链接`
- 可选的 `Code`、`代码地址` 或 `代码链接`

程序读取题目单元格的字体颜色和粗体状态。默认检查所有可见工作表，也可以指定一个工作表。

## 在 Codex 中使用

可以直接使用 `$download-papers-from-list` 调用。

### 范式一：只下载 Word 中标红的论文

> 使用 `$download-papers-from-list`，读取 `D:\论文清单\papers.docx`，只下载题目标红的论文，保存到 `D:\论文清单\PDF`，并输出下载失败报告。

对应筛选模式：`red`。

### 范式二：只下载 Markdown 中加粗的论文

> 使用 `$download-papers-from-list`，只下载 `papers.md` 中题目加粗的论文。先预演并核对筛选结果，然后下载到 `downloaded-papers` 文件夹。

对应筛选模式：`bold`。

### 范式三：下载 Excel 中所有被突出标记的论文

> 使用 `$download-papers-from-list`，读取 `cross-modal-papers.xlsx`，下载题目标红或加粗的论文，已有文件不要覆盖，并列出失败原因。

对应筛选模式：`marked`。

### 范式四：只下载同时标红且加粗的论文

> 使用 `$download-papers-from-list`，只下载清单中题目同时标红并且加粗的论文。

对应筛选模式：`red-and-bold`。

### 范式五：下载清单中的全部论文

> 使用 `$download-papers-from-list`，下载这个 Excel 清单里的全部论文。不要覆盖已经存在的 PDF，最后输出完整报告。

对应筛选模式：`all`。

### 范式六：格式筛选后再按题目筛选

> 使用 `$download-papers-from-list`，只下载 Word 中标红并且题目包含 `cross-modal` 或 `multimodal` 的论文。

这类请求会把格式条件和题目正则条件组合使用。

## 直接运行下载程序

下载 Word 中标红的论文：

```powershell
python scripts/download_papers_from_list.py "D:\论文清单\papers.docx" `
  --output-dir "D:\论文清单\PDF" `
  --select red
```

下载 Markdown 中加粗的论文：

```powershell
python scripts/download_papers_from_list.py "D:\论文清单\papers.md" `
  --output-dir "D:\论文清单\PDF" `
  --select bold
```

下载 Excel 中标红或加粗的论文，并限定工作表：

```powershell
python scripts/download_papers_from_list.py "D:\论文清单\papers.xlsx" `
  --output-dir "D:\论文清单\PDF" `
  --select marked `
  --sheet "CVPR"
```

先预演，不执行下载：

```powershell
python scripts/download_papers_from_list.py "D:\论文清单\papers.xlsx" `
  --output-dir "D:\论文清单\PDF" `
  --select red `
  --dry-run
```

按题目正则表达式继续筛选：

```powershell
python scripts/download_papers_from_list.py "D:\论文清单\papers.docx" `
  --output-dir "D:\论文清单\PDF" `
  --select red `
  --title-regex "cross[- ]modal|multimodal"
```

授权覆盖已有文件：

```powershell
python scripts/download_papers_from_list.py "D:\论文清单\papers.xlsx" `
  --output-dir "D:\论文清单\PDF" `
  --select all `
  --overwrite
```

## 主要参数

| 参数 | 说明 |
|---|---|
| `--output-dir` | PDF 和报告的输出目录，必填 |
| `--select all` | 下载全部论文 |
| `--select red` | 只下载题目标红的论文 |
| `--select bold` | 只下载题目加粗的论文 |
| `--select marked` | 下载题目标红或加粗的论文 |
| `--select red-and-bold` | 下载题目同时标红且加粗的论文 |
| `--sheet NAME` | 只读取指定的 Excel 工作表 |
| `--title-regex PATTERN` | 使用正则表达式进一步筛选题目 |
| `--dry-run` | 只生成下载计划和报告，不下载 PDF |
| `--overwrite` | 覆盖已经存在的同名 PDF |
| `--workers N` | 并发下载数量，默认值为 4，范围为 1–16 |
| `--timeout N` | 单次网络请求超时秒数，默认值为 30 |
| `--max-mb N` | 单个 PDF 的最大体积，默认值为 200 MB |
| `--report-prefix NAME` | 修改报告文件名前缀 |

## 输出结果

默认在输出目录生成：

```text
PDF输出目录/
├── 论文题目一.pdf
├── 论文题目二.pdf
├── paper_download_report.md
└── paper_download_report.csv
```

报告包含以下状态：

- `success`：PDF 下载并验证成功。
- `failed`：下载地址缺失、访问失败、不是 PDF、详情页没有明确 PDF 地址等。
- `skipped`：目标 PDF 已经存在，并且没有启用覆盖。
- `planned`：预演模式中计划下载的论文。
- `not selected`：不满足本次筛选条件的论文。

## 安全边界

- 只访问公开的 HTTP 或 HTTPS 地址。
- 拒绝本机、局域网、私有网段和保留地址。
- 不绕过登录、验证码或付费墙。
- 不根据题目擅自搜索清单之外的论文。
- 只有验证为 PDF 的响应才会保留为论文文件。

如果网站的下载按钮完全依赖 JavaScript、登录、验证码或人工同意流程，并且页面 HTML 中没有公开下载地址，程序会记录失败原因，不会模拟绕过这些限制。

完整的 Codex 工作流程参见 [SKILL.md](SKILL.md)。
