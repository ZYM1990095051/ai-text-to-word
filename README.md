<div align="center">

# AI Text To Word

**把 AI 生成的含公式文本，整理成更适合 WPS / Word 的 Markdown 和 DOCX。**

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)
![Pandoc](https://img.shields.io/badge/Pandoc-Required-1A1A1A?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?style=flat-square&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)

</div>

## Why This Project

很多 AI 生成的讲义、题解、学习资料都混着中文、LaTeX 公式和化学公式。  
这些内容直接复制到 WPS / Word 后，常见问题是：

- 公式分隔符原样暴露，阅读体验很差
- 数学公式和化学公式混排时容易乱
- `\ce{H+}`、`\ce{OH-}` 这类化学式经常退化成普通文本
- 手动逐个“插入公式 -> 转为专业格式”很费时间

这个项目的目标就是把这件事自动化。

## What It Does

项目当前支持三段式流程：

```text
TXT  ->  Markdown  ->  DOCX
```

也支持直接一键：

```text
TXT  ->  DOCX
```

核心能力：

- 自动识别常见文本编码，如 UTF-8、UTF-16、GBK、GB18030
- 规范化 AI 文本里的数学公式分隔符
- 将 Markdown 数学公式导出为 Office 公式对象
- 对常见化学公式写法做预处理，减少导出后乱码或退化
- 支持单文件和整目录批量转换

## Quick Start

> 如果你只想直接把 `.txt` 变成 Word，优先用这一条：

```bash
python txt_to_word.py
```

默认流程：

1. 从 `原始文本/` 读取 `.txt`
2. 生成中间 Markdown 到 `output_md/`
3. 生成最终 Word 文档到 `output_docx/`

### 分步运行

先做 `txt -> md`：

```bash
python txt_to_md.py
```

再做 `md -> docx`：

```bash
python md_to_docx.py
```

### 指定单个文件

```bash
python txt_to_word.py "原始文本\01.txt" --md-output "output_md\01.md" --docx-output "output_docx\01.docx"
```

### 使用参考模板

```bash
python txt_to_word.py "原始文本\01.txt" --reference-doc "your-template.docx"
```

## Requirements

- Python 3.10+
- `pandoc`

先确认 `pandoc` 已安装：

```bash
pandoc --version
```

## Project Structure

```text
AI文本转换/
├─ txt_to_md.py       # TXT -> Markdown
├─ md_to_docx.py      # Markdown -> DOCX
├─ txt_to_word.py     # 一键 TXT -> DOCX
├─ 原始文本/           # 输入文本
├─ output_md/         # 中间 Markdown 输出
├─ output_docx/       # 最终 DOCX 输出
├─ README.md
├─ LICENSE
└─ .gitignore
```

## Scripts

| Script | Purpose |
| --- | --- |
| `txt_to_md.py` | 读取原始 `.txt`，统一编码、换行和公式分隔符，输出 `.md` |
| `md_to_docx.py` | 调用 `pandoc` 把 Markdown 转成 `.docx`，尽量保留公式为 Office 公式对象 |
| `txt_to_word.py` | 串联前两步，适合日常直接使用 |

## Example

仓库里已经包含一组示例文件：

- 输入样例：[原始文本/01.txt](./原始文本/01.txt)
- Markdown 输出示例：[output_md/01.md](./output_md/01.md)
- DOCX 输出示例：`output_docx/01.docx`

## Formula Handling

这个项目目前主要处理两类内容：

- 数学公式：把 `\(...\)`、`\[...\]` 等形式整理成更标准的 Markdown 数学格式
- 化学公式：对 `\ce{...}` 这类常见写法做预处理，提升导出到 Word / WPS 时的可读性

最终的 DOCX 生成依赖 `pandoc`。  
这意味着它不是简单把公式当普通文本复制进去，而是尽量生成 Word 能识别的公式内容。

## Current Limits

- 优先覆盖常见数学公式和常见化学公式
- 极少数冷门 LaTeX 宏可能还需要额外规则
- WPS 和 Word 的显示效果通常接近，但个别场景仍可能有差异

## Roadmap

- 更强的化学公式兼容
- 更完善的 DOCX 样式控制
- 图形界面
- 拖拽式批量转换
- 打包成 Windows 可执行程序

## License

MIT
