# AI Text To Word

把 AI 生成的含公式文本整理成 Markdown 和 Word 文档，尽量减少复制到 WPS / Word 后出现的公式乱码问题。

这个项目现在支持三步：

1. `txt -> md`
2. `md -> docx`
3. `txt -> docx` 一键转换

## 这个项目解决什么问题

很多 AI 生成的学习资料、讲义、题解都会夹杂 LaTeX 公式或化学公式。  
直接复制进 WPS / Word 时，经常会出现下面几类问题：

- 中文和公式混在一起时格式很乱
- `\(...\)`、`\[...\]` 这类公式分隔符不能直接好看地显示
- 化学公式如 `\ce{H+}`、`\ce{OH-}` 导出后容易变成普通文本
- 手动在 WPS 里逐个“插入公式 -> 转为专业格式”很费时间

这个仓库的思路是：

- 先把原始 TXT 统一整理成更稳定的 Markdown
- 再把 Markdown 通过 `pandoc` 转成 DOCX
- 在导出 DOCX 前，先预处理一部分化学公式写法，让 Word / WPS 更容易识别成公式对象

## 先看这些文件

你现在只要先认识下面这些就够了：

- `README.md`
  这是项目说明书。以后别人点进你的 GitHub，第一个看的就是它。
- `txt_to_md.py`
  把原始 `.txt` 转成 `.md`。
- `md_to_docx.py`
  把 `.md` 转成 `.docx`。
- `txt_to_word.py`
  一键执行完整流程，最适合你平时直接用。
- `原始文本/`
  放你准备转换的 `.txt` 文件。
- `output_md/`
  程序自动生成的 Markdown 输出目录。
- `output_docx/`
  程序自动生成的 Word 输出目录。
- `.gitignore`
  告诉 Git 哪些文件不要上传，比如缓存文件、输出文件、你自己的临时笔记。
- `LICENSE`
  开源许可证，表示别人可以在什么规则下使用你的项目。
- `.git/`
  Git 自己的内部记录目录，不用手动改它。

## 哪些文件你平时最常碰

你以后大多数时候只需要看这 4 个地方：

- `README.md`
- `txt_to_word.py`
- `原始文本/`
- `output_docx/`

如果你只是想“把文本转成 Word”，最简单的理解方式就是：

1. 把 `.txt` 放进 `原始文本/`
2. 运行 `python txt_to_word.py`
3. 去 `output_docx/` 找结果

## 当前功能

### `txt_to_md.py`

- 自动尝试 UTF-8、UTF-16、GBK、GB18030 等常见编码
- 统一换行和常见不可见字符
- 把 `\(...\)` 转成 `$...$`
- 把 `\[...\]` 转成 `$$...$$`
- 支持单文件和整目录批量转换

### `md_to_docx.py`

- 调用 `pandoc` 生成真正的 `.docx`
- 把 Markdown 数学公式导出为 Office 公式对象
- 对 `\ce{...}` 这类化学公式做预处理
- 支持单文件和整目录批量转换
- 支持 `--reference-doc` 复用你的 DOCX 样式模板

### `txt_to_word.py`

- 直接把 `txt` 一键转成 `docx`
- 中间会自动生成 Markdown
- 适合日常直接出 Word 文档

## 环境要求

- Python 3.10+
- 已安装 `pandoc`

先在终端里确认：

```bash
pandoc --version
```

## 快速开始

### 1. 一键 TXT 转 DOCX

```bash
python txt_to_word.py
```

默认流程：

- 从 `原始文本` 读取 `.txt`
- 生成到 `output_md`
- 再输出到 `output_docx`

### 2. 只做 TXT 转 Markdown

```bash
python txt_to_md.py
```

### 3. 只做 Markdown 转 DOCX

```bash
python md_to_docx.py
```

### 4. 指定单个文件

```bash
python txt_to_word.py "原始文本\01.txt" --md-output "output_md\01.md" --docx-output "output_docx\01.docx"
```

### 5. 使用 Word / WPS 样式模板

```bash
python txt_to_word.py "原始文本\01.txt" --reference-doc "your-template.docx"
```

## 示例

当前仓库里已经放了一个示例流程：

- 输入样例：`原始文本/01.txt`
- Markdown 输出：`output_md/01.md`
- DOCX 输出：`output_docx/01.docx`

## 怎么上传到 GitHub

下面是最适合新手的一套流程。

### 第一步：去 GitHub 新建一个空仓库

在 GitHub 网页上：

1. 登录你的 GitHub 账号
2. 点右上角 `+`
3. 点 `New repository`
4. 仓库名可以填：`ai-text-to-word`
5. 先不要勾选 `Add a README file`
6. 点 `Create repository`

创建完后，GitHub 会给你一个仓库地址，格式大概像这样：

```bash
https://github.com/你的用户名/ai-text-to-word.git
```

### 第二步：如果 Git 还没设置名字和邮箱，先设置一次

如果你是第一次用 Git，在终端里执行：

```bash
git config --global user.name "你的名字"
git config --global user.email "你的邮箱"
```

### 第三步：在当前项目目录执行下面这些命令

```bash
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/你的用户名/ai-text-to-word.git
git push -u origin main
```

### 第四步：刷新 GitHub 页面

如果上传成功，你会在 GitHub 页面看到这些内容：

- `README.md`
- 3 个 Python 脚本
- `原始文本/` 示例文件
- `LICENSE`

而下面这些不会被上传，这是正常的：

- `output_md/`
- `output_docx/`
- `__pycache__/`
- `txt.txt`

因为 `.gitignore` 已经帮你排除了它们。

## 已知限制

- 目前优先处理常见数学公式和常见化学公式
- 如果文本里用了非常冷门的 LaTeX 宏，可能还需要继续补规则
- WPS 对公式兼容通常不错，但和 Word 最终显示仍可能有少量差异

## 后面还能继续加什么

- 图形界面
- 拖拽文件转换
- 批量处理整个资料目录
- 更强的化学公式/数学公式兼容
- 直接打包成 Windows 可执行程序

## License

MIT
