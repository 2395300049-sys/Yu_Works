# Yu_Works — 自动化排版与可编辑公式解析器

一键将 Markdown / 纯文本 / Word 文档转换为严格排版的 `.docx`，支持 LaTeX 数学公式转 Word 原生公式。

## 直接下载

在 [Releases](https://github.com/2395300049-sys/Yu_Works/releases/latest) 下载已打包版本。当前发布包为 macOS Apple Silicon 版；源码可用于 Windows、macOS 和 Linux 自行构建。

macOS 首次打开若被 Gatekeeper 拦截，请在 Finder 中右键应用并选择“打开”。发布包未经 Apple 公证。

## 功能

- **多格式输入**：`.md` `.txt` `.docx` 均可作为输入
- **Markdown 解析**：`#` `##` `###` 标题、`` ``` `` 代码块自动识别
- **纯文本智能识别**：`第X章` `1.` `1.1` `一、` `摘要` 等编号自动转为标题层级
- **LaTeX → Word 公式**：`$\frac{a}{b}$` `\sqrt{x}` `\int_0^1` 转为 Word 原生 OMML 公式（根号带横线、分数有分数线、积分有上下限）
- **中英混排**：中文宋体 + 英文/数字 Times New Roman 自动分流
- **三线表**：`.docx` 重格式化时保留表格内容，并统一为顶线、表头线和底线的学术三线表
- **图片保留**：`.docx` 中的图片/Visio 绘图等深拷贝保留
- **自动编号提取**：Word 自动编号（`1. / 1.1 / 1.1.1`）自动解析并转换为文本
- **封面预留**：Markdown 输入自动生成空白封面页

## 排版规范

| 元素 | 字体 | 字号 | 加粗 | 对齐 |
|------|------|------|:--:|:--:|
| 一级标题 (`#` / `第X章`) | 宋体 | 16pt (三号) | ✅ | `第X章` 居中 |
| 二级标题 (`##` / `1.1`) | 宋体 | 14pt (四号) | ✅ | 左对齐 |
| 三级标题 (`###` / `1.1.1`) | 宋体 | 12pt (小四) | ✅ | 左对齐 |
| 正文 | 宋体 + TNR | 12pt (小四) | — | 两端对齐 |
| 代码块 | Times New Roman | 11pt | — | 左对齐 |
| 表格 | — | 10.5pt (五号) | — | 上下左右居中 |

- 行距：固定 20 磅
- 正文首行缩进：2 字符（0.85cm）
- 一级/二级标题段前段后：0.5 行

## 安装

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
python -m pip install -r requirements.txt
```

## 使用

### GUI 桌面应用

```bash
python gui_main.py
```

macOS 可使用项目内的专用环境启动：

```bash
./launch_yu_works_macos.sh
```

macOS 独立应用的重建命令：

```bash
.venv-macos/bin/python -m PyInstaller --noconfirm --clean Yu_Works-mac.spec
```

`launch_yu_works.bat` 和 `Yu_Works.lnk` 仅供 Windows 使用；macOS 请使用
`launch_yu_works_macos.sh` 或桌面的 `Yu_Works.app`。

- 📂 文件上传 — 选择 `.md` / `.docx` 文件
- ✏️ 文本输入 — 直接粘贴 Markdown 内容
- 输出模式 — 导出 `.docx` 或直接打开预览

### 命令行

```bash
python format_conversion.py input.md output.docx
python format_conversion.py input.docx output.docx
```

不带参数则弹出文件选择框；拖拽文件到 `Yu_Works.exe` 上自动生成 `*_Ultra.docx`。

### Web 服务

```bash
pip install flask
python web_app.py
```

手机同 WiFi 下浏览器访问 `http://电脑IP:5050`，Safari/Chrome 添加到主屏幕即可像 App 使用（PWA）。

## Markdown 语法

```markdown
# 一级标题
## 二级标题
### 三级标题

正文段落。公式 $\int_0^1 x dx$ 转 Word 原生公式。

​```python
print('代码块')
​```
```

纯文本自动识别（即使没有 `#` 标记）：

| 输入 | 识别为 |
|------|--------|
| `第一章 概述` | 一级标题（居中 + 新页） |
| `1. 项目背景` | 一级标题 |
| `1.1 市场分析` | 二级标题 |
| `1.1.1 数据来源` | 三级标题 |
| `一、研究意义` | 二级标题 |
| `摘要` / `Abstract` | 一级标题（居中 + 新页） |

## LaTeX 公式支持

公式用 `$...$`（行内）或 `$$...$$`（块级）包裹：

- **分数**：`\frac{a}{b}` → (a)/(b)
- **根号**：`\sqrt{x}` → √(x)
- **积分**：`\int_0^1` → ∫₀¹
- **求和**：`\sum_{i=1}^n` → Σⁿᵢ₌₁
- **希腊字母**：`\alpha \beta \gamma \pi \theta \Delta \Sigma`
- **函数**：`\sin \cos \tan \arcsin \log \ln \lim`
- **运算符**：`\times \cdot \div \pm \oplus \otimes`
- **关系符**：`\approx \neq \leq \geq \equiv \to \implies`
- **重音**：`\bar{A} \hat{x} \tilde{y} \vec{v} \dot{x}`
- **括号**：`\left( \right)` — 自动适应高度

所有公式均可选转为 **Word 原生公式**（OMML）——根号带横线、分数有分数线、积分有上下限，渲染效果与 Word 公式编辑器一致。

## AI 内容清洗管线

针对 DeepSeek / 豆包 / Kimi / ChatGPT 等 AI 平台输出，引擎内置多道清洗：

- **强力反转义**：修复 AI 过度转义的 Markdown（`\*\*bold\*\*` → `**bold**`、`1\.` → `1.`、`\[链接\]` → `[链接]`）
- **DeepSeek 思考剥离**：自动移除 `<think>...</think>` 标签
- **豆包公式修复**：`\$\$` 转义还原、`\[...\]` 统一为 `$$...$$`
- **重影去重**：消除网页剪贴板造成的长重复短语（`velocityvelocity` → `velocity`）
- **零宽字符清洗**：移除不可见 Unicode 干扰字符
- **纯文本 `**` 解析**：`.md` 模式下的 `**加粗**` 自动转为 Word 原生加粗

## 关键设计

| 机制 | 说明 |
|---|---|
| **四级加粗回退** | 直接格式 → 字符样式（Strong）→ 段落样式 → 默认 |
| **冒号雷达** | 全段加粗时以冒号为界，前半保留加粗、后半降级细体 |
| **XML 自动编号探测** | 读取 `w:numPr` 节点识别 Word 原生列表 |
| **表格 sectPr 前置插入** | 表格插入到 `<w:sectPr>` 之前，解决表格沉底 |
| **所见即所得** | 禁用 Word 原生自动编号追加，标题前缀不切除 |
| **延迟导入** | `table_processor.py` 函数内 `import`，规避循环依赖 |

## 打包为 .exe

项目使用 PyInstaller spec 文件管理打包配置，自动收集所有依赖：

```bash
# GUI 桌面版（无控制台窗口，29MB）
pyinstaller Yu_Works.spec

# 命令行版（带控制台，15.7MB）
pyinstaller Yu_Works-CLI.spec

# Web 服务器版（带控制台显示访问地址，15.3MB）
pyinstaller WebApp.spec
```

每个 spec 自动打包 `format_conversion.py` / `latex_to_omml.py` / `table_processor.py` / `Montserrat-Bold.ttf` / `3.ico`，GUI 版额外收集 `customtkinter` + `tkinterdnd2` 的二进制数据。

## 项目结构

```
format_conversion.py   — 核心排版引擎（Markdown 解析 + .docx 重格式化 + AI 清洗管线）
latex_to_omml.py       — LaTeX → Word OMML 公式转换器
table_processor.py     — Markdown 表格渲染插件（延迟导入，零循环依赖）
gui_main.py            — 桌面 GUI（CustomTkinter + tkinterdnd2 拖拽）
web_app.py             — Web 服务（Flask + PWA）
Yu_Works.spec           — GUI 版 PyInstaller 打包配置
Yu_Works-CLI.spec       — CLI 版 PyInstaller 打包配置
WebApp.spec            — Web 版 PyInstaller 打包配置
3.ico                  — 应用图标
Montserrat-Bold.ttf    — 英文标题字体
```

## 许可

[MIT](LICENSE)
