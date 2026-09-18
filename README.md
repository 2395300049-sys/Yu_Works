# Yu_Works 2.0 — 基础排版版

Yu_Works 将 `.md`、`.txt` 或 `.docx` 整理为统一格式的 Word 文档，并保留可编辑公式、图片和表格等基础能力。

本版本专注基础排版，不包含内容创作、学校模板克隆、复杂模式切换和 Web 服务。

## 最简单的使用方法

1. 选择或拖入 `.md`、`.txt`、`.docx` 文件。
2. 标题格式不需要调整时，直接保留默认值。
3. 点击“开始基础排版”。
4. 在 Word 或 WPS 中检查结果并自动生成目录。

## 标题默认值

| 级别 | 中文字体 | 字号 | 段前 | 段后 | 对齐 | 段前分页 |
|---|---|---:|---:|---:|---|---|
| 一级标题 | 黑体 | 小三 | 40磅 | 20磅 | 居中 | 是 |
| 二级标题 | 黑体 | 四号 | 24磅 | 6磅 | 左对齐 | 否 |
| 三级标题 | 黑体 | 小四 | 12磅 | 6磅 | 左对齐 | 否 |

界面允许分别修改一级至三级标题的中文字体、字号、段前、段后和是否居中。一级标题始终启用段前分页。

## 目录规则

Yu_Works 不生成、不重排目录。检测到输入 Word 文档中的目录区域时会跳过该区域；Markdown 或纯文本中的“目录”章节也会跳过。请在输出文档中通过 Word/WPS 的“引用 → 目录”自动生成。

## 本地运行

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python gui_main.py
```

Windows 激活虚拟环境：

```bat
.venv\Scripts\activate
```

## 打包

Windows：

```bash
python -m PyInstaller --noconfirm --clean Yu_Works.spec
```

macOS：

```bash
python -m PyInstaller --noconfirm --clean Yu_Works-mac.spec
```

## 验证

```bash
python -m unittest discover -s tests -v
```

## 许可

[MIT](LICENSE)
