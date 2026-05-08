# ThesisEZ：一键把"裸奔"草稿变成标准格式毕设

> AI + Pandoc 双引擎，让用户**不用学 Markdown**、**不用拆文件**，写完 Word 草稿就能一键交差。

## 它解决什么问题

学位论文的"内容"和"格式"是两件事，但学校通常只发一份格式要求 + 一份 Word 模板，让你自己去对齐。多数同学就是在最后一周对着字号、行距、目录域、题注编号反复磨。

ThesisEZ 的思路：
- 你写**只管内容**的草稿（字号字体随便）
- AI 解析草稿结构，输出标准 Markdown
- 渲染引擎用 Word 模板套出符合规范的 docx

---

## 上手前要装的东西

### 1. Python 3.9 或更高

终端跑 `python --version`，如果版本号不对或者提示找不到命令，去 [python.org](https://www.python.org/downloads/) 下个 3.9+。**Windows 安装时务必勾上 "Add Python to PATH"**，不然后面所有命令都用不了。

### 2. Pandoc

> **不要 `pip install pandoc`！** 那个包只是个 Python 包装器，不是 pandoc 本体，装完照样会报"找不到 pandoc"。

正确方式二选一：

**方式 A：conda（推荐）**
```bash
conda install -c conda-forge pandoc
```

**方式 B：直接下安装包**
- Windows / macOS：去 [pandoc.org/installing.html](https://pandoc.org/installing.html) 下安装包，双击装
- macOS 用 Homebrew：`brew install pandoc`
- Linux：`sudo apt install pandoc`（Ubuntu/Debian）

装完后终端跑 `pandoc --version`，能看到版本号就成了。

### 3. Python 依赖包

```bash
pip install -r requirements.txt
```

这一行装的是 `python-docx`、`openai`、`PyYAML` 三个库。

### 4. DeepSeek API Key

去 [platform.deepseek.com](https://platform.deepseek.com/) 注册，充几块钱（按token量估计运行一次两万字论文需要差不多0.1元或更低），拿到一个 `sk-` 开头的 key。然后在终端里设置环境变量：

**Windows（CMD）：**
```cmd
set DEEPSEEK_API_KEY=sk-你的key
```

**Windows（PowerShell）：**
```powershell
$env:DEEPSEEK_API_KEY="sk-你的key"
```

**macOS / Linux：**
```bash
export DEEPSEEK_API_KEY=sk-你的key
```

> 注意：上面这种设法**只在当前终端窗口有效**。关掉终端就没了，下次跑要重新 set 一遍。想永久生效请使用"环境变量永久设置"。

---

## 用户要做的事（5 步）

1. 下载本项目
2. 把自己的草稿改名 `输入.docx` 放到根目录（**不需要任何格式**，只要满足下面的弱约定）
3. 编辑 `info.yaml` 填封面信息（姓名、学号、导师、题目）
4. （有图就放）在根目录建 `图表/` 文件夹，把图片命名成 `图3-1.png`、表格数据存成 `表3-1.csv`（很遗憾，目前没法支持公式，需后续自己手动加入）
5. 终端跑：
   ```bash
   python main.py
   ```
6. 拿到 `ThesisEZ.docx`，Word 打开，**弹"是否更新域"点是**（这一步是为了刷新目录和题注编号）。完事。

---

## 草稿的弱约定

只要你大致这样写，AI 都能识别（**格式不用对齐，字号字体随便**）：

```
不吃饭会很饿       ← 第一段当标题

摘要
xxxxxxx 中文摘要 xxxxx
关键词：A；B；C

ABSTRACT
xxxxxxx english abstract xxxxx
Keywords: A, B, C

一、绪论                                  ← 一级标题：中文数字+顿号
1.1 研究背景                              ← 二级标题：N.M
1.1.1 国内现状                            ← 三级标题：N.M.K
正文段落 ...
图3-1 实验流程图                          ← 图占位
表3-1 数据汇总                            ← 表占位

二、相关工作
...

参考文献
[1] 张三. 论文标题 [J]. ...
[2] ...

致谢
感谢 xxx
```

---

## 项目结构

| 文件 | 作用 |
| --- | --- |
| `main.py` | 一键入口，串起 AI 解析 + 渲染 |
| `aichuli.py` | AI 解析器：把"输入.docx"拆成结构化 Markdown + 题注 + 章节 |
| `render.py` | 内置渲染引擎：基于 Pandoc 把 Markdown + Word 模板合成 docx，再做后处理（三线表、题注、分节符、自动目录、公式编号……） |
| `info.yaml` | 封面信息（姓名、学号、导师、题目等） |
| `template/` | Word 模板（默认放的是 SJTU 的，**换成任何学校的 .docx 模板都行**） |
| `图表/` | 用户手动放图片和数据，命名 `图3-1.png` / `表3-1.csv` |

---

## 换成自己学校的模板

`template/` 里默认放了 SJTU 的中文模板（仅作示例）。换法：
1. 把你学校官方发的 `.docx` 模板放到 `template/` 下
2. 在 `render.py` 里把 `TEMPLATE_ZH`、`TEMPLATE_EN` 路径改成你的文件名即可
3. 模板里需要保留 `Heading 1/2/3`、`图`、`表` 等标准样式名，渲染引擎会按这些样式套版

---

## 常见报错速查

| 报错信息 | 原因 | 解决 |
| --- | --- | --- |
| `pandoc: command not found` / `'pandoc' 不是内部或外部命令` | pandoc 没装或没进 PATH | 按上面"第 2 步"用 conda 或安装包重装 |
| `ModuleNotFoundError: No module named 'docx'` | 装错包了，装成了 Py2 老库 | `pip uninstall docx` 然后 `pip install python-docx` |
| `Could not fetch resource image/auto/鍥?` | Windows 中文路径 GBK 乱码 | 已自动处理，确认图片放在 `图表/` 而不是其它地方 |
| `DEEPSEEK_API_KEY not set` | 环境变量没 set 或者重开了终端 | 重新 `set` / `export` 一遍 |
| Word 打开后目录是空的 / 题注编号没出来 | 没点"更新域" | 关掉重开，弹窗点"是"；或者 Ctrl+A 全选 → F9 |

---

## 致谢

- 引擎：Pandoc + python-docx + DeepSeek


`template/` 里默认放了 SJTU 的中文模板（仅作示例）。换法：
1. 把你学校官方发的 `.docx` 模板放到 `template/` 下
2. 在 `render.py` 里把 `TEMPLATE_ZH`、`TEMPLATE_EN` 路径改成你的文件名即可
3. 模板里需要保留 `Heading 1/2/3`、`图`、`表`、`公式` 等标准样式名，渲染引擎会按这些样式套版

## 致谢

- 引擎：Pandoc + python-docx + DeepSeek
