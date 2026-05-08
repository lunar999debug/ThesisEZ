"""AI 处理器 —— 把用户的「输入.docx」草稿拆解成渲染引擎需要的素材。

可独立运行用于调试 prompt：
    python aichuli.py 输入.docx

会生成 / 覆盖：
    md/201-abstract.md
    md/3xx-*.md       （按一级标题切分，文件名前缀决定章节归属）
    md/304-conclusion.md
    md/401-reference.md
    101-setup.md      （由 info.yaml 渲染）
    image/auto/       （从「图表/」目录搬运改名后的图片/表格）
"""
from __future__ import annotations

import csv
import json
import os
import re
import shutil
import sys
from pathlib import Path

import yaml
from docx import Document
from openai import OpenAI  # DeepSeek 兼容 OpenAI SDK

HERE = Path(__file__).resolve().parent
INPUT_DOCX = HERE / "输入.docx"
INFO_YAML = HERE / "info.yaml"
FIGURE_DIR = HERE / "图表"

MD_DIR = HERE / "md"
IMAGE_DIR = HERE / "image" / "auto"
SETUP_MD = HERE / "101-setup.md"

# ─────────────────────────────────────────────────────────────
# 1. 读 docx，拆段落（保留顺序）
# ─────────────────────────────────────────────────────────────

def read_docx_paragraphs(path: Path) -> list[str]:
    doc = Document(str(path))
    paras = [p.text.strip() for p in doc.paragraphs]
    return [p for p in paras if p]  # 去空行


# ─────────────────────────────────────────────────────────────
# 2. 调 DeepSeek 给每段打标签
# ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """你是一个学位论文结构解析器。用户会给你一篇论文草稿的段落列表（按原顺序），你要给每一段打一个标签，输出严格的 JSON 数组，长度与输入完全相等。

可用标签集合（**只能用这些**）：
- title              ：论文中文标题（通常是第一段）
- abstract_zh_header ：「摘要」二字标题
- abstract_zh        ：中文摘要正文
- keywords_zh        ：「关键词：xxx；xxx」整段
- abstract_en_header ：「ABSTRACT」标题
- abstract_en        ：英文摘要正文
- keywords_en        ：「Keywords: xxx, xxx」整段
- h1                 ：一级标题（识别规则：「一、xxx」「二、xxx」「三、xxx」「四、xxx」「五、xxx」「六、xxx」「七、xxx」「八、xxx」「九、xxx」等中文数字+顿号开头，且整段是纯标题）
- h2                 ：二级标题（识别规则：「1.1 xxx」「2.3xxx」等 N.M 数字开头，且整段是纯标题）
- h3                 ：三级标题（识别规则：「1.1.1 xxx」等 N.M.K 数字开头，且整段是纯标题）

⚠️【关于标题的关键判断】纯标题的特征：长度通常 ≤ 20 字、没有句号、没有冒号后接长正文。
   如果出现「1.1.1 主观饥饿评分：采用 10 分制量表，受试者根据自身感受评分...」这种「数字前缀 + 短词 + 冒号 + 一大段正文」的形态，**整段视为 para**（普通段落），而不是 h3。
   判断口诀：**标题是引出下文的，不是直接讲内容的**。一段话如果自己已经在「讲事情」（出现了句号、长描述），不管前面有没有数字前缀，都标 para。
- para               ：正文普通段落
- figure_ref         ：图占位符（识别规则：单独成段、形如「图3-1」或「图3-1 描述文字」）
- table_ref          ：表占位符（识别规则：单独成段、形如「表3-1」或「表3-1 描述文字」）
- ref_header         ：「参考文献」标题
- ref_item           ：参考文献条目（以 [N] 开头）
- ack_header         ：「致谢」标题
- ack                ：致谢正文
- ignore             ：要丢弃的段（如全空白）

【关键规则】
1. 输出必须是合法 JSON 数组，每个元素形如 {"i": 段落序号, "type": "标签", "clean": "去掉序号后的纯净文本"}
2. clean 字段：
   - h1 去掉「一、」「二、」前缀
   - h2 去掉「1.1 」前缀
   - h3 去掉「1.1.1 」前缀
   - figure_ref / table_ref 保留原文，并额外给出 meta: {"id": "3-1", "caption": "描述文字或空"}
   - ref_item 保留原文（包含 [1] 编号）
   - 其他标签 clean 与原文一致
3. 数组长度必须与输入段落数严格相等，一一对应。
4. 如果用户只写了中文摘要/关键词，没写英文版，**不要**自己生成，让对应位置缺失即可（后续由代码补 placeholder）。
"""

USER_PROMPT_TEMPLATE = """以下是论文草稿的段落列表（每段前面是序号）：

{numbered}

请输出 JSON 数组。"""


def call_deepseek(paragraphs: list[str]) -> list[dict]:
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        sys.exit("❌ 请先设置环境变量 DEEPSEEK_API_KEY")

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    numbered = "\n".join(f"[{i}] {p}" for i, p in enumerate(paragraphs))
    user_msg = USER_PROMPT_TEMPLATE.format(numbered=numbered)

    print(f"→ 调 DeepSeek，共 {len(paragraphs)} 段...")
    resp = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_msg},
        ],
        response_format={"type": "json_object"},
        temperature=0.0,
    )
    raw = resp.choices[0].message.content
    data = json.loads(raw)
    # DeepSeek 在 json_object 模式下经常包一层 {"result": [...]}，做兼容
    if isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list):
                data = v
                break
    if len(data) != len(paragraphs):
        print(f"⚠️  AI 返回 {len(data)} 段，输入 {len(paragraphs)} 段，可能错位，请检查")
    return _sanitize_tags(data, paragraphs)


# “标题+正文”裹在一起的段落拆分正则：
#   「1.1.1 主观饥饿评分：采用 10 分制量表...」
#   「主观饥饿评分：采用 10 分制量表...」
# 只要出现 「照明短词 + ：/: + 长正文」都会识别。
_HEAD_BODY_PATTERN = re.compile(
    r"^(?P<head>[^。、；;，,\n]{2,30})[：:](?P<body>.+)$",
    re.DOTALL,
)


def _split_head_body(clean_text: str):
    """如果 clean_text 是「短标题：长正文」结构，拆为 (head, body) 返回；
    否则返回 None。
    """
    m = _HEAD_BODY_PATTERN.match(clean_text.strip())
    if not m:
        return None
    head = m.group("head").strip()
    body = m.group("body").strip()
    # 标题不能太长，正文不能太短，避免误伤「关键词：xxx」这种。
    if len(head) > 30 or len(body) < 10:
        return None
    return head, body


# 硬规则正则：识别「数字.数字.数字 xxx」开头的段落。例如：
#   "3.3.1主观饥饿评分：采用 10 分制量表..."
#   "3.3.1 主观饥饿评分：采用..."  (数字后有空格)
#   "3.3.1 行为观察"  (纯标题，不带正文)
_H3_HARD_PATTERN = re.compile(r"^\s*(\d+)\.(\d+)\.(\d+)\s*(.*)$", re.DOTALL)
_H2_HARD_PATTERN = re.compile(r"^\s*(\d+)\.(\d+)\s*(.*)$", re.DOTALL)
_H1_HARD_PATTERN = re.compile(r"^\s*([一二三四五六七八九十])、\s*(.*)$", re.DOTALL)


def _hard_detect_heading(raw: str):
    """基于原文正则硬识别标题。返回 (level, head, body) 或 None。
    level 为 'h1'/'h2'/'h3'；body 可能为空字串（纯标题场景）。
    """
    text = raw.strip()
    # 先试 h3（优先级最高，避免 3.3.1 被 h2 吃掉）
    m = _H3_HARD_PATTERN.match(text)
    if m:
        rest = m.group(4).strip()
        head, body = _split_head_or_pure(rest)
        return "h3", head, body
    m = _H2_HARD_PATTERN.match(text)
    if m:
        rest = m.group(3).strip()
        head, body = _split_head_or_pure(rest)
        return "h2", head, body
    m = _H1_HARD_PATTERN.match(text)
    if m:
        rest = m.group(2).strip()
        head, body = _split_head_or_pure(rest)
        return "h1", head, body
    return None


def _split_head_or_pure(text: str):
    """辅助函数：如果 text 形如 「短词：长正文」，拆为 (head, body)；
    否则 (text, '')。
    """
    split = _split_head_body(text)
    if split is not None:
        return split
    return text, ""


def _sanitize_tags(tags: list[dict], paragraphs: list[str]) -> list[dict]:
    """对 AI 输出做后处理，三道防线：
    1. 硬规则优先：任何开头是「N.M.K」「N.M」「一、」的段落，都强制识别为
       对应级别的标题 + 可选正文（必要时拆为两段）。
       这是主防线——不依赖 AI 判断，只看原文结构。
    2. AI 已标为 h2/h3 但仍护袋正文的，拆为两段。
    3. AI 误标为 h2/h3 的超长句子，降级为 para。
    """
    out: list[dict] = []
    hard_count = 0
    split_count = 0
    demote_count = 0

    # 同时需要识别“仅限于正文区”的标题——摘要部分不应被拆（例如
    # 「关键词：xxx」开头不会匹配上面任何一个硬规则，但为防万一加个 section 划分）。
    in_body_zone = False  # 是否已进入正文（遇到首个 h1 后为 True）

    for i, tag in enumerate(tags):
        ttype = tag.get("type", "para")
        raw = paragraphs[i] if i < len(paragraphs) else ""
        clean = tag.get("clean", raw)

        # 一旦进入以后，该 zone 保持 True。
        if ttype in ("h1",) or (in_body_zone is False and ttype == "para" and _H1_HARD_PATTERN.match(raw.strip())):
            in_body_zone = True
        # 摘要/参考文献/致谢 header 遇到时，zone 重置
        if ttype in ("abstract_zh_header", "abstract_en_header", "ref_header", "ack_header"):
            in_body_zone = False

        # 硬规则：只在正文区识别（避免动摘要里的「关键词：xxx」）
        # 主要针对 h2/h3；h1 AI 一般不会错，不动。
        if in_body_zone and ttype in ("para", "h2", "h3"):
            detected = _hard_detect_heading(raw)
            if detected is not None:
                level, head, body = detected
                if level in ("h2", "h3"):
                    new_tag = {"type": level, "clean": head}
                    out.append(new_tag)
                    if body:
                        out.append({"type": "para", "clean": body, "_synthetic": True})
                    if ttype != level:
                        hard_count += 1
                    elif body:
                        split_count += 1
                    continue

        # AI 标为 h2/h3 但拆分、降级逻辑（保留原有）
        if ttype in ("h2", "h3"):
            split = _split_head_body(clean)
            if split is not None:
                head, body = split
                tag["clean"] = head
                out.append(tag)
                out.append({"type": "para", "clean": body, "_synthetic": True})
                split_count += 1
                continue
            if ("。" in raw and len(raw) > 40):
                tag["type"] = "para"
                tag["clean"] = raw
                demote_count += 1
                out.append(tag)
                continue
        out.append(tag)

    if hard_count:
        print(f"→ 硬规则重识别 {hard_count} 个标题（原被 AI 误标为 para）")
    if split_count:
        print(f"→ 后处理：拆分 {split_count} 个「标题+正文」混合段")
    if demote_count:
        print(f"→ 后处理：降级 {demote_count} 个超长误标题为 para")
    return out


# ─────────────────────────────────────────────────────────────
# 3. 把标签结果转成 markdown 文件
# ─────────────────────────────────────────────────────────────

def csv_to_md_table(csv_path: Path) -> str:
    with csv_path.open(encoding="utf-8") as f:
        rows = list(csv.reader(f))
    if not rows:
        return ""
    header = "| " + " | ".join(rows[0]) + " |"
    sep = "| " + " | ".join("---" for _ in rows[0]) + " |"
    body = "\n".join("| " + " | ".join(r) + " |" for r in rows[1:])
    return f"{header}\n{sep}\n{body}"


# 题注中使用的特殊前缀：渲染引擎会识别它，提取出「N-M」作为
# 手动编号，不再插入 SEQ 项，从而题注输出「图 N-M 描述」。
MANUAL_NUM_PREFIX = "@@NUM@@"


def _build_caption_alt(num_id: str, caption: str) -> str:
    """生成题注用的 alt 文本，带上手动编号前缀。
    渲染引擎会检测到该前缀并跳过 SEQ 插入。
    """
    text = caption.strip() if caption else ""
    return f"{MANUAL_NUM_PREFIX}{num_id}|{text}"


def resolve_figure(fig_id: str, caption: str) -> str:
    """图3-1 → ![@@NUM@@3-1|caption](image/auto/fig-3-1.png)
    注意：拷贝时会重命名为纯 ASCII（fig-N-M.ext），避免 Windows 下
    pandoc 读取中文路径出现乱码 / 找不到文件。
    """
    safe_id = fig_id.replace("/", "-").replace("\\", "-")
    alt = _build_caption_alt(fig_id, caption)
    for ext in (".png", ".jpg", ".jpeg", ".pdf"):
        src = FIGURE_DIR / f"图{fig_id}{ext}"
        if src.is_file():
            IMAGE_DIR.mkdir(parents=True, exist_ok=True)
            ascii_name = f"fig-{safe_id}{ext}"
            dst = IMAGE_DIR / ascii_name
            shutil.copy2(src, dst)
            rel = f"image/auto/{ascii_name}"
            return f"![{alt}]({rel})"
    # 文件未找到：不输出 markdown 图片语法（否则 pandoc 会报错），
    # 而是输出一行明文提示，后期用户可以在 Word 里手动插图。
    print(f"⚠️  未找到图{fig_id}.[png|jpg|jpeg|pdf]，已在输出中保留文本占位")
    return f"【图 {fig_id}　{caption or '待插入'}】"


def resolve_table(tbl_id: str, caption: str) -> str:
    safe_id = tbl_id.replace("/", "-").replace("\\", "-")
    alt = _build_caption_alt(tbl_id, caption)
    src = FIGURE_DIR / f"表{tbl_id}.csv"
    if src.is_file():
        # pandoc 语法：以 ": alt" 开头的行会被识别为表题注。
        return f": {alt}\n\n{csv_to_md_table(src)}"
    # csv 文件未找到：不输出表语法，输出明文占位。
    print(f"⚠️  未找到表{tbl_id}.csv，已在输出中保留文本占位")
    return f"【表 {tbl_id}　{caption or '待插入'}】"


def emit_markdown(paragraphs: list[str], tags: list[dict]) -> None:
    """把段落+标签序列写成 md/ 下的多个文件。"""
    MD_DIR.mkdir(exist_ok=True)
    for f in MD_DIR.glob("*.md"):
        f.unlink()

    # ─ 摘要文件：预填「# 摘要」标题，避免被其他文本抢占首行 ─
    abs_lines = ["# 摘要", ""]
    abs_header_emitted = True
    # ─ 正文：按 h1 切分到不同文件 ─
    body_chapters: list[tuple[str, list[str]]] = []  # [(章标题, [行...])]
    cur_chapter: list[str] | None = None
    # ─ 参考文献 / 致谢 ─
    ref_lines = ["# 参考文献", ""]
    ack_lines = ["# 致谢", ""]

    section = "abstract"  # abstract → body → ref → ack
    title = ""

    # 按 tags 遍历（后处理可能会拆出新的 synthetic tag，tags 长度 ≥ paragraphs）
    for idx, tag in enumerate(tags):
        t = tag.get("type", "para")
        # 原始 paragraph 仅侜 clean 缺失时作为 fallback，synthetic tag 只靠 clean。
        fallback = paragraphs[idx] if idx < len(paragraphs) else ""
        clean = tag.get("clean", fallback)
        meta = tag.get("meta", {})

        if t == "title":
            title = clean
            # title 不写入任何 md：模板§1（封面）已包含论文标题
            continue
        elif t == "abstract_zh_header":
            # 首行已预填「# 摘要」，避免重复
            if not abs_header_emitted:
                abs_lines += ["# 摘要", ""]
                abs_header_emitted = True
        elif t == "abstract_zh":
            abs_lines += [clean, ""]
        elif t == "keywords_zh":
            abs_lines += [f"**{clean}**" if not clean.startswith("**") else clean, ""]
        elif t == "abstract_en_header":
            abs_lines += ["# ABSTRACT", ""]
        elif t == "abstract_en":
            abs_lines += [clean, ""]
        elif t == "keywords_en":
            abs_lines += [f"**{clean}**" if not clean.startswith("**") else clean, ""]
        elif t == "h1":
            section = "body"
            cur_chapter = []
            body_chapters.append((clean, cur_chapter))
        elif t == "h2":
            (cur_chapter if cur_chapter is not None else []).extend([f"## {clean}", ""])
        elif t == "h3":
            (cur_chapter if cur_chapter is not None else []).extend([f"### {clean}", ""])
        elif t == "para":
            target = cur_chapter if section == "body" and cur_chapter is not None else abs_lines
            target.extend([clean, ""])
        elif t == "figure_ref":
            line = resolve_figure(meta.get("id", ""), meta.get("caption", ""))
            (cur_chapter if cur_chapter is not None else abs_lines).extend([line, ""])
        elif t == "table_ref":
            block = resolve_table(meta.get("id", ""), meta.get("caption", ""))
            (cur_chapter if cur_chapter is not None else abs_lines).extend([block, ""])
        elif t == "ref_header":
            section = "ref"
        elif t == "ref_item":
            ref_lines.append(clean)
            ref_lines.append("")
        elif t == "ack_header":
            section = "ack"
        elif t == "ack":
            ack_lines.extend([clean, ""])
        # ignore → skip

    # 写文件
    (MD_DIR / "201-abstract.md").write_text("\n".join(abs_lines), encoding="utf-8")

    # 正文章节：301-, 302-, ...
    if not body_chapters:
        body_chapters = [("绪论", ["（草稿无正文）"])]
    for idx, (chap_title, lines) in enumerate(body_chapters, start=1):
        # 最后一章如果叫"总结/结论"，落到 304
        is_last_concl = (idx == len(body_chapters)) and re.search(r"(总结|结论|conclusion)", chap_title, re.I)
        prefix = "304" if is_last_concl else f"30{idx}"
        # 避免和 304 冲突
        if not is_last_concl and prefix == "304":
            prefix = "303"
        fname = f"{prefix}-chapter{idx}.md"
        content = [f"# {chap_title}", ""] + lines
        (MD_DIR / fname).write_text("\n".join(content), encoding="utf-8")

    (MD_DIR / "401-reference.md").write_text("\n".join(ref_lines), encoding="utf-8")
    (MD_DIR / "404-acknowledgement.md").write_text("\n".join(ack_lines), encoding="utf-8")

    print(f"✓ 写入 md/ 完成：{[f.name for f in sorted(MD_DIR.glob('*.md'))]}")


# ─────────────────────────────────────────────────────────────
# 4. 用 info.yaml 生成 101-setup.md
# ─────────────────────────────────────────────────────────────

SETUP_TEMPLATE = """| 需填写字段 | 模板字段内容（用于对照填写，请勿修改） | 填写字段内容（用于填写和修改）   |
| ---------- | ------------ | ---|
| 姓名   | 张三 | {姓名}  |
| Author    | Zhang San | {Author} |
| 学号       | 520XXXXXXXX | {学号} |
| 导师 | 李四 | {导师} |
| Supervisor | Li Si | {Supervisor}|
| 中文学院名 | 机械与动力工程学院 | {中文学院名} |
| 英文学院名 | School of XXXXXXX | {英文学院名} |
| 专业名称 | 工业工程 | {专业名称}|
| 申请学位层次 | 学士 | {申请学位层次} |
| 年份 | 20XX | {年份} |
| 中文论文标题 | 上海交通大学学位论文格式模板 | {中文论文标题} |
| 英文论文标题 | DISSERTATION TEMPLATE FOR BACHELOR DEGREE OF ENGINEERING IN HANGHAI JIAO TONG UNIVERSITY | {英文论文标题} |
"""


def emit_setup_md() -> None:
    info = yaml.safe_load(INFO_YAML.read_text(encoding="utf-8"))
    SETUP_MD.write_text(SETUP_TEMPLATE.format(**info), encoding="utf-8")
    print(f"✓ 写入 {SETUP_MD.name}")


# ─────────────────────────────────────────────────────────────
# 5. 总入口
# ─────────────────────────────────────────────────────────────

def process(input_path: Path = INPUT_DOCX) -> None:
    if not input_path.is_file():
        sys.exit(f"❌ 找不到 {input_path}")
    paragraphs = read_docx_paragraphs(input_path)
    print(f"读到 {len(paragraphs)} 段")
    tags = call_deepseek(paragraphs)
    # 调试落盘
    (HERE / "_ai_tags.json").write_text(
        json.dumps(tags, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("✓ AI 标签已存到 _ai_tags.json（方便排查）")
    emit_markdown(paragraphs, tags)
    emit_setup_md()


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT_DOCX
    process(target)
