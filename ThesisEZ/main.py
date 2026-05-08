"""一键入口：输入.docx → ThesisEZ.docx

用户只需要：
    1. 把自己的论文草稿命名为「输入.docx」放到本目录
    2. 改 info.yaml 里的封面信息
    3. （可选）把图片放到「图表/」目录，命名「图3-1.png」「表3-1.csv」
    4. export DEEPSEEK_API_KEY=sk-xxx
    5. python main.py
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

import aichuli

HERE = Path(__file__).resolve().parent


def clean_artifacts() -> None:
    """清理上次运行的产物，避免脏数据干扰。"""
    for d in ("md", "image"):
        p = HERE / d
        if p.exists():
            shutil.rmtree(p)
    for f in ("101-setup.md", "ThesisEZ.docx", "_ai_tags.json"):
        p = HERE / f
        if p.exists():
            p.unlink()


def main() -> None:
    print("═" * 50)
    print(" ThesisEZ：一键论文格式化")
    print("═" * 50)

    # 1. 清场
    clean_artifacts()

    # 2. AI 处理
    print("\n[1/2] AI 解析草稿...")
    aichuli.process()

    # 3. 渲染
    print("\n[2/2] 调用渲染引擎...")
    # 强制子进程使用 UTF-8，避免 Windows GBK 控制台遇到中文路径/
    # 警告信息时出现乱码（pandoc 报 「锕」之类乱码就是这个问题）。
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    result = subprocess.run(
        [sys.executable, str(HERE / "render.py"), "--no-bibtex"],
        cwd=str(HERE),
        env=env,
    )
    if result.returncode != 0:
        sys.exit("❌ 渲染引擎执行失败")

    out = HERE / "ThesisEZ.docx"
    if out.is_file():
        print(f"\n✅ 完成！输出文件：{out}")
    else:
        sys.exit("❌ 没有生成 ThesisEZ.docx")


if __name__ == "__main__":
    main()
