"""
网页转整页长图
==============
用 Edge headless 把 dashboard/index.html 渲染成竖版长图 PNG，
再用 PIL 裁剪底部空白。输出 dashboard/看板长图.png
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image
import numpy as np

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
HTML_PATH = Path(__file__).resolve().parent / "dashboard" / "index.html"
OUT_PATH = Path(__file__).resolve().parent / "dashboard" / "看板长图.png"
WIDTH = 1240
HEIGHT = 6200  # 先给足高度，稍后裁剪


def render_and_crop() -> Path:
    if not HTML_PATH.exists():
        print(f"❌ 找不到 {HTML_PATH}")
        sys.exit(1)

    # 临时截图放系统 temp 目录（避免沙盒删除限制），用绝对路径
    tmp_png = Path(tempfile.gettempdir()) / "dashboard_fullshot_tmp.png"
    url = HTML_PATH.as_uri()

    cmd = [
        EDGE,
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--force-device-scale-factor=2",  # 2x 高清
        f"--window-size={WIDTH},{HEIGHT}",
        f"--screenshot={tmp_png}",  # 绝对路径
        url,
    ]
    print(f"🖥️  渲染: {url}")
    subprocess.run(cmd, check=True, capture_output=True, timeout=120)

    if not tmp_png.exists():
        print("❌ 截图失败")
        sys.exit(1)

    print("✂️  裁剪底部空白...")
    img = Image.open(tmp_png).convert("RGB")
    w, h = img.size
    arr = np.array(img)
    # 背景色 #f5f5f5 = (245,245,245)
    bg = np.array([245, 245, 245])
    is_bg = np.all(np.abs(arr.astype(int) - bg) < 8, axis=2)  # (h,w)
    row_has_content = ~is_bg.all(axis=1)  # (h,)
    last_content = int(np.max(np.where(row_has_content)[0]))
    crop_bottom = last_content + 30  # 留 30px 余量
    img_cropped = img.crop((0, 0, w, min(crop_bottom, h)))
    img_cropped.save(OUT_PATH)
    print(f"✅ 长图已生成: {OUT_PATH} ({img_cropped.size[0]}x{img_cropped.size[1]})")
    return OUT_PATH


if __name__ == "__main__":
    render_and_crop()
