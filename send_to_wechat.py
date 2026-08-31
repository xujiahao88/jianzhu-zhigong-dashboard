"""
发送长图到微信「文件传输助手」
================================
⚠️ 前置条件：
    1. 微信 PC 版（新版 4.0 / Weixin.exe）已登录
    2. 用户已手动切换到「文件传输助手」会话窗口
    3. 微信窗口在前台
    4. 微信窗口在屏幕上固定位置（避免被遮挡）

流程：
    1. 把长图 PNG 复制到 Windows 剪贴板
    2. 激活微信窗口
    3. Ctrl+V 粘贴图片
    4. 间隔 0.5 秒
    5. Enter 发送

为什么这样：
    - 微信 4.0 是 Qt 重写的客户端，控件树对 Windows 自动化完全不可见
      （pywinauto 只能看到 2 个空 Pane）
    - 拖拽图片需要精确定位聊天输入框，容易失败
    - 用剪贴板 + 键盘快捷键最稳
"""
from __future__ import annotations

import io
import sys
import time
from io import BytesIO
from pathlib import Path

import win32clipboard
from PIL import Image
import pywinauto
import pyautogui


IMG_PATH = Path(__file__).resolve().parent / "dashboard" / "看板长图.png"
WECHAT_WINDOW_TITLE = "微信"  # 微信 4.0 主窗口标题


def image_to_clipboard(img_path: Path):
    """把图片放到 Windows 剪贴板（CF_DIB 格式）。"""
    img = Image.open(img_path).convert("RGB")
    output = BytesIO()
    img.save(output, "BMP")
    data = output.getvalue()[14:]  # 去掉 BMP 文件头（14 字节），只留 DIB
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()
    print(f"📋 图片已放入剪贴板 ({img.size[0]}x{img.size[1]})")


def activate_wechat():
    """激活微信窗口到前台。"""
    desktop = pywinauto.Desktop(backend="uia")
    for w in desktop.windows():
        if w.window_text() == WECHAT_WINDOW_TITLE:
            w.set_focus()
            print(f"🪟 已激活微信窗口 ({w.rectangle()})")
            return w
    print("⚠️ 没找到微信窗口，请确认已登录并打开")
    return None


def main():
    if not IMG_PATH.exists():
        print(f"❌ 找不到图片: {IMG_PATH}")
        print("   请先运行 html_to_image.py 生成看板长图")
        sys.exit(1)

    dry_run = "--dry-run" in sys.argv
    if dry_run:
        print("🧪 DRY-RUN 模式：只复制图片到剪贴板，不发送")

    print("⏰ 5 秒后开始...")
    print("   请确保：① 微信已登录  ② 文件传输助手会话已打开  ③ 微信窗口在前台")
    for i in range(5, 0, -1):
        print(f"   {i}...", end="\r", flush=True)
        time.sleep(1)
    print()

    image_to_clipboard(IMG_PATH)

    win = activate_wechat()
    if win is None:
        sys.exit(1)

    if dry_run:
        print("✅ Dry-run 完成：图片已复制到剪贴板，微信窗口已激活")
        return

    time.sleep(0.5)
    print("📎 Ctrl+V 粘贴图片...")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.0)
    print("📤 Enter 发送...")
    pyautogui.press("enter")
    print("✅ 发送完成")


if __name__ == "__main__":
    main()