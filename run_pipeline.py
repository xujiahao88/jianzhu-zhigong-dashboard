"""
完整流水线：更新数据 → 出图 → 发微信文件传输助手
==================================================

用法：
    python run_pipeline.py <源Excel路径>              # 完整流程
    python run_pipeline.py <源Excel路径> --dry-run   # 只更新+出图，不发微信
    python run_pipeline.py <源Excel路径> --send      # 强制发（跳过确认）
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

WORK_DIR = Path(__file__).resolve().parent
PYTHON = Path(r"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe")


def run_step(name: str, cmd: list[str], cwd: Path = WORK_DIR):
    print(f"\n{'='*60}\n🚀 {name}\n{'='*60}")
    env = os.environ.copy()
    env["PYTHONPATH"] = str(WORK_DIR)
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(cmd, cwd=cwd, env=env)
    if result.returncode != 0:
        print(f"❌ {name} 失败 (exit {result.returncode})")
        sys.exit(result.returncode)
    print(f"✅ {name} 完成")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("src", help="源 Excel 路径")
    ap.add_argument("--dry-run", action="store_true", help="只更新+出图，不发微信")
    ap.add_argument("--send", action="store_true", help="强制发微信（跳过 dry-run 默认）")
    ap.add_argument("--years", type=int, default=5, help="看板图表保留近 N 年")
    ap.add_argument("--recent-days", type=int, default=10, help="表格显示近 N 天")
    args = ap.parse_args()

    src = Path(args.src)
    if not src.exists():
        print(f"❌ 找不到源文件: {src}")
        sys.exit(1)

    # 1) 更新数据
    run_step(
        "更新数据 → 钢材直供量统计.xlsx",
        [str(PYTHON), str(WORK_DIR / "update_zhigong.py"), str(src)],
    )

    # 2) 重建 dashboard
    run_step(
        "重建看板（图表 + 表格）",
        [str(PYTHON), str(WORK_DIR / "build_dashboard.py"),
         "--years", str(args.years),
         "--recent-days", str(args.recent_days)],
    )

    # 3) 网页转长图
    run_step(
        "网页转整页长图",
        [str(PYTHON), str(WORK_DIR / "html_to_image.py")],
    )

    # 4) 发微信
    if args.dry_run and not args.send:
        print("\n⏭️  跳过微信发送（--dry-run 模式）")
        print(f"   长图已生成: {WORK_DIR / 'dashboard' / '看板长图.png'}")
        print("   手动预览/发送，或重新跑加 --send")
        return

    run_step(
        "发送长图到微信文件传输助手",
        [str(PYTHON), str(WORK_DIR / "send_to_wechat.py")],
    )

    print("\n" + "="*60)
    print("🎉 完整流水线跑完")
    print(f"   长图: {WORK_DIR / 'dashboard' / '看板长图.png'}")
    print("="*60)


if __name__ == "__main__":
    main()