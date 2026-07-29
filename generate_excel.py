"""Windows 免安装 Excel 生成器入口。"""

from __future__ import annotations

import sys
import traceback

from scripts.excel_builder import build_workbooks


def main() -> int:
    try:
        build_workbooks()
        print("\n生成成功：请查看 output 文件夹。")
        return 0
    except Exception:
        print("\n生成失败，错误信息如下：\n")
        traceback.print_exc()
        return 1
    finally:
        if "--no-pause" not in sys.argv:
            input("\n按回车键关闭窗口……")


if __name__ == "__main__":
    raise SystemExit(main())
