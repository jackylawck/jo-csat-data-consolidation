# main.py
import logging
import os
import sys

try:
    import config
except ImportError:
    raise ImportError("❌ 找不到 config.py，請確認專案目錄完整！")

from processor import generate_enterprise_report


def main():
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
    )

    if len(sys.argv) < 2:
        input_path = input(
            "請輸入或直接拖曳 Microsoft Forms RAW Excel 檔案至此："
        ).strip("\"'")
    else:
        input_path = sys.argv[1].strip("\"'")

    if not os.path.exists(input_path):
        print(f"❌ 找不到檔案：{input_path}")
        input("按 Enter 鍵離開...")
        return

    print("🚀 正在執行問卷資料整合與報表生成...")
    try:
        with open(input_path, "rb") as f:
            raw_bytes = f.read()

        output_io = generate_enterprise_report(raw_bytes, config)

        output_filename = "2025年客戶滿意度問卷調查資料整合_已生成.xlsx"
        output_dir = os.path.dirname(input_path) or "."
        output_path = os.path.join(output_dir, output_filename)

        with open(output_path, "wb") as f:
            f.write(output_io.getvalue())

        print("=" * 60)
        print("🎉 整合成功！")
        print(f"📁 報表已輸出至：{output_path}")
        print("=" * 60)
    except Exception as e:
        print(f"❌ 執行失敗，錯誤詳情：{str(e)}")

    input("\n按 Enter 鍵離開...")


if __name__ == "__main__":
    main()
