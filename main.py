import os
import sys
import shutil
import logging
import traceback
from modules.dieu_chinh_kho import DieuChinhKho

# Cấu hình logging
logging.basicConfig(
    filename="error.log",  # Ghi lỗi vào file error.log
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def extract_resources():
    """Giải nén các file cần thiết từ file .exe vào thư mục cố định."""
    # Lấy đường dẫn thư mục chứa main.py hoặc file .exe
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    resources_dir = os.path.join(base_dir, "resources")
    chromium_dir = os.path.join(resources_dir, "chromium-1169") if getattr(sys, 'frozen', False) else os.path.join(base_dir, "chromium-1169")

    # Kiểm tra nếu thư mục chromium-1169 chưa tồn tại
    if not os.path.exists(chromium_dir):
        if getattr(sys, 'frozen', False):
            # Chế độ đóng gói (PyInstaller)
            print("Đang giải nén Chromium...")
            source_chromium_dir = os.path.join(sys._MEIPASS, "chromium-1169")
            shutil.copytree(source_chromium_dir, chromium_dir)
            print(f"Đã giải nén Chromium vào: {chromium_dir}")
        else:
            # Chế độ phát triển
            print(f"Thư mục Chromium không tồn tại trong chế độ phát triển: {chromium_dir}")
            raise FileNotFoundError(f"Thư mục Chromium không tồn tại: {chromium_dir}")
    else:
        print("Thư mục Chromium đã tồn tại. Không cần giải nén lại.")

def main():
    try:
        # Giải nén các file cần thiết (nếu cần)
        extract_resources()

        while True:
            print("\nMenu chính:")
            print("1. Điều chỉnh kho")
            print("0. Thoát")

            choice = input("Chọn một tùy chọn: ")

            if choice == "1":
                dieu_chinh_kho_menu()
            elif choice == "0":
                print("Thoát chương trình.")
                break
            else:
                print("Tùy chọn không hợp lệ. Vui lòng chọn lại.")
    except Exception as e:
        # Ghi lỗi vào file log
        logging.error("Lỗi xảy ra: %s", traceback.format_exc())
        print("Đã xảy ra lỗi. Vui lòng kiểm tra file error.log để biết thêm chi tiết.")

def dieu_chinh_kho_menu():
    dieu_chinh_kho = DieuChinhKho()

    while True:
        print("\nMenu Điều chỉnh kho:")
        print("1. Tạo file DCT/DCG")
        print("2. Import file DCT lên web")
        print("3. Import file DCG lên web")
        print("0. Quay lại menu chính")

        choice = input("Chọn một tùy chọn: ")

        if choice == "1":
            dieu_chinh_kho.process_and_create_files()
        elif choice == "2":
            dieu_chinh_kho.import_dct_to_web()
        elif choice == "3":
            dieu_chinh_kho.import_dcg_to_web()
        elif choice == "0":
            print("Quay lại menu chính.")
            break
        else:
            print("Tùy chọn không hợp lệ. Vui lòng chọn lại.")

if __name__ == "__main__":
    main()
