import os
import sys
import json
from curl_cffi import requests
import pandas as pd
import datetime

# ==========================================
# CẤU HÌNH HỆ THỐNG
# ==========================================
LOGIN_URL = "https://prod-swa-app-be.smartlogix.biz/api/public/auth-portal/login"
DATA_URL = "https://portal-be.sabeco.vn/api/inventories/getListDateShelfLife"

# Đường dẫn đến file Excel đích trên máy của bạn (Tự động nhận theo thư mục chứa script)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TARGET_EXCEL_PATH = os.path.join(SCRIPT_DIR, "Báo cáo Cận date (T2 Hàng tuần).xlsx")
# Tên sheet trong file Excel đích cần dán dữ liệu mới
TARGET_SHEET_NAME = "data"

# Tài khoản đăng nhập Sabeco Portal
USERNAME = "ntb-hoangtien"
PASSWORD = "hehuha170714@"

# Bảng tra cứu tên kho từ mã kho
WHSE_NAME_MAP = {
    "030": "Kho Đà Nẵng",
    "036": "Kho Quảng Ngãi",
    "040": "Kho Đắk Lắk",
    "041": "Kho Gia Lai",
    "050": "Kho Nha Trang",
    "051": "Kho Ninh Thuận",
    "052": "Kho Bình Thuận",
    "05KH": "Kho NTB tại NM BSG Khánh Hòa",
    "05NT": "Kho NTB tại NM BSG Ninh Thuận",
    "151": "NM BSG Phú Yên",
    "157": "NM BSG Quảng Ngãi",
    "S30": "Kho TĐ BSG tại Đà Nẵng",
    "SKH": "Kho TĐ BSG tại NM BSG Khánh Hòa",
    "164": "NM BSG Khánh Hòa",
    "S40": "Kho TĐ BSG tại Daklak",
    "166": "NM BSG Lâm Đồng",
    "149": "NM BSG Đắk Lắk",
    "148": "NM BSG Qui Nhơn",
    "04DL": "Kho TN tại NM BSG Daklak",
    "161": "NM BSG Ninh Thuận",
    "032": "Kho Bình Định"
}

# Bảng tra cứu đơn vị tính (KEG và CUON để trống giống hệt file xuất từ Web)
UOM_MAP = {
    "KET": "Két",
    "CAI": "Cái",
    "THUNG": "Thùng",
    "CHAI": "Chai",
    "KEG": "",
    "CUON": ""
}

def log(msg):
    print(f"[INFO] {msg}", flush=True)

def error_log(msg):
    print(f"[ERROR] {msg}", file=sys.stderr, flush=True)

def format_date(iso_str):
    """Chuyển đổi ngày ISO dạng 2026-03-03T17:00:00.000Z sang định dạng Excel DD/MM/YYYY"""
    if not iso_str:
        return ""
    try:
        parts = iso_str.split("T")[0].split("-")
        if len(parts) == 3:
            return f"{parts[2]}/{parts[1]}/{parts[0]}" # DD/MM/YYYY
    except Exception:
        pass
    return iso_str

def to_number(val):
    """Chuyển đổi về dạng số nguyên hoặc số thực sạch"""
    if val is None or val == "":
        return ""
    try:
        num = float(val)
        return int(num) if num.is_integer() else num
    except Exception:
        return val

def main():
    log("=== KHỞI ĐỘNG LUỒNG TỰ ĐỘNG TẢI DỮ LIỆU HẠN SỬ DỤNG (SABECO) ===")
    
    target_file = os.path.abspath(TARGET_EXCEL_PATH)
    username_val = USERNAME
    password_val = PASSWORD
    
    log("Đang sử dụng thông tin tài khoản đăng nhập Sabeco Portal được cấu hình sẵn...")
            
    # 1. Gửi yêu cầu đăng nhập để lấy Token mới
    log("Đang gửi yêu cầu đăng nhập lấy Token mới...")
    login_payload = {
        "username": username_val,
        "password": password_val
    }
    
    try:
        r = requests.post(LOGIN_URL, json=login_payload, headers={"Content-Type": "application/json"}, timeout=15, impersonate="chrome110")
        r.raise_for_status()
        login_res = r.json()
        token = login_res.get("token") or login_res.get("Token")
        if not token:
            error_log("Không tìm thấy mã Token trong dữ liệu đăng nhập trả về.")
            sys.exit(1)
        log("Đăng nhập thành công, đã nhận được Token.")
    except Exception as e:
        error_log(f"Lỗi đăng nhập thất bại: {e}")
        sys.exit(1)
        
    # 2. Gọi API để tải danh sách dữ liệu hạn sử dụng
    log("Đang gọi API Sabeco tải dữ liệu hạn sử dụng...")
    data_headers = {
        "Content-Type": "application/json",
        "Token": token,
        "Whseid": "148",
        "Origin": "https://portal.sabeco.vn",
        "Referer": "https://portal.sabeco.vn/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.0.0 Safari/537.36"
    }
    
    data_payload = {
        "whereClause": "1>0",
        "orderBy": "",
        "skip": 0,
        "limit": 100000
    }
    
    try:
        r = requests.post(DATA_URL, json=data_payload, headers=data_headers, timeout=30, impersonate="chrome110")
        r.raise_for_status()
        response_json = r.json()
        
        # Trích xuất mảng danh sách từ JSON
        raw_items = []
        if isinstance(response_json, list):
            raw_items = response_json
        elif isinstance(response_json, dict):
            raw_items = response_json.get("res") or []
            
            if not raw_items:
                data_val = response_json.get("data")
                if isinstance(data_val, list):
                    raw_items = data_val
                elif isinstance(data_val, dict):
                    raw_items = data_val.get("list") or data_val.get("rows") or data_val.get("data") or []
                
            if not raw_items:
                raw_items = response_json.get("result") or response_json.get("items") or response_json.get("rows") or []
                if isinstance(raw_items, dict):
                    raw_items = raw_items.get("list") or raw_items.get("rows") or []
            
        if not raw_items:
            error_log("Dữ liệu trống hoặc API trả về sai định dạng danh sách.")
            sys.exit(1)
            
        log(f"Đã tải về thành công {len(raw_items)} dòng dữ liệu từ Sabeco Portal.")
    except Exception as e:
        error_log(f"Lỗi tải dữ liệu thất bại: {e}")
        sys.exit(1)

    # 3. Ánh xạ trực tiếp các trường API vào đúng 23 cột đầu tiên (A đến W)
    # Không ánh xạ 2 cột cuối X và Y vì đây là cột công thức tính toán của Excel Table
    mapped_data = []
    for item in raw_items:
        # Tên đơn vị
        branch = str(item.get("branchname", "")).strip()
        
        # Lấy mã kho và tên kho
        ma_kho = str(item.get("whseid", "")).strip()
        
        # LỌC KHO THEO YÊU CẦU DỰ ÁN
        if ma_kho not in ["052", "05NT", "05KH", "SKH"]:
            continue
            
        whse_name_raw = str(item.get("whsename", "")).strip()
        ten_kho = WHSE_NAME_MAP.get(ma_kho, whse_name_raw)

        # Định dạng đơn vị tính
        uom_raw = str(item.get("uom", "")).strip()
        dvt = UOM_MAP.get(uom_raw.upper(), uom_raw)

        # Trạng thái
        status_raw = str(item.get("status", "")).strip().upper()
        trang_thai = "CÓ THỂ XUẤT" if status_raw == "OK" else status_raw
        
        # LỌC TRẠNG THÁI THEO Slicer Cũ
        if trang_thai not in ["CÓ THỂ XUẤT", "D"]:
            continue

        # Ngày sản xuất và hạn sử dụng
        nsx = format_date(item.get("lottable04", ""))
        hsd = format_date(item.get("lottable05", ""))

        # Tính phần trăm HSD còn lại quy đổi
        pct_raw = item.get("percentshl", 0)
        pct_val = 0
        try:
            pct_val = float(pct_raw)
        except Exception:
            pct_val = pct_raw

        row = {
            "TÊN ĐƠN VỊ": branch,
            "TÊN KHO": ten_kho,
            "MÃ KHO": ma_kho,
            "MÃ HÀNG": str(item.get("sku", "")).strip(),
            "TÊN HÀNG": str(item.get("skuname", "")).strip(),
            "ĐƠN VỊ TÍNH": dvt,
            "NHÓM HÀNG": str(item.get("packtype", "")).strip(),
            "TRẠNG THÁI": trang_thai,
            "SỐ LÔ": str(item.get("lottable06", "")).strip(),
            "KMDB": str(item.get("lottable01", "")).strip(),
            "NSX": format_date(item.get("lottable04", "")),
            "HSD": hsd,
            "SỐ LƯỢNG": to_number(item.get("qty", 0)),
            "SỐ LƯỢNG PL": to_number(item.get("qtyavailablepl", 0)),
            "(%) HSD": pct_val,
            "SỐ NGÀY CÒN LẠI": to_number(item.get("dayoff", 0)),
            "SỐ NGÀY HSD": to_number(item.get("shelflife", 0)),
            "(%) KHẢ DỤNG": to_number(item.get("usable", 0)),
            "(%) GẦN HẾT HẠN": to_number(item.get("nearlyexpired", 0)),
            "(%) HẾT HẠN": to_number(item.get("expired", 0)),
            "TRẠNG THÁI HSD": str(item.get("statusshelflife", "")).strip() if item.get("statusshelflife") is not None else "",
            "STOCKDATE": to_number(item.get("stockdate", "")),
            "VỊ TRÍ": str(item.get("location", "")).strip()
        }
        
        # LOGIC LỌC CẬN DATE THEO FORMULA CŨ CỦA EXCEL
        # =_xlfn.IFS(AND(LEFT(C2,1)="S",O2<=75),"cận date",O2<=50,"cận date",O2>50,"")
        is_candate = False
        if ma_kho.startswith("S") and pct_val <= 75:
            is_candate = True
        elif pct_val <= 50:
            is_candate = True
            
        if is_candate:
            mapped_data.append(row)
        
    # 4. Ghi dữ liệu vào file data.js dưới dạng hằng số mảng
    log("Đang ghi dữ liệu vào tệp data.js...")
    try:
        data_js_path = os.path.join(SCRIPT_DIR, "data.js")
        now_str = datetime.datetime.now().strftime('%H:%M:%S %d/%m/%Y')
        import random
        random_id = random.randint(1000, 9999)
        with open(data_js_path, "w", encoding="utf-8") as f:
            f.write(f"const DATA_UPDATED_TIME = '{now_str}';\n")
            f.write(f"const DATA_RUN_ID = '{random_id}';\n")
            f.write("const DATA_CANDATE = ")
            json.dump(mapped_data, f, ensure_ascii=False, indent=2)
            f.write(";\n")
        log(f"Đã ghi thành công {len(mapped_data)} dòng dữ liệu vào {data_js_path}")
    except Exception as e:
        error_log(f"Lỗi khi ghi tệp data.js: {e}")
        sys.exit(1)
        
    log("=== LUỒNG TỰ ĐỘNG HOÀN THÀNH THÀNH CÔNG ===")

import traceback
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = traceback.format_exc()
        with open("data.js", "w", encoding="utf-8") as f:
            f.write(f"var error_log = {json.dumps(error_msg)};\n")
        print(error_msg)
        sys.exit(0)
