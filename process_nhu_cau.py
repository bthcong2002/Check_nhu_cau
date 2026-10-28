import os
import sys
import time
import pandas as pd
import openpyxl

# =============================================================================
# CẤU HÌNH ĐƯỜNG DẪN CỐ ĐỊNH
# =============================================================================
INPUT_FILE = r"C:\MWG\Thử nghiệm\Check Nhu cầu\Nhu cầu 27.10\Nhucầumuahàngtheosiêuthịtheongàyđặthàng_93c16d50ddac.xlsx"
OUTPUT_FILE = r"C:\MWG\Thử nghiệm\Check Nhu cầu\Nhu cầu 27.10\nhucau_da_xu_ly.xlsx"

def run_process():
    if not os.path.exists(INPUT_FILE):
        print(f"[LỖI] Không tìm thấy file gốc: {INPUT_FILE}")
        return

    print("=" * 80)
    print(" BẮT ĐẦU XỬ LÝ THEO TOÀN BỘ LUỒNG CHUẨN")
    print(f" - File đầu vào : {INPUT_FILE}")
    print(f" - File xuất ra : {OUTPUT_FILE}")
    print("=" * 80)
    t0 = time.time()

    # =========================================================================
    # BƯỚC 1: Đọc dữ liệu ban đầu & Tạo cột Map
    # =========================================================================
    print("[1/5] Đang đọc file Excel gốc...")
    df = pd.read_excel(INPUT_FILE)
    print(f"      -> Đã đọc {len(df):,} dòng và {len(df.columns)} cột.")

    # Chuẩn hóa dữ liệu chuỗi và số
    df['Khu vực mua hàng'] = df['Khu vực mua hàng'].fillna('').astype(str)
    df['Tên sản phẩm'] = df['Tên sản phẩm'].fillna('').astype(str)
    df['Số lượng cần mua'] = pd.to_numeric(df['Số lượng cần mua'], errors='coerce').fillna(0)
    df['SL mapping'] = pd.to_numeric(df['SL mapping'], errors='coerce').fillna(0)
    df['Khả năng giao NCC'] = pd.to_numeric(df['Khả năng giao NCC'], errors='coerce').fillna(0)

    # Tạo cột Map = Trim(Khu vực mua hàng) - Trim(Tên sản phẩm)
    print("[2/5] Đang tạo cột 'Map' và tính Sheet2 (Pivot Nhu cầu vs Mapping)...")
    col_map = df['Khu vực mua hàng'].str.strip() + '-' + df['Tên sản phẩm'].str.strip()
    
    if 'Map' in df.columns:
        df['Map'] = col_map
    else:
        if 'Số lượng cầu ban đầu' in df.columns:
            pos = df.columns.get_loc('Số lượng cầu ban đầu') + 1
            df.insert(pos, 'Map', col_map)
        else:
            df['Map'] = col_map

    # =========================================================================
    # BƯỚC 2: Tạo Sheet2 (Pivot Nhu cầu vs SL mapping, Tỉ lệ, NOTE)
    # =========================================================================
    pivot2 = df.groupby(
        ['Khu vực mua hàng', 'Tên sản phẩm', 'Map'],
        as_index=False,
        dropna=False
    )[['Số lượng cần mua', 'SL mapping']].sum()

    pivot2.rename(columns={
        'Số lượng cần mua': 'Sum of Số lượng cần mua',
        'SL mapping': 'Sum of SL mapping'
    }, inplace=True)

    # Tính tỉ lệ mapping/nhu cầu
    req_sum = pivot2['Sum of Số lượng cần mua']
    map_sum = pivot2['Sum of SL mapping']
    safe_ratio = map_sum / req_sum.replace(0, float('nan'))
    pivot2['Tỉ lệ mapping/nhu cầu'] = safe_ratio.fillna(0.0)

    # Điều kiện NOTE: nếu chia 0 lỗi -> 'Không'; nếu tỉ lệ < 85% -> 'Có', ngược lại 'Không'
    cond_co = (req_sum > 0) & (pivot2['Tỉ lệ mapping/nhu cầu'] < 0.85)
    pivot2['NOTE'] = cond_co.map({True: 'Có', False: 'Không'})

    # Map ngược lại cột NOTE vào Sheet1 qua cột Map
    note_dict = dict(zip(pivot2['Map'], pivot2['NOTE']))
    note_series = df['Map'].map(note_dict).fillna('Không')

    if 'NOTE' in df.columns:
        df['NOTE'] = note_series
    else:
        note_pos = df.columns.get_loc('Map') + 1
        df.insert(note_pos, 'NOTE', note_series)

    # =========================================================================
    # BƯỚC 3: Tạo Sheet3 (Pivot NCC: Rows: Khu vực, SP, NCC, Map | Values: Max Khả năng giao)
    # Lọc: bỏ blank ở cột NCC, bỏ Khả năng giao <= 1 (chỉ lấy > 1)
    # Sort: theo Map và Max khả năng giao giảm dần (ưu tiên NCC năng lực cao nhất)
    # =========================================================================
    print("[3/5] Đang tạo Sheet3 (Pivot NCC & Lọc năng lực giao cao nhất)...")
    ncc_clean = df['Tên Chi nhánh nhà cung cấp'].fillna('').astype(str).str.strip()
    df_valid_ncc = df[
        (ncc_clean != '') & 
        (ncc_clean.str.lower() != 'none') & 
        (ncc_clean.str.lower() != 'nan') &
        (ncc_clean != '(blank)')
    ].copy()

    pivot3 = df_valid_ncc.groupby(
        ['Khu vực mua hàng', 'Tên sản phẩm', 'Tên Chi nhánh nhà cung cấp', 'Map'],
        as_index=False,
        dropna=False
    )['Khả năng giao NCC'].max()

    pivot3.rename(columns={'Khả năng giao NCC': 'Max of Khả năng giao NCC'}, inplace=True)

    # Bỏ dữ liệu <= 1 (chỉ giữ lại > 1)
    pivot3 = pivot3[pivot3['Max of Khả năng giao NCC'] > 1].copy()

    # Sort theo Map và Max khả năng giao NCC từ cao đến thấp
    pivot3.sort_values(
        by=['Map', 'Max of Khả năng giao NCC'],
        ascending=[True, False],
        inplace=True
    )
    pivot3.reset_index(drop=True, inplace=True)
    print(f"      -> Sheet3 có {len(pivot3):,} dòng NCC hợp lệ sau lọc và sort.")

    # Lấy NCC có năng lực cao nhất cho từng Map (dòng đầu tiên của mỗi Map sau sort giảm dần)
    best_ncc_df = pivot3.drop_duplicates(subset=['Map'], keep='first')
    best_ncc_dict = dict(zip(best_ncc_df['Map'], best_ncc_df['Tên Chi nhánh nhà cung cấp']))

    # =========================================================================
    # BƯỚC 4: Tạo Sheet "PO bổ sung"
    # - Lọc Sheet1: NOTE = "Có" và "Mã PO" trống/blank
    # - Làm trắng hoàn toàn cột "Tên Chi nhánh nhà cung cấp" cũ
    # - Map NCC mới từ Sheet3 (ưu tiên năng lực cao nhất)
    # =========================================================================
    print("[4/5] Đang tạo Sheet 'PO bổ sung' và map mới NCC ưu tiên...")
    ma_po_str = df['Mã PO'].fillna('').astype(str).str.strip()
    cond_ma_po_blank = (ma_po_str == '') | (ma_po_str.str.lower() == 'nan') | (ma_po_str.str.lower() == 'none')
    
    cond_po_bo_sung = (df['NOTE'] == 'Có') & cond_ma_po_blank
    df_po_bo_sung = df[cond_po_bo_sung].copy()

    # QUAN TRỌNG: Làm trắng sạch cột 'Tên Chi nhánh nhà cung cấp' cũ trước khi map
    df_po_bo_sung['Tên Chi nhánh nhà cung cấp'] = None

    # Map mới Tên Chi nhánh NCC từ Sheet3 theo cột Map
    # Nếu Map nào không có NCC > 1 ở Sheet3 thì để trống (None), không giữ NCC cũ
    df_po_bo_sung['Tên Chi nhánh nhà cung cấp'] = df_po_bo_sung['Map'].map(best_ncc_dict)

    so_dong_po = len(df_po_bo_sung)
    so_dong_co_ncc = df_po_bo_sung['Tên Chi nhánh nhà cung cấp'].notna().sum()
    print(f"      -> Tổng dòng PO bổ sung: {so_dong_po:,} dòng.")
    print(f"      -> Số dòng được map NCC mới (năng lực > 1): {so_dong_co_ncc:,} dòng.")
    print(f"      -> Số dòng để trống (không có NCC thỏa > 1): {so_dong_po - so_dong_co_ncc:,} dòng.")

    # =========================================================================
    # BƯỚC 5: Xuất toàn bộ ra file Excel
    # =========================================================================
    print(f"[5/5] Đang xuất file kết quả: {OUTPUT_FILE} ...")
    
    # Kiểm tra xem file có đang bị mở bởi Excel không
    target_file = OUTPUT_FILE
    try:
        # Thử mở file ghi kiểm tra lock
        if os.path.exists(target_file):
            with open(target_file, 'a'):
                pass
    except PermissionError:
        target_file = r"C:\MWG\CHECK PO\nhucau_da_xu_ly_moi.xlsx"
        print(f"[CẢNH BÁO] File '{OUTPUT_FILE}' đang mở trong Excel!")
        print(f"            Đang tự động lưu ra file: {target_file}")

    with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        pivot2.to_excel(writer, sheet_name='Sheet2', index=False)
        pivot3.to_excel(writer, sheet_name='Sheet3', index=False)
        df_po_bo_sung.to_excel(writer, sheet_name='PO bổ sung', index=False)

        # Định dạng % cho cột tỉ lệ ở Sheet2
        ws2 = writer.sheets['Sheet2']
        ratio_col = pivot2.columns.get_loc('Tỉ lệ mapping/nhu cầu') + 1
        for row in range(2, len(pivot2) + 2):
            ws2.cell(row=row, column=ratio_col).number_format = '0.0%'

    total_time = time.time() - t0
    print("=" * 80)
    print(f" XỬ LÝ THÀNH CÔNG TRONG {total_time:.1f} GIÂY!")
    print(f" - File kết quả: {target_file}")
    print(f"   1. Sheet1      : {len(df):,} dòng (Dữ liệu gốc + Map + NOTE)")
    print(f"   2. Sheet2      : {len(pivot2):,} dòng (Pivot Nhu cầu vs SL mapping, Tỉ lệ %, NOTE)")
    print(f"   3. Sheet3      : {len(pivot3):,} dòng (Pivot NCC > 1, sort giảm dần theo năng lực)")
    print(f"   4. PO bổ sung  : {len(df_po_bo_sung):,} dòng (Làm trắng NCC cũ -> map NCC mới từ Sheet3)")
    print("=" * 80)

if __name__ == '__main__':
    run_process()
