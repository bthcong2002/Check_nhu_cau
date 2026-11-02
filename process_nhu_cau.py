import gc
import os
import sys
import time
import pandas as pd

# =============================================================================
# CẤU HÌNH ĐƯỜNG DẪN CỐ ĐỊNH
# =============================================================================
INPUT_FILE = r"C:\MWG\Thử nghiệm\Check Nhu cầu\Nhu cầu 31.10\Nhucầumuahàngtheosiêuthịtheongàyđặthàng_9b844ec743c6.xlsx"
OUTPUT_FILE = r"C:\MWG\Thử nghiệm\Check Nhu cầu\Nhu cầu 31.10\nhucau_da_xu_ly.xlsx"

# Đặt True để xuất đủ 4 Sheets (Sheet1, Sheet2, Sheet3, PO bổ sung)
# Đặt False nếu chỉ muốn xuất nhanh 3 Sheets kết quả (Sheet2, Sheet3, PO bổ sung)
INCLUDE_SHEET1 = True


def read_excel_fast(file_path: str) -> pd.DataFrame:
    """
    Ưu tiên đọc file Excel bằng engine 'calamine' (Rust - nhanh gấp 4-5 lần và nhẹ RAM),
    tự động fallback về 'openpyxl' nếu môi trường chưa cài python-calamine.
    """
    try:
        return pd.read_excel(file_path, engine="calamine")
    except Exception:
        return pd.read_excel(file_path)


def stream_df_to_xlsxwriter_sheet(workbook, sheet_name: str, df: pd.DataFrame, pct_col_name: str = None):
    """
    Ghi cuốn chiếu (Streaming) từng cụm 5,000 dòng xuống ổ cứng bằng xlsxwriter constant_memory.
    Giúp quá trình xuất 350,000+ dòng cực nhanh và chỉ tốn < 10 MB RAM.
    """
    ws = workbook.add_worksheet(sheet_name)
    header_fmt = workbook.add_format({"bold": True})
    pct_fmt = workbook.add_format({"num_format": "0.0%"}) if pct_col_name else None

    cols = list(df.columns)
    pct_col_idx = cols.index(pct_col_name) if (pct_col_name and pct_col_name in cols) else None

    if pct_col_idx is not None:
        ws.set_column(pct_col_idx, pct_col_idx, 18, pct_fmt)

    ws.write_row(0, 0, cols, header_fmt)

    chunk_size = 5000
    total = len(df)
    for start in range(0, total, chunk_size):
        chunk = df.iloc[start : start + chunk_size]
        rows = chunk.astype(object).where(pd.notna(chunk), None).values.tolist()
        for offset, row_vals in enumerate(rows):
            r_idx = start + offset + 1
            ws.write_row(r_idx, 0, row_vals)
            if pct_col_idx is not None and row_vals[pct_col_idx] is not None:
                ws.write_number(r_idx, pct_col_idx, float(row_vals[pct_col_idx]), pct_fmt)
        del chunk, rows


def run_process():
    if not os.path.exists(INPUT_FILE):
        print(f"[LỖI] Không tìm thấy file gốc: {INPUT_FILE}")
        return

    print("=" * 80)
    print(" BẮT ĐẦU XỬ LÝ THEO TOÀN BỘ LUỒNG CHUẨN (BẢN TỐI ƯU TỐC ĐỘ & RAM)")
    print(f" - File đầu vào : {INPUT_FILE}")
    print(f" - File xuất ra : {OUTPUT_FILE}")
    print("=" * 80)
    t0 = time.time()

    # =========================================================================
    # BƯỚC 1: Đọc dữ liệu ban đầu & Tạo cột Map
    # =========================================================================
    print("[1/5] Đang đọc file Excel gốc (Engine tốc độ cao)...")
    t_read = time.time()
    df = read_excel_fast(INPUT_FILE)
    total_rows = len(df)
    print(f"      -> Đã đọc {total_rows:,} dòng và {len(df.columns)} cột trong {time.time() - t_read:.1f}s.")

    # Chuẩn hóa dữ liệu chuỗi và số
    df["Khu vực mua hàng"] = df["Khu vực mua hàng"].fillna("").astype(str)
    df["Tên sản phẩm"] = df["Tên sản phẩm"].fillna("").astype(str)
    df["Số lượng cần mua"] = pd.to_numeric(df["Số lượng cần mua"], errors="coerce").fillna(0)
    df["SL mapping"] = pd.to_numeric(df["SL mapping"], errors="coerce").fillna(0)
    df["Khả năng giao NCC"] = pd.to_numeric(df["Khả năng giao NCC"], errors="coerce").fillna(0)

    # Tạo cột Map = Trim(Khu vực mua hàng) - Trim(Tên sản phẩm)
    print("[2/5] Đang tạo cột 'Map' và tính Sheet2 (Pivot Nhu cầu vs Mapping)...")
    col_map = df["Khu vực mua hàng"].str.strip() + "-" + df["Tên sản phẩm"].str.strip()

    if "Map" in df.columns:
        df["Map"] = col_map
    else:
        if "Số lượng cầu ban đầu" in df.columns:
            pos = df.columns.get_loc("Số lượng cầu ban đầu") + 1
            df.insert(pos, "Map", col_map)
        else:
            df["Map"] = col_map
    del col_map

    # =========================================================================
    # BƯỚC 2: Tạo Sheet2 (Pivot Nhu cầu vs SL mapping, Tỉ lệ, NOTE)
    # =========================================================================
    pivot2 = df.groupby(
        ["Khu vực mua hàng", "Tên sản phẩm", "Map"],
        as_index=False,
        dropna=False,
    )[["Số lượng cần mua", "SL mapping"]].sum()

    pivot2.rename(
        columns={
            "Số lượng cần mua": "Sum of Số lượng cần mua",
            "SL mapping": "Sum of SL mapping",
        },
        inplace=True,
    )

    # Tính tỉ lệ mapping/nhu cầu
    req_sum = pivot2["Sum of Số lượng cần mua"]
    map_sum = pivot2["Sum of SL mapping"]
    safe_ratio = map_sum / req_sum.replace(0, float("nan"))
    pivot2["Tỉ lệ mapping/nhu cầu"] = safe_ratio.fillna(0.0)

    # Điều kiện NOTE: nếu chia 0 lỗi -> 'Không'; nếu tỉ lệ < 85% -> 'Có', ngược lại 'Không'
    cond_co = (req_sum > 0) & (pivot2["Tỉ lệ mapping/nhu cầu"] < 0.85)
    pivot2["NOTE"] = cond_co.map({True: "Có", False: "Không"})

    # Map ngược lại cột NOTE vào Sheet1 qua cột Map
    note_dict = dict(zip(pivot2["Map"], pivot2["NOTE"]))
    note_series = df["Map"].map(note_dict).fillna("Không")
    del note_dict, req_sum, map_sum, safe_ratio, cond_co

    if "NOTE" in df.columns:
        df["NOTE"] = note_series
    else:
        note_pos = df.columns.get_loc("Map") + 1
        df.insert(note_pos, "NOTE", note_series)
    del note_series
    sheet2_rows = len(pivot2)

    # =========================================================================
    # BƯỚC 3: Tạo Sheet3 (Chỉ trích đúng 5 cột cần thiết, không nhân bản 48 cột)
    # =========================================================================
    print("[3/5] Đang tạo Sheet3 (Pivot NCC & Lọc năng lực giao cao nhất)...")
    ncc_clean = df["Tên Chi nhánh nhà cung cấp"].fillna("").astype(str).str.strip()
    mask_valid_ncc = (
        (ncc_clean != "")
        & (ncc_clean.str.lower() != "none")
        & (ncc_clean.str.lower() != "nan")
        & (ncc_clean != "(blank)")
    )
    del ncc_clean

    cols_s3 = [
        "Khu vực mua hàng",
        "Tên sản phẩm",
        "Tên Chi nhánh nhà cung cấp",
        "Map",
        "Khả năng giao NCC",
    ]
    pivot3 = (
        df.loc[mask_valid_ncc, cols_s3]
        .groupby(
            ["Khu vực mua hàng", "Tên sản phẩm", "Tên Chi nhánh nhà cung cấp", "Map"],
            as_index=False,
            dropna=False,
        )["Khả năng giao NCC"]
        .max()
    )
    del mask_valid_ncc

    pivot3.rename(columns={"Khả năng giao NCC": "Max of Khả năng giao NCC"}, inplace=True)

    # Bỏ dữ liệu <= 1 (chỉ giữ lại > 1)
    pivot3 = pivot3[pivot3["Max of Khả năng giao NCC"] > 1].copy()

    # Sort theo Map và Max khả năng giao NCC từ cao đến thấp
    pivot3.sort_values(
        by=["Map", "Max of Khả năng giao NCC"],
        ascending=[True, False],
        inplace=True,
    )
    pivot3.reset_index(drop=True, inplace=True)
    sheet3_rows = len(pivot3)
    print(f"      -> Sheet3 có {sheet3_rows:,} dòng NCC hợp lệ sau lọc và sort.")

    # Lấy NCC có năng lực cao nhất cho từng Map (dòng đầu tiên của mỗi Map sau sort giảm dần)
    best_ncc_df = pivot3.drop_duplicates(subset=["Map"], keep="first")
    best_ncc_dict = dict(zip(best_ncc_df["Map"], best_ncc_df["Tên Chi nhánh nhà cung cấp"]))
    del best_ncc_df

    # =========================================================================
    # BƯỚC 4: Tạo Sheet "PO bổ sung"
    # =========================================================================
    print("[4/5] Đang tạo Sheet 'PO bổ sung' và map mới NCC ưu tiên...")
    ma_po_str = df["Mã PO"].fillna("").astype(str).str.strip()
    cond_ma_po_blank = (
        (ma_po_str == "")
        | (ma_po_str.str.lower() == "nan")
        | (ma_po_str.str.lower() == "none")
    )
    del ma_po_str

    cond_po_bo_sung = (df["NOTE"] == "Có") & cond_ma_po_blank
    del cond_ma_po_blank

    df_po_bo_sung = df.loc[cond_po_bo_sung].copy()
    del cond_po_bo_sung

    # Làm trắng sạch cột 'Tên Chi nhánh nhà cung cấp' cũ trước khi map
    df_po_bo_sung["Tên Chi nhánh nhà cung cấp"] = None
    df_po_bo_sung["Tên Chi nhánh nhà cung cấp"] = df_po_bo_sung["Map"].map(best_ncc_dict)
    del best_ncc_dict

    so_dong_po = len(df_po_bo_sung)
    so_dong_co_ncc = int(df_po_bo_sung["Tên Chi nhánh nhà cung cấp"].notna().sum())
    print(f"      -> Tổng dòng PO bổ sung: {so_dong_po:,} dòng.")
    print(f"      -> Số dòng được map NCC mới (năng lực > 1): {so_dong_co_ncc:,} dòng.")
    print(f"      -> Số dòng để trống (không có NCC thỏa > 1): {so_dong_po - so_dong_co_ncc:,} dòng.")

    # =========================================================================
    # BƯỚC 5: Xuất toàn bộ ra file Excel (Streaming constant_memory siêu tốc)
    # =========================================================================
    print(f"[5/5] Đang xuất file kết quả: {OUTPUT_FILE} ...")

    target_file = OUTPUT_FILE
    try:
        if os.path.exists(target_file):
            with open(target_file, "a"):
                pass
    except PermissionError:
        base, ext = os.path.splitext(OUTPUT_FILE)
        target_file = f"{base}_moi{ext}"
        print(f"[CẢNH BÁO] File '{OUTPUT_FILE}' đang mở trong Excel!")
        print(f"            Đang tự động lưu ra file: {target_file}")

    try:
        import xlsxwriter

        wb = xlsxwriter.Workbook(
            target_file,
            {
                "constant_memory": True,
                "strings_to_urls": False,
                "nan_inf_to_errors": True,
                "default_date_format": "yyyy-mm-dd",
            },
        )
        if INCLUDE_SHEET1:
            stream_df_to_xlsxwriter_sheet(wb, "Sheet1", df)
        del df
        gc.collect()

        stream_df_to_xlsxwriter_sheet(wb, "Sheet2", pivot2, pct_col_name="Tỉ lệ mapping/nhu cầu")
        del pivot2
        gc.collect()

        stream_df_to_xlsxwriter_sheet(wb, "Sheet3", pivot3)
        del pivot3
        gc.collect()

        stream_df_to_xlsxwriter_sheet(wb, "PO bổ sung", df_po_bo_sung)
        del df_po_bo_sung
        wb.close()
        gc.collect()
    except ImportError:
        with pd.ExcelWriter(target_file, engine="openpyxl") as writer:
            if INCLUDE_SHEET1:
                df.to_excel(writer, sheet_name="Sheet1", index=False)
            del df
            gc.collect()

            pivot2.to_excel(writer, sheet_name="Sheet2", index=False)
            pivot3.to_excel(writer, sheet_name="Sheet3", index=False)
            df_po_bo_sung.to_excel(writer, sheet_name="PO bổ sung", index=False)

            ws2 = writer.sheets["Sheet2"]
            ratio_col = pivot2.columns.get_loc("Tỉ lệ mapping/nhu cầu") + 1
            for row in range(2, len(pivot2) + 2):
                ws2.cell(row=row, column=ratio_col).number_format = "0.0%"

    total_time = time.time() - t0
    print("=" * 80)
    print(f" XỬ LÝ THÀNH CÔNG TRONG {total_time:.1f} GIÂY!")
    print(f" - File kết quả: {target_file}")
    if INCLUDE_SHEET1:
        print(f"   1. Sheet1      : {total_rows:,} dòng (Dữ liệu gốc + Map + NOTE)")
    print(f"   2. Sheet2      : {sheet2_rows:,} dòng (Pivot Nhu cầu vs SL mapping, Tỉ lệ %, NOTE)")
    print(f"   3. Sheet3      : {sheet3_rows:,} dòng (Pivot NCC > 1, sort giảm dần theo năng lực)")
    print(f"   4. PO bổ sung  : {so_dong_po:,} dòng (Làm trắng NCC cũ -> map NCC mới từ Sheet3)")
    print("=" * 80)


if __name__ == "__main__":
    run_process()
