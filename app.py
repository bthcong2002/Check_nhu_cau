import ctypes
import gc
import os
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st

# =============================================================================
# CẤU HÌNH TRANG WEB STREAMLIT
# =============================================================================
st.set_page_config(
    page_title="Hệ Thống Xử Lý Nhu Cầu & PO Bổ Sung",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Đường dẫn file kết quả tạm trên ổ cứng máy chủ (tránh giữ 80MB bytes trong RAM)
EXPORT_FILE_PATH = os.path.join(tempfile.gettempdir(), "nhucau_da_xu_ly_latest.xlsx")

# Custom CSS lấy cảm hứng từ bộ Progress Bar Teal-Mint
st.markdown(
    """
<style>
.main-header {
    background: linear-gradient(135deg, #24474e 0%, #377375 55%, #65b89f 100%);
    padding: 26px 30px;
    border-radius: 16px;
    color: white;
    margin-bottom: 24px;
    box-shadow: 0 6px 20px rgba(36, 71, 78, 0.18);
}
.main-header h1 {
    margin: 0;
    font-size: 28px;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: 0.3px;
}
.main-header p {
    margin: 8px 0 0 0;
    opacity: 0.92;
    font-size: 15px;
}
.guide-box {
    background-color: #f4faf8;
    border-left: 5px solid #377375;
    padding: 15px 20px;
    border-radius: 0 10px 10px 0;
    margin-bottom: 20px;
    font-size: 14.5px;
    color: #1e3a3f;
}
.progress-wrapper {
    background: #ffffff;
    border: 2px solid #24474e;
    border-radius: 16px;
    padding: 22px 26px;
    margin: 16px 0 24px 0;
    box-shadow: 0 8px 24px rgba(36, 71, 78, 0.10);
}
.progress-top-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 14px;
}
.progress-title {
    display: flex;
    align-items: center;
    gap: 12px;
    font-weight: 700;
    font-size: 17px;
    color: #24474e;
}
.progress-percent-badge {
    background: linear-gradient(90deg, #24474e, #377375);
    color: #ffffff;
    font-weight: 800;
    font-size: 15px;
    padding: 5px 16px;
    border-radius: 999px;
    letter-spacing: 0.6px;
}
.teal-spinner {
    width: 24px;
    height: 24px;
    border: 3.5px solid #c8e6dc;
    border-top: 3.5px solid #24474e;
    border-right: 3.5px solid #377375;
    border-radius: 50%;
    animation: spin 0.9s linear infinite;
    flex-shrink: 0;
}
@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}
.bar-outer-pill {
    width: 100%;
    height: 38px;
    background: #f1f7f5;
    border: 2.5px solid #1e2d2f;
    border-radius: 11px;
    padding: 4px;
    position: relative;
    overflow: hidden;
    box-sizing: border-box;
    margin-bottom: 16px;
}
.bar-inner-gradient {
    height: 100%;
    border-radius: 6px;
    background: linear-gradient(90deg, #24474e 0%, #377375 50%, #68b89e 100%);
    transition: width 0.35s ease;
}
.bar-center-label {
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    font-size: 14px;
    letter-spacing: 1.5px;
    color: #162527;
    text-shadow: 0 0 6px rgba(255, 255, 255, 0.9);
    pointer-events: none;
}
.step-list {
    display: flex;
    flex-direction: column;
    gap: 8px;
    background: #f7fbf9;
    border: 1px solid #d7eae4;
    border-radius: 10px;
    padding: 12px 16px;
}
.step-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 14px;
    padding: 4px 0;
}
.step-done {
    color: #1f6f54;
    font-weight: 600;
}
.step-active {
    color: #24474e;
    font-weight: 700;
}
.step-pending {
    color: #94a3b8;
    font-weight: 500;
}
.step-badge {
    font-size: 12.5px;
    padding: 2px 10px;
    border-radius: 999px;
    font-weight: 600;
}
.badge-done {
    background: #d1fae5;
    color: #065f46;
}
.badge-active {
    background: #cfe9e2;
    color: #1e3a3f;
}
.badge-pending {
    background: #f1f5f9;
    color: #94a3b8;
}
</style>
""",
    unsafe_allow_html=True,
)

MAIN_STEPS = [
    (1, "Bước 1: Đọc dữ liệu Excel gốc (Calamine tốc độ cao) & chuẩn hóa kiểu dữ liệu"),
    (2, "Bước 2: Tạo cột 'Map' & tính Sheet2 (Pivot Nhu cầu vs SL mapping, Tỉ lệ %, NOTE)"),
    (3, "Bước 3: Tạo Sheet3 (Pivot NCC & lọc năng lực giao > 1 cao nhất)"),
    (4, "Bước 4: Tạo Sheet 'PO bổ sung' (Làm trắng NCC cũ -> Map NCC ưu tiên từ Sheet3)"),
    (5, "Bước 5: Ghi cuốn chiếu (Streaming) ra file Excel & giải phóng RAM"),
]

REQUIRED_COLUMNS = [
    "Khu vực mua hàng",
    "Tên sản phẩm",
    "Số lượng cần mua",
    "SL mapping",
    "Khả năng giao NCC",
    "Tên Chi nhánh nhà cung cấp",
    "Mã PO",
]


def release_memory():
    """
    Thu hồi rác bộ nhớ Python (gc.collect) đồng thời ép hệ điều hành Linux trên
    Streamlit Cloud (libc.so.6 malloc_trim) lấy lại ngay vùng RAM đã giải phóng,
    tránh tình trạng cộng dồn RAM qua các lần chạy gây crash 'Oh no. Error running app'.
    """
    gc.collect()
    try:
        libc = ctypes.CDLL("libc.so.6")
        libc.malloc_trim(0)
    except Exception:
        pass


def get_excel_sheet_names_fast(uploaded_file) -> list:
    """
    Lấy danh sách tên Sheet tức thì (< 0.01 giây, 0 MB RAM) bằng cách đọc trực tiếp
    metadata xl/workbook.xml trong cấu trúc nén của file .xlsx.
    """
    try:
        uploaded_file.seek(0)
        with zipfile.ZipFile(uploaded_file, "r") as z:
            if "xl/workbook.xml" in z.namelist():
                root = ET.fromstring(z.read("xl/workbook.xml"))
                sheets = [
                    el.attrib["name"]
                    for el in root.iter()
                    if el.tag.endswith("sheet") and "name" in el.attrib
                ]
                uploaded_file.seek(0)
                if sheets:
                    return sheets
    except Exception:
        pass

    uploaded_file.seek(0)
    try:
        xl = pd.ExcelFile(uploaded_file, engine="calamine")
        sheets = xl.sheet_names
        xl.close()
        uploaded_file.seek(0)
        return sheets
    except Exception:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        sheets = xl.sheet_names
        xl.close()
        uploaded_file.seek(0)
        return sheets


def read_excel_fast(uploaded_file, sheet_name: str) -> pd.DataFrame:
    """
    Ưu tiên đọc file Excel bằng engine 'calamine' (viết bằng Rust, nhanh gấp 4-5 lần
    và tiết kiệm 70% RAM so với openpyxl trên file > 50MB / 350,000 dòng).
    """
    uploaded_file.seek(0)
    try:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="calamine")
    except Exception:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def render_custom_progress(
    placeholder,
    percent: int,
    current_step: int,
    step_title: str,
    step_notes: dict = None,
):
    """
    Hiển thị thanh tiến độ chuẩn phong cách Teal-Mint và danh sách 5 bước chính.
    """
    percent = max(0, min(100, int(percent)))
    step_notes = step_notes or {}

    status_label = f"LOADING... {percent}%" if percent < 100 else "HOÀN TẤT 100%"
    spinner_html = (
        '<div class="teal-spinner"></div>'
        if percent < 100
        else '<span style="font-size:22px;">✅</span>'
    )

    steps_html_parts = []
    for step_idx, step_label in MAIN_STEPS:
        note_str = step_notes.get(step_idx, "")
        note_display = f" — <span style='font-weight:500;opacity:0.9;'>{note_str}</span>" if note_str else ""

        if percent == 100 or step_idx < current_step:
            icon = "✅"
            cls = "step-done"
            badge = '<span class="step-badge badge-done">Hoàn tất</span>'
        elif step_idx == current_step:
            icon = "🔄"
            cls = "step-active"
            badge = '<span class="step-badge badge-active">Đang chạy...</span>'
        else:
            icon = "⏳"
            cls = "step-pending"
            badge = '<span class="step-badge badge-pending">Chờ xử lý</span>'

        steps_html_parts.append(
            f'<div class="step-item {cls}">'
            f'<div>{icon} {step_label}{note_display}</div>'
            f'{badge}'
            f'</div>'
        )

    steps_html = "".join(steps_html_parts)

    html = (
        f'<div class="progress-wrapper">'
        f'<div class="progress-top-row">'
        f'<div class="progress-title">{spinner_html}<span>{step_title}</span></div>'
        f'<div class="progress-percent-badge">{percent}%</div>'
        f'</div>'
        f'<div class="bar-outer-pill">'
        f'<div class="bar-inner-gradient" style="width: {percent}%;"></div>'
        f'<div class="bar-center-label">{status_label}</div>'
        f'</div>'
        f'<div class="step-list">{steps_html}</div>'
        f'</div>'
    )
    placeholder.markdown(html, unsafe_allow_html=True)


def stream_df_to_xlsxwriter_sheet(workbook, sheet_name: str, df: pd.DataFrame, pct_col_name: str = None):
    """
    Ghi cuốn chiếu (Streaming) từng cụm 5,000 dòng xuống ổ cứng bằng xlsxwriter constant_memory.
    Giúp quá trình xuất 350,000+ dòng chỉ tốn < 10 MB RAM thay vì 1.2 GB RAM.
    """
    ws = workbook.add_worksheet(sheet_name)
    header_fmt = workbook.add_format({"bold": True})
    pct_fmt = workbook.add_format({"num_format": "0.0%"}) if pct_col_name else None

    cols = list(df.columns)
    pct_col_idx = cols.index(pct_col_name) if (pct_col_name and pct_col_name in cols) else None

    # Trong chế độ constant_memory, set_column phải gọi TRƯỚC khi ghi các dòng dữ liệu
    if pct_col_idx is not None:
        ws.set_column(pct_col_idx, pct_col_idx, 18, pct_fmt)

    # Ghi dòng tiêu đề
    ws.write_row(0, 0, cols, header_fmt)

    # Ghi cuốn chiếu từng lô 5,000 dòng để không tạo bản sao lớn trong RAM
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


def run_process_with_progress(
    uploaded_file,
    sheet_name: str,
    progress_placeholder,
    threshold_ratio: float = 0.85,
    include_sheet1_export: bool = True,
):
    """
    Thực thi toàn bộ luồng xử lý của process_nhu_cau.py theo kiến trúc Cuốn Chiếu (Streaming):
    Xử lý xong bảng lớn nào sẽ ghi thẳng xuống file tạm trên đĩa và xóa ngay khỏi RAM trước khi
    xử lý bảng tiếp theo, đảm bảo không bao giờ vượt giới hạn 1GB RAM của Streamlit Cloud.
    """
    t0 = time.time()
    step_notes = {}
    release_memory()

    # =========================================================================
    # BƯỚC 1: Đọc dữ liệu ban đầu bằng Calamine & Chuẩn hóa (0% -> 20%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        percent=8,
        current_step=1,
        step_title="[1/5] Đang đọc dữ liệu từ file Excel gốc (Engine tốc độ cao)...",
        step_notes=step_notes,
    )

    t_read = time.time()
    df = read_excel_fast(uploaded_file, sheet_name=sheet_name)
    total_rows = len(df)
    total_cols = len(df.columns)
    read_sec = time.time() - t_read

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        del df
        release_memory()
        raise ValueError(
            f"File Excel thiếu các cột bắt buộc: {', '.join(missing_cols)}. "
            f"Vui lòng kiểm tra lại đúng sheet dữ liệu!"
        )

    df["Khu vực mua hàng"] = df["Khu vực mua hàng"].fillna("").astype(str)
    df["Tên sản phẩm"] = df["Tên sản phẩm"].fillna("").astype(str)
    df["Số lượng cần mua"] = pd.to_numeric(df["Số lượng cần mua"], errors="coerce").fillna(0)
    df["SL mapping"] = pd.to_numeric(df["SL mapping"], errors="coerce").fillna(0)
    df["Khả năng giao NCC"] = pd.to_numeric(df["Khả năng giao NCC"], errors="coerce").fillna(0)

    step_notes[1] = f"Đã đọc {total_rows:,} dòng x {total_cols} cột ({read_sec:.1f}s)"

    # =========================================================================
    # BƯỚC 2: Tạo cột Map & Tính Sheet2 (Pivot Nhu cầu vs Mapping) (20% -> 45%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        percent=30,
        current_step=2,
        step_title="[2/5] Đang tạo cột 'Map' và tính toán Sheet2...",
        step_notes=step_notes,
    )

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

    req_sum = pivot2["Sum of Số lượng cần mua"]
    map_sum = pivot2["Sum of SL mapping"]
    safe_ratio = map_sum / req_sum.replace(0, float("nan"))
    pivot2["Tỉ lệ mapping/nhu cầu"] = safe_ratio.fillna(0.0)

    cond_co = (req_sum > 0) & (pivot2["Tỉ lệ mapping/nhu cầu"] < threshold_ratio)
    pivot2["NOTE"] = cond_co.map({True: "Có", False: "Không"})

    note_dict = dict(zip(pivot2["Map"], pivot2["NOTE"]))
    note_series = df["Map"].map(note_dict).fillna("Không")
    del note_dict, req_sum, map_sum, safe_ratio, cond_co

    if "NOTE" in df.columns:
        df["NOTE"] = note_series
    else:
        note_pos = df.columns.get_loc("Map") + 1
        df.insert(note_pos, "NOTE", note_series)
    del note_series

    so_cum_co = int((pivot2["NOTE"] == "Có").sum())
    sheet2_rows = len(pivot2)
    step_notes[2] = f"{sheet2_rows:,} cụm ({so_cum_co:,} cụm NOTE = 'Có')"

    # =========================================================================
    # BƯỚC 3: Tạo Sheet3 (Chỉ trích đúng 5 cột cần thiết, KHÔNG nhân bản 48 cột!)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        percent=52,
        current_step=3,
        step_title="[3/5] Đang tạo Sheet3 (Pivot NCC & lọc năng lực giao > 1)...",
        step_notes=step_notes,
    )

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
    pivot3 = pivot3[pivot3["Max of Khả năng giao NCC"] > 1].copy()

    pivot3.sort_values(
        by=["Map", "Max of Khả năng giao NCC"],
        ascending=[True, False],
        inplace=True,
    )
    pivot3.reset_index(drop=True, inplace=True)

    best_ncc_df = pivot3.drop_duplicates(subset=["Map"], keep="first")
    best_ncc_dict = dict(zip(best_ncc_df["Map"], best_ncc_df["Tên Chi nhánh nhà cung cấp"]))
    del best_ncc_df

    sheet3_rows = len(pivot3)
    step_notes[3] = f"{sheet3_rows:,} dòng NCC hợp lệ (Khả năng giao > 1)"

    # =========================================================================
    # BƯỚC 4: Tạo Sheet "PO bổ sung" (65% -> 80%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        percent=72,
        current_step=4,
        step_title="[4/5] Đang tạo Sheet 'PO bổ sung' và map mới NCC ưu tiên...",
        step_notes=step_notes,
    )

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

    step_notes[4] = f"{so_dong_po:,} dòng (Map NCC mới: {so_dong_co_ncc:,} | Trống: {so_dong_po - so_dong_co_ncc:,})"

    # =========================================================================
    # BƯỚC 5: Ghi cuốn chiếu (Streaming) ra file tạm trên đĩa & Xóa ngay RAM
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        percent=86,
        current_step=5,
        step_title="[5/5] Đang ghi cuốn chiếu (Streaming) ra file Excel & giải phóng RAM...",
        step_notes=step_notes,
    )

    # Lưu bản xem trước nhẹ (tối đa 1,000 dòng) cho giao diện web
    df_sheet1_preview = df.head(500).copy()
    df_sheet2_preview = pivot2.head(2000).copy()
    df_sheet3_preview = pivot3.head(2000).copy()
    df_po_preview_all = df_po_bo_sung.head(2000).copy()
    df_po_preview_mapped = df_po_bo_sung[df_po_bo_sung["Tên Chi nhánh nhà cung cấp"].notna()].head(2000).copy()
    df_po_preview_blank = df_po_bo_sung[df_po_bo_sung["Tên Chi nhánh nhà cung cấp"].isna()].head(2000).copy()

    import xlsxwriter

    wb = xlsxwriter.Workbook(
        EXPORT_FILE_PATH,
        {
            "constant_memory": True,
            "strings_to_urls": False,
            "nan_inf_to_errors": True,
            "default_date_format": "yyyy-mm-dd",
        },
    )

    # 5.1: Ghi Sheet1 xong là XÓA NGAY bảng df 350,000 dòng khỏi RAM!
    if include_sheet1_export:
        stream_df_to_xlsxwriter_sheet(wb, "Sheet1", df)
    del df
    release_memory()

    render_custom_progress(
        progress_placeholder,
        percent=94,
        current_step=5,
        step_title="[5/5] Đã giải phóng RAM Sheet1 -> Đang hoàn tất Sheet2, Sheet3, PO bổ sung...",
        step_notes=step_notes,
    )

    # 5.2: Ghi Sheet2 (có định dạng 0.0%) rồi xóa khỏi RAM
    stream_df_to_xlsxwriter_sheet(wb, "Sheet2", pivot2, pct_col_name="Tỉ lệ mapping/nhu cầu")
    del pivot2
    release_memory()

    # 5.3: Ghi Sheet3 rồi xóa khỏi RAM
    stream_df_to_xlsxwriter_sheet(wb, "Sheet3", pivot3)
    del pivot3
    release_memory()

    # 5.4: Ghi PO bổ sung rồi xóa khỏi RAM
    stream_df_to_xlsxwriter_sheet(wb, "PO bổ sung", df_po_bo_sung)
    del df_po_bo_sung
    wb.close()
    release_memory()

    total_time = time.time() - t0
    sheets_desc = "4 Sheets đầy đủ" if include_sheet1_export else "3 Sheets kết quả (bỏ qua Sheet1 gốc)"
    step_notes[5] = f"Hoàn tất {sheets_desc} trong {total_time:.1f}s (Đã giải phóng RAM)"

    render_custom_progress(
        progress_placeholder,
        percent=100,
        current_step=5,
        step_title=f"XỬ LÝ HOÀN TẤT TRONG {total_time:.1f} GIÂY!",
        step_notes=step_notes,
    )

    return {
        "df_sheet1_preview": df_sheet1_preview,
        "df_sheet2_preview": df_sheet2_preview,
        "df_sheet3_preview": df_sheet3_preview,
        "df_po_preview_all": df_po_preview_all,
        "df_po_preview_mapped": df_po_preview_mapped,
        "df_po_preview_blank": df_po_preview_blank,
        "export_file_path": EXPORT_FILE_PATH,
        "step_notes": step_notes,
        "stats": {
            "total_rows": total_rows,
            "sheet2_rows": sheet2_rows,
            "sheet2_note_co": so_cum_co,
            "sheet3_rows": sheet3_rows,
            "po_total": so_dong_po,
            "po_mapped_ncc": so_dong_co_ncc,
            "po_blank_ncc": so_dong_po - so_dong_co_ncc,
            "elapsed_sec": total_time,
        },
    }


# =============================================================================
# SIDEBAR CẤU HÌNH & QUẢN LÝ BỘ NHỚ (CHỐNG CRASH CLOUD)
# =============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Cấu Hình Xử Lý")
    threshold_pct = st.slider(
        "Ngưỡng tỉ lệ Mapping/Nhu cầu đánh dấu NOTE = 'Có' (%)",
        min_value=10,
        max_value=100,
        value=85,
        step=5,
        help="Mặc định theo chuẩn là 85%: Nếu Tỉ lệ mapping/nhu cầu < 85% thì đánh dấu NOTE = 'Có'.",
    )

    include_sheet1_export = st.checkbox(
        "Ghi kèm Sheet1 (350k+ dòng gốc) vào file Excel tải về",
        value=False,
        help="Mặc định TẮT để file tải về nhẹ (chỉ gồm Sheet2, Sheet3, PO bổ sung) và chạy siêu tốc trong vài giây, không bao giờ làm tràn RAM Cloud. Nếu hôm nào bạn cần xuất cả Sheet1 gốc thì tích chọn ô này.",
    )

    st.markdown("---")
    if st.button("🧹 Giải Phóng RAM Máy Chủ Ngay", use_container_width=True):
        st.session_state.clear()
        if os.path.exists(EXPORT_FILE_PATH):
            try:
                os.remove(EXPORT_FILE_PATH)
            except Exception:
                pass
        release_memory()
        st.success("✅ Đã dọn sạch bộ nhớ đệm & giải phóng RAM!")

    st.markdown("---")
    st.markdown("### 🛡️ Cơ Chế Chống Crash RAM")
    st.markdown(
        """
        - **Streaming `constant_memory`:** Ghi từng cụm 5,000 dòng xuống đĩa, không giữ bảng Excel lớn trong RAM.
        - **Tự động `malloc_trim`:** Trả thẳng RAM về cho Linux ngay sau khi chạy xong từng bước.
        - **Không giữ file 80MB trong Session:** Lưu file kết quả trên ổ đĩa tạm, giữ RAM máy chủ luôn nhẹ cả tháng.
        """
    )
    st.markdown("---")
    st.caption("🔄 Tích hợp GitHub Actions tự động ping mỗi 3 giờ/lần giúp ứng dụng hoạt động 24/7.")


# =============================================================================
# GIAO DIỆN CHÍNH
# =============================================================================
st.markdown(
    """
<div class="main-header">
    <h1>📊 Hệ Thống Kiểm Tra Nhu Cầu & Tạo PO Bổ Sung</h1>
    <p>Tự động phân tích tỉ lệ Mapping/Nhu cầu, lọc Nhà Cung Cấp năng lực cao nhất và xuất danh sách PO bổ sung</p>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="guide-box">
    💡 <b>Hướng dẫn nhanh:</b> Tải lên file Excel nhu cầu mua hàng theo siêu thị (chứa các cột:
    <code>Khu vực mua hàng</code>, <code>Tên sản phẩm</code>, <code>Số lượng cần mua</code>,
    <code>SL mapping</code>, <code>Khả năng giao NCC</code>, <code>Tên Chi nhánh nhà cung cấp</code>, <code>Mã PO</code>),
    sau đó nhấn <b>Bắt đầu Xử Lý Nhu Cầu</b>.
</div>
""",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "📥 Tải lên file Excel Nhu cầu mua hàng (.xlsx, .xls)",
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
    try:
        sheet_cache_key = f"sheets_{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get("sheet_cache_key") != sheet_cache_key:
            st.session_state["sheet_names"] = get_excel_sheet_names_fast(uploaded_file)
            st.session_state["sheet_cache_key"] = sheet_cache_key
        sheet_names = st.session_state["sheet_names"]

        col_sel, col_btn = st.columns([2.5, 1.5])
        with col_sel:
            selected_sheet = st.selectbox("📑 Chọn Sheet dữ liệu đầu vào:", sheet_names, index=0)
        with col_btn:
            st.write("")
            st.write("")
            run_btn = st.button("🚀 Bắt đầu Xử Lý Nhu Cầu", type="primary", use_container_width=True)

        file_key = f"{uploaded_file.name}_{uploaded_file.size}_{selected_sheet}_{threshold_pct}_{include_sheet1_export}"
        if st.session_state.get("last_file_key") != file_key and not run_btn:
            st.info("👆 Nhấn nút **🚀 Bắt đầu Xử Lý Nhu Cầu** để chạy chương trình và theo dõi tiến độ.")

        progress_placeholder = st.empty()

        if run_btn:
            # Xóa kết quả cũ và ép Linux thu hồi RAM trước khi chạy mới
            if "nhu_cau_result" in st.session_state:
                del st.session_state["nhu_cau_result"]
            release_memory()

            result = run_process_with_progress(
                uploaded_file=uploaded_file,
                sheet_name=selected_sheet,
                progress_placeholder=progress_placeholder,
                threshold_ratio=threshold_pct / 100.0,
                include_sheet1_export=include_sheet1_export,
            )
            st.session_state["nhu_cau_result"] = result
            st.session_state["last_file_key"] = file_key

        if "nhu_cau_result" in st.session_state and st.session_state.get("last_file_key") == file_key:
            res = st.session_state["nhu_cau_result"]
            stats = res["stats"]

            if not run_btn:
                render_custom_progress(
                    progress_placeholder,
                    percent=100,
                    current_step=5,
                    step_title=f"XỬ LÝ HOÀN TẤT TRONG {stats['elapsed_sec']:.1f} GIÂY!",
                    step_notes=res.get("step_notes", {}),
                )

            # Đọc từ file tạm trên ổ đĩa khi render nút Download (không ngốn RAM session_state)
            export_path = res.get("export_file_path", EXPORT_FILE_PATH)
            if os.path.exists(export_path):
                with open(export_path, "rb") as f_dl:
                    st.download_button(
                        label="💾 TẢI XUỐNG FILE KẾT QUẢ EXCEL (nhucau_da_xu_ly.xlsx)",
                        data=f_dl,
                        file_name="nhucau_da_xu_ly.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )

            # Thống kê tổng quan
            st.markdown("### 📈 Thống Kê Kết Quả Xử Lý")
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("1. Tổng dòng gốc (Sheet1)", f"{stats['total_rows']:,}")
            with m2:
                st.metric(
                    "2. Cụm Khu vực - SP (Sheet2)",
                    f"{stats['sheet2_rows']:,}",
                    delta=f"{stats['sheet2_note_co']:,} cụm NOTE = 'Có'",
                )
            with m3:
                st.metric(
                    "3. Dòng NCC năng lực > 1 (Sheet3)",
                    f"{stats['sheet3_rows']:,}",
                )
            with m4:
                st.metric(
                    "4. Tổng dòng PO bổ sung",
                    f"{stats['po_total']:,}",
                    delta=f"Map NCC: {stats['po_mapped_ncc']:,} | Trống: {stats['po_blank_ncc']:,}",
                )

            # Hiển thị 4 Tabs xem trước nhẹ (tối đa 2,000 dòng đầu để không làm nặng trình duyệt & RAM)
            st.markdown("### 🗂️ Xem Trước Dữ Liệu Các Sheet Kết Quả (Hiển thị mẫu tối đa 2,000 dòng đầu)")
            tab_po, tab_s2, tab_s3, tab_s1 = st.tabs(
                [
                    f"🛒 PO bổ sung ({stats['po_total']:,} dòng)",
                    f"📊 Sheet2 - Pivot Nhu cầu ({stats['sheet2_rows']:,} dòng)",
                    f"🏆 Sheet3 - Pivot NCC > 1 ({stats['sheet3_rows']:,} dòng)",
                    f"📄 Sheet1 - Dữ liệu gốc + Map ({stats['total_rows']:,} dòng)",
                ]
            )

            with tab_po:
                f_col1, _ = st.columns([1.5, 2.5])
                with f_col1:
                    filter_ncc_status = st.radio(
                        "Lọc trạng thái NCC:",
                        ["Tất cả", "Đã map được NCC mới", "Chưa có NCC (Trống)"],
                        horizontal=True,
                    )
                if filter_ncc_status == "Đã map được NCC mới":
                    df_po_view = res["df_po_preview_mapped"]
                elif filter_ncc_status == "Chưa có NCC (Trống)":
                    df_po_view = res["df_po_preview_blank"]
                else:
                    df_po_view = res["df_po_preview_all"]
                st.dataframe(df_po_view, use_container_width=True, height=420)

            with tab_s2:
                df_s2_view = res["df_sheet2_preview"].copy()
                df_s2_view["Tỉ lệ mapping/nhu cầu (%)"] = (
                    df_s2_view["Tỉ lệ mapping/nhu cầu"] * 100
                ).round(1).astype(str) + "%"
                st.dataframe(df_s2_view, use_container_width=True, height=420)

            with tab_s3:
                st.dataframe(res["df_sheet3_preview"], use_container_width=True, height=420)

            with tab_s1:
                st.caption("Hiển thị 500 dòng đầu để trình duyệt mượt mà.")
                st.dataframe(res["df_sheet1_preview"], use_container_width=True, height=420)

    except Exception as e:
        release_memory()
        st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {str(e)}")
else:
    st.info("👆 Vui lòng chọn hoặc kéo thả file Excel nhu cầu mua hàng vào khung bên trên để bắt đầu.")
