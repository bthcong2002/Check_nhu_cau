import io
import time
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

# Custom CSS lấy cảm hứng từ bộ Progress Bar Teal-Mint của người dùng
st.markdown(
    """
    <style>
    /* Header chính */
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

    /* Hộp hướng dẫn */
    .guide-box {
        background-color: #f4faf8;
        border-left: 5px solid #377375;
        padding: 15px 20px;
        border-radius: 0 10px 10px 0;
        margin-bottom: 20px;
        font-size: 14.5px;
        color: #1e3a3f;
    }

    /* Khung chứa Thanh Tiến Độ Custom (theo mẫu ảnh upload) */
    .progress-wrapper {
        background: #ffffff;
        border: 2px solid #24474e;
        border-radius: 18px;
        padding: 22px 26px;
        margin: 18px 0 24px 0;
        box-shadow: 0 8px 24px rgba(36, 71, 78, 0.10);
    }

    .progress-top-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 12px;
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
        font-size: 16px;
        padding: 5px 16px;
        border-radius: 999px;
        letter-spacing: 0.8px;
    }

    /* Spinner tròn đồng bộ tông màu Teal-Mint */
    .teal-spinner {
        width: 26px;
        height: 26px;
        border: 3.5px solid #c8e6dc;
        border-top: 3.5px solid #24474e;
        border-right: 3.5px solid #377375;
        border-radius: 50%;
        animation: spin 0.9s linear infinite;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }

    /* Kiểu 1: Thanh Gradient mượt có chữ LOADING... bên trong */
    .bar-outer-pill {
        width: 100%;
        height: 40px;
        background: #ffffff;
        border: 3px solid #1e2d2f;
        border-radius: 12px;
        padding: 4px;
        position: relative;
        overflow: hidden;
        box-sizing: border-box;
    }
    .bar-inner-gradient {
        height: 100%;
        border-radius: 7px;
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
        font-size: 15px;
        letter-spacing: 1.5px;
        color: #162527;
        text-shadow: 0 0 6px rgba(255, 255, 255, 0.85);
        pointer-events: none;
    }

    /* Kiểu 2: Thanh sọc chéo chuyển động (Striped) */
    .bar-outer-striped {
        width: 100%;
        height: 28px;
        background: #ffffff;
        border: 2.5px solid #1e2d2f;
        border-radius: 999px;
        padding: 4px;
        margin-top: 12px;
        box-sizing: border-box;
        overflow: hidden;
    }
    .bar-inner-striped {
        height: 100%;
        border-radius: 999px;
        background: repeating-linear-gradient(
            -55deg,
            #24474e,
            #24474e 12px,
            #ffffff 12px,
            #ffffff 18px,
            #3f8482 18px,
            #3f8482 30px,
            #ffffff 30px,
            #ffffff 36px,
            #68b89e 36px,
            #68b89e 48px
        );
        background-size: 68px 68px;
        animation: moveStripes 1.2s linear infinite;
        transition: width 0.35s ease;
    }
    @keyframes moveStripes {
        0% { background-position: 0 0; }
        100% { background-position: 68px 0; }
    }

    /* Kiểu 3: Các khối phân đoạn (Segmented Blocks) */
    .segmented-row {
        display: grid;
        grid-template-columns: repeat(10, 1fr);
        gap: 6px;
        margin-top: 12px;
        padding: 5px;
        border: 2.5px solid #1e2d2f;
        border-radius: 10px;
        background: #ffffff;
    }
    .seg-block {
        height: 20px;
        border-radius: 4px;
        transition: all 0.25s ease;
    }

    .progress-step-detail {
        margin-top: 10px;
        font-size: 14px;
        color: #375a60;
        font-weight: 500;
        display: flex;
        justify-content: space-between;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_custom_progress(placeholder, percent: int, step_title: str, detail_text: str = ""):
    """
    Hiển thị thanh tiến độ chuẩn phong cách Teal-Mint theo hình mẫu người dùng yêu cầu:
    - Thanh chính: Gradient đậm -> nhạt kèm chữ LOADING... XX%
    - Thanh phụ: 10 khối phân đoạn (Segmented Blocks) đổi màu theo từng 10%
    """
    percent = max(0, min(100, int(percent)))
    seg_colors = [
        "#24474e", "#295259", "#2f5e64", "#366b70", "#3d787b",
        "#468685", "#4f948e", "#58a194", "#60ad99", "#68b89e",
    ]
    active_blocks = percent // 10
    blocks_html = ""
    for idx in range(10):
        if idx < active_blocks:
            bg = seg_colors[idx]
            opacity = "1"
        elif idx == active_blocks and percent % 10 > 0:
            bg = seg_colors[idx]
            opacity = "0.45"
        else:
            bg = "#e9f2ef"
            opacity = "1"
        blocks_html += f'<div class="seg-block" style="background:{bg}; opacity:{opacity};"></div>'

    status_label = f"LOADING... {percent}%" if percent < 100 else f"HOÀN TẤT 100%"
    spinner_html = '<div class="teal-spinner"></div>' if percent < 100 else '<span style="font-size:22px;">✅</span>'

    html = f"""
    <div class="progress-wrapper">
        <div class="progress-top-row">
            <div class="progress-title">
                {spinner_html}
                <span>{step_title}</span>
            </div>
            <div class="progress-percent-badge">{percent}%</div>
        </div>

        <!-- Thanh Gradient chính kèm chữ LOADING... -->
        <div class="bar-outer-pill">
            <div class="bar-inner-gradient" style="width: {percent}%;"></div>
            <div class="bar-center-label">{status_label}</div>
        </div>

        <!-- Thanh Sọc Chéo Chuyển Động -->
        <div class="bar-outer-striped">
            <div class="bar-inner-striped" style="width: {percent}%;"></div>
        </div>

        <!-- Thanh Khối Phân Đoạn 10 Mức -->
        <div class="segmented-row">
            {blocks_html}
        </div>

        <div class="progress-step-detail">
            <span>{detail_text}</span>
            <span>Tiến độ: {percent}/100%</span>
        </div>
    </div>
    """
    placeholder.markdown(html, unsafe_allow_html=True)


REQUIRED_COLUMNS = [
    "Khu vực mua hàng",
    "Tên sản phẩm",
    "Số lượng cần mua",
    "SL mapping",
    "Khả năng giao NCC",
    "Tên Chi nhánh nhà cung cấp",
    "Mã PO",
]


def run_process_with_progress(uploaded_file, sheet_name, progress_placeholder, threshold_ratio=0.85):
    """
    Thực thi toàn bộ luồng xử lý của process_nhu_cau.py và cập nhật thanh tiến độ % theo thời gian thực.
    Trả về dict chứa các DataFrame kết quả, file Excel bytes và thống kê.
    """
    t0 = time.time()

    # =========================================================================
    # BƯỚC 1: Đọc dữ liệu ban đầu & Chuẩn hóa (0% -> 25%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        5,
        "[1/5] Đang đọc dữ liệu từ file Excel gốc...",
        "Đang nạp dữ liệu vào bộ nhớ...",
    )

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    total_rows = len(df)
    total_cols = len(df.columns)

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        raise ValueError(
            f"File Excel thiếu các cột bắt buộc: {', '.join(missing_cols)}. "
            f"Vui lòng kiểm tra lại đúng sheet dữ liệu!"
        )

    render_custom_progress(
        progress_placeholder,
        18,
        "[1/5] Đang chuẩn hóa kiểu dữ liệu chuỗi và số...",
        f"Đã đọc {total_rows:,} dòng và {total_cols} cột.",
    )

    df["Khu vực mua hàng"] = df["Khu vực mua hàng"].fillna("").astype(str)
    df["Tên sản phẩm"] = df["Tên sản phẩm"].fillna("").astype(str)
    df["Số lượng cần mua"] = pd.to_numeric(df["Số lượng cần mua"], errors="coerce").fillna(0)
    df["SL mapping"] = pd.to_numeric(df["SL mapping"], errors="coerce").fillna(0)
    df["Khả năng giao NCC"] = pd.to_numeric(df["Khả năng giao NCC"], errors="coerce").fillna(0)

    # =========================================================================
    # BƯỚC 2: Tạo cột Map & Tính Sheet2 (Pivot Nhu cầu vs Mapping) (25% -> 50%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        30,
        "[2/5] Đang tạo cột 'Map' và tính toán Sheet2 (Pivot Nhu cầu vs SL mapping)...",
        "Công thức Map = Trim(Khu vực mua hàng) - Trim(Tên sản phẩm)",
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

    if "NOTE" in df.columns:
        df["NOTE"] = note_series
    else:
        note_pos = df.columns.get_loc("Map") + 1
        df.insert(note_pos, "NOTE", note_series)

    render_custom_progress(
        progress_placeholder,
        48,
        "[2/5] Đã hoàn tất Sheet2 & map ngược cột NOTE về Sheet1",
        f"Sheet2 có {len(pivot2):,} cụm Khu vực - Sản phẩm.",
    )

    # =========================================================================
    # BƯỚC 3: Tạo Sheet3 (Pivot NCC & lọc năng lực giao cao nhất) (50% -> 70%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        58,
        "[3/5] Đang tạo Sheet3 (Pivot NCC & lọc Khả năng giao > 1)...",
        "Sắp xếp theo Map và Max khả năng giao giảm dần (ưu tiên NCC năng lực cao nhất)...",
    )

    ncc_clean = df["Tên Chi nhánh nhà cung cấp"].fillna("").astype(str).str.strip()
    df_valid_ncc = df[
        (ncc_clean != "")
        & (ncc_clean.str.lower() != "none")
        & (ncc_clean.str.lower() != "nan")
        & (ncc_clean != "(blank)")
    ].copy()

    pivot3 = df_valid_ncc.groupby(
        ["Khu vực mua hàng", "Tên sản phẩm", "Tên Chi nhánh nhà cung cấp", "Map"],
        as_index=False,
        dropna=False,
    )["Khả năng giao NCC"].max()

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

    best_ncc_df = pivot3.drop_duplicates(subset=["Map"], keep="first")
    best_ncc_dict = dict(zip(best_ncc_df["Map"], best_ncc_df["Tên Chi nhánh nhà cung cấp"]))

    render_custom_progress(
        progress_placeholder,
        70,
        "[3/5] Đã xác định danh sách NCC ưu tiên năng lực cao nhất",
        f"Sheet3 có {len(pivot3):,} dòng NCC hợp lệ (Khả năng giao > 1).",
    )

    # =========================================================================
    # BƯỚC 4: Tạo Sheet "PO bổ sung" (70% -> 85%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        78,
        "[4/5] Đang tạo Sheet 'PO bổ sung' và map mới NCC ưu tiên...",
        "Lọc NOTE = 'Có' & Mã PO trống -> Làm trắng NCC cũ -> Map NCC mới từ Sheet3...",
    )

    ma_po_str = df["Mã PO"].fillna("").astype(str).str.strip()
    cond_ma_po_blank = (
        (ma_po_str == "")
        | (ma_po_str.str.lower() == "nan")
        | (ma_po_str.str.lower() == "none")
    )

    cond_po_bo_sung = (df["NOTE"] == "Có") & cond_ma_po_blank
    df_po_bo_sung = df[cond_po_bo_sung].copy()

    # Làm trắng sạch cột 'Tên Chi nhánh nhà cung cấp' cũ trước khi map
    df_po_bo_sung["Tên Chi nhánh nhà cung cấp"] = None
    df_po_bo_sung["Tên Chi nhánh nhà cung cấp"] = df_po_bo_sung["Map"].map(best_ncc_dict)

    so_dong_po = len(df_po_bo_sung)
    so_dong_co_ncc = int(df_po_bo_sung["Tên Chi nhánh nhà cung cấp"].notna().sum())

    # =========================================================================
    # BƯỚC 5: Xuất toàn bộ 4 Sheets ra bộ nhớ Excel (85% -> 100%)
    # =========================================================================
    render_custom_progress(
        progress_placeholder,
        88,
        "[5/5] Đang kết xuất 4 Sheets ra file Excel chuẩn (.xlsx)...",
        f"Đang ghi Sheet1 ({len(df):,} dòng), Sheet2, Sheet3 và PO bổ sung ({so_dong_po:,} dòng)...",
    )

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        render_custom_progress(
            progress_placeholder,
            93,
            "[5/5] Đang ghi Sheet2, Sheet3 và định dạng tỉ lệ %...",
            "Đã ghi xong Sheet1 -> Đang ghi Sheet2, Sheet3, PO bổ sung...",
        )
        pivot2.to_excel(writer, sheet_name="Sheet2", index=False)
        pivot3.to_excel(writer, sheet_name="Sheet3", index=False)
        df_po_bo_sung.to_excel(writer, sheet_name="PO bổ sung", index=False)

        # Định dạng % cho cột tỉ lệ ở Sheet2
        ws2 = writer.sheets["Sheet2"]
        ratio_col = pivot2.columns.get_loc("Tỉ lệ mapping/nhu cầu") + 1
        for row in range(2, len(pivot2) + 2):
            ws2.cell(row=row, column=ratio_col).number_format = "0.0%"

    output_buffer.seek(0)
    total_time = time.time() - t0

    render_custom_progress(
        progress_placeholder,
        100,
        f"XỬ LÝ HOÀN TẤT TRONG {total_time:.1f} GIÂY!",
        f"Tổng PO bổ sung: {so_dong_po:,} dòng (Map được NCC mới: {so_dong_co_ncc:,} | Trống: {so_dong_po - so_dong_co_ncc:,})",
    )

    return {
        "df_sheet1": df,
        "df_sheet2": pivot2,
        "df_sheet3": pivot3,
        "df_po_bo_sung": df_po_bo_sung,
        "excel_bytes": output_buffer.getvalue(),
        "stats": {
            "total_rows": total_rows,
            "sheet2_rows": len(pivot2),
            "sheet2_note_co": int((pivot2["NOTE"] == "Có").sum()),
            "sheet3_rows": len(pivot3),
            "po_total": so_dong_po,
            "po_mapped_ncc": so_dong_co_ncc,
            "po_blank_ncc": so_dong_po - so_dong_co_ncc,
            "elapsed_sec": total_time,
        },
    }


# =============================================================================
# SIDEBAR CẤU HÌNH & THÔNG TIN LUỒNG XỬ LÝ
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

    st.markdown("---")
    st.markdown("### 📋 Quy Trình 5 Bước Chuẩn")
    st.markdown(
        """
        1. **Bước 1:** Đọc file gốc & chuẩn hóa dữ liệu số/chuỗi, tạo cột `Map = Khu vực - Sản phẩm`.
        2. **Bước 2 (Sheet2):** Pivot tổng `Số lượng cần mua` & `SL mapping`, tính `Tỉ lệ mapping/nhu cầu` và gán `NOTE` (`Có` nếu `< 85%`), map ngược vào `Sheet1`.
        3. **Bước 3 (Sheet3):** Pivot `Max Khả năng giao NCC`, lọc bỏ `<= 1`, sắp xếp giảm dần theo năng lực giao của từng `Map`.
        4. **Bước 4 (PO bổ sung):** Lọc các dòng `NOTE = 'Có'` và `Mã PO` trống, xóa sạch NCC cũ và gán NCC có năng lực giao cao nhất từ `Sheet3`.
        5. **Bước 5:** Xuất file Excel gồm đủ 4 Sheets (`Sheet1`, `Sheet2`, `Sheet3`, `PO bổ sung`).
        """
    )
    st.markdown("---")
    st.caption("🔄 Đã tích hợp GitHub Actions tự động ping mỗi 3 giờ/lần giúp ứng dụng luôn hoạt động 24/7.")


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
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names

        col_sel, col_btn = st.columns([2.5, 1.5])
        with col_sel:
            selected_sheet = st.selectbox("📑 Chọn Sheet dữ liệu đầu vào:", sheet_names, index=0)
        with col_btn:
            st.write("")
            st.write("")
            run_btn = st.button("🚀 Bắt đầu Xử Lý Nhu Cầu", type="primary", use_container_width=True)

        # Reset kết quả cũ nếu đổi file
        file_key = f"{uploaded_file.name}_{uploaded_file.size}_{selected_sheet}_{threshold_pct}"
        if st.session_state.get("last_file_key") != file_key and not run_btn:
            st.info("👆 Nhấn nút **🚀 Bắt đầu Xử Lý Nhu Cầu** để chạy chương trình và theo dõi thanh tiến độ.")

        progress_placeholder = st.empty()

        if run_btn:
            result = run_process_with_progress(
                uploaded_file=uploaded_file,
                sheet_name=selected_sheet,
                progress_placeholder=progress_placeholder,
                threshold_ratio=threshold_pct / 100.0,
            )
            st.session_state["nhu_cau_result"] = result
            st.session_state["last_file_key"] = file_key

        if "nhu_cau_result" in st.session_state and st.session_state.get("last_file_key") == file_key:
            res = st.session_state["nhu_cau_result"]
            stats = res["stats"]

            if not run_btn:
                render_custom_progress(
                    progress_placeholder,
                    100,
                    f"ĐÃ XỬ LÝ HOÀN TẤT ({stats['elapsed_sec']:.1f} GIÂY)",
                    f"Tổng PO bổ sung: {stats['po_total']:,} dòng | Đã map NCC mới: {stats['po_mapped_ncc']:,} dòng",
                )

            # Nút tải file nổi bật ngay phía trên
            st.download_button(
                label="💾 TẢI XUỐNG FILE KẾT QUẢ ĐẦY ĐỦ 4 SHEETS (nhucau_da_xu_ly.xlsx)",
                data=res["excel_bytes"],
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

            # Hiển thị 4 Tabs tương ứng 4 Sheets
            st.markdown("### 🗂️ Xem Trước Dữ Liệu Các Sheet Kết Quả")
            tab_po, tab_s2, tab_s3, tab_s1 = st.tabs(
                [
                    f"🛒 PO bổ sung ({stats['po_total']:,} dòng)",
                    f"📊 Sheet2 - Pivot Nhu cầu ({stats['sheet2_rows']:,} dòng)",
                    f"🏆 Sheet3 - Pivot NCC > 1 ({stats['sheet3_rows']:,} dòng)",
                    f"📄 Sheet1 - Dữ liệu gốc + Map ({stats['total_rows']:,} dòng)",
                ]
            )

            with tab_po:
                f_col1, f_col2 = st.columns([1.5, 2.5])
                with f_col1:
                    filter_ncc_status = st.radio(
                        "Lọc trạng thái NCC:",
                        ["Tất cả", "Đã map được NCC mới", "Chưa có NCC (Trống)"],
                        horizontal=True,
                    )
                df_po_view = res["df_po_bo_sung"]
                if filter_ncc_status == "Đã map được NCC mới":
                    df_po_view = df_po_view[df_po_view["Tên Chi nhánh nhà cung cấp"].notna()]
                elif filter_ncc_status == "Chưa có NCC (Trống)":
                    df_po_view = df_po_view[df_po_view["Tên Chi nhánh nhà cung cấp"].isna()]
                st.dataframe(df_po_view, use_container_width=True, height=420)

            with tab_s2:
                df_s2_view = res["df_sheet2"].copy()
                df_s2_view["Tỉ lệ mapping/nhu cầu (%)"] = (
                    df_s2_view["Tỉ lệ mapping/nhu cầu"] * 100
                ).round(1).astype(str) + "%"
                st.dataframe(df_s2_view, use_container_width=True, height=420)

            with tab_s3:
                st.dataframe(res["df_sheet3"], use_container_width=True, height=420)

            with tab_s1:
                st.caption("Hiển thị tối đa 1,000 dòng đầu để trình duyệt mượt mà (File tải về chứa đầy đủ 100% dòng).")
                st.dataframe(res["df_sheet1"].head(1000), use_container_width=True, height=420)

    except Exception as e:
        st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {str(e)}")
else:
    st.info("👆 Vui lòng chọn hoặc kéo thả file Excel nhu cầu mua hàng vào khung bên trên để bắt đầu.")
