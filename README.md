# 📊 Hệ Thống Kiểm Tra Nhu Cầu & Tạo PO Bổ Sung (Streamlit Web App)

Ứng dụng web xây dựng bằng **Streamlit** giúp tự động hóa toàn bộ quy trình kiểm tra nhu cầu mua hàng theo siêu thị (`process_nhu_cau.py`), tính tỉ lệ Mapping/Nhu cầu, lọc Nhà Cung Cấp có năng lực giao cao nhất và xuất file `nhucau_da_xu_ly.xlsx` gồm đầy đủ 4 Sheets chuẩn.

---

## ✨ Tính Năng Nổi Bật
- **Thanh tiến độ (Progress Bar) trực quan theo phần trăm (`0%` -> `100%`):**
  - Thiết kế đồng bộ tông màu Teal-Mint hiện đại gồm: Thanh Gradient kèm chữ `LOADING... XX%`, Thanh sọc chéo chuyển động (Striped Animation), và Thanh 10 khối phân đoạn (Segmented Blocks).
  - Hiển thị chi tiết từng bước từ `[1/5]` đến `[5/5]` khi chương trình đang chạy.
- **Giữ nguyên 100% logic chuẩn từ `process_nhu_cau.py`:**
  1. **Sheet1:** Dữ liệu gốc + cột `Map` (`Trim(Khu vực mua hàng) - Trim(Tên sản phẩm)`) + cột `NOTE`.
  2. **Sheet2:** Pivot tổng `Số lượng cần mua` & `SL mapping`, tính `Tỉ lệ mapping/nhu cầu` (định dạng `0.0%` trong Excel) và đánh dấu `NOTE = 'Có'` nếu tỉ lệ `< 85%`.
  3. **Sheet3:** Pivot `Max Khả năng giao NCC`, lọc bỏ các dòng trống hoặc `<= 1`, sắp xếp giảm dần theo năng lực giao cho từng `Map`.
  4. **PO bổ sung:** Lọc các dòng `NOTE = 'Có'` và `Mã PO` trống từ `Sheet1`, xóa trắng hoàn toàn NCC cũ và map NCC có năng lực giao cao nhất từ `Sheet3`.
- **Tự động Ping mỗi 3 tiếng/lần chống ngủ (GitHub Actions):**
  - Tích hợp sẵn workflow `.github/workflows/keep_alive.yml` chạy định kỳ mỗi 3 giờ (`0 */3 * * *`) để giữ ứng dụng Streamlit Cloud luôn thức.

---

## 🛠️ Chạy Thử Nghiệm Trên Máy (Local)

```bash
cd "C:\MWG\Thử nghiệm\Check Nhu cầu"
pip install -r requirements.txt
streamlit run app.py
```

---

## 🌐 Hướng Dẫn Đưa Lên GitHub & Deploy Streamlit Cloud

### Bước 1: Đẩy code lên GitHub Repository mới
Mở Terminal / PowerShell tại thư mục `C:\MWG\Thử nghiệm\Check Nhu cầu`:
```bash
git init
git add app.py process_nhu_cau.py requirements.txt README.md .gitignore .github/
git commit -m "Deploy ứng dụng Check Nhu Cầu lên Streamlit kèm Keep Alive 3h"
git branch -M main
git remote add origin https://github.com/<TEN_TAI_KHOAN_GITHUB>/<TEN_REPO>.git
git push -u origin main
```

### Bước 2: Deploy trên Streamlit Community Cloud
1. Truy cập [https://share.streamlit.io/](https://share.streamlit.io/) và đăng nhập bằng GitHub.
2. Nhấn **Create app** (hoặc **New app**).
3. Chọn:
   - **Repository:** `<TEN_TAI_KHOAN_GITHUB>/<TEN_REPO>`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Nhấn **Deploy!**.

### Bước 3: Cập nhật link App vào GitHub Actions (Ping 3h/lần)
Sau khi có đường dẫn web `.streamlit.app`, bạn có thể:
- **Cách 1 (Khuyên dùng):** Vào GitHub Repo -> **Settings** -> **Secrets and variables** -> **Actions** -> **New repository secret**:
  - **Name:** `STREAMLIT_APP_URL`
  - **Secret:** Dán đường link Streamlit của bạn (ví dụ: `https://check-nhu-cau-mwg.streamlit.app/`)
- **Cách 2:** Sửa trực tiếp đường link mặc định trong file `.github/workflows/keep_alive.yml`.
