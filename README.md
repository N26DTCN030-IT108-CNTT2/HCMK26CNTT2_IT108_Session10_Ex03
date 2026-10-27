# HCMK26CNTT2_IT108_Session10_Ex03
Bài tập số 3
# Bộ Hồ Sơ Giải Pháp Tự Động Hóa Giao Vận FastMove (Session 12 Ex03)

Bài giải chuẩn hóa quy trình giao vận, khắc phục sự cố tra cứu cước phí thủ công 16 tiếng/ngày và tối ưu hóa vận hành kho cho Công ty Giao hàng FastMove.

---

# Phần 1. Phân Tích Mô Hình VLOOKUP Trong Nghiệp Vụ Kho (Mapping Process)

Ví ẩn dụ nghiệp vụ tra cứu giá cước vận đơn tương đương với hành động **"Nhân viên kho tra cứu hồ sơ cứng trong lưu trữ"**:

| Tham số hàm VLOOKUP | Tên kỹ thuật trong Excel | Hành động tương đương của Nhân viên kho | Giá trị thực thi trong bài toán FastMove |
| :--- | :--- | :--- | :--- |
| **Lookup_value** | Giá trị tìm kiếm | Đọc mã vận đơn in trên kiện hàng đang cầm trên tay | `A2` (Mã vận đơn cần tra cước, ví dụ: `FM001`) |
| **Table_array** | Vùng bảng tra cứu | Đi tới tủ hồ sơ niêm yết Bảng giá cước quy định của công ty | `BangGiaCuoc!A$2:C$500` (Cố định bằng dấu `$`) |
| **Col_index_num** | Cột lấy dữ liệu | Dóng mắt sang Cột số 2 trong bảng giá để lấy con số tiền cước | `2` (Cột chứa đơn giá cước quy định) |
| **Range_lookup** | Kiểu tìm kiếm | Yêu cầu mã trên đơn phải khớp chính xác 100% với mã trên bảng | `FALSE` (Tìm kiếm chính xác tuyệt đối) |

---

# Phần 2. Triển Khai Giải Pháp Công Nghệ Cho FastMove

## 1. File Word Quy Trình Vận Hành Kho (`Quy_Trinh_Van_Hanh_Kho_FastMove.docx`)
* **Thương hiệu & Màu sắc:** Sử dụng hệ màu nhận diện thương hiệu FastMove (Màu chủ đạo Xanh Navy `#002D62` cho `Heading 1` và Màu nhấn Cam `#FF6600` cho `Heading 2`).
* **Chuẩn hóa Heading Styles:** Phân cấp rõ ràng các bước tiếp nhận hàng, quét mã bar-code, truy vấn đơn giá tự động và bàn giao cho Shipper.
* **Auto TOC:** Tích hợp Mục lục tự động `Table of Contents` tại trang 1, đảm bảo số trang luôn được cập nhật chính xác thời gian thực.

---

## 2. File Excel Bảng Tra Cứu Giá Cước (`Bang_Tra_Cuu_Gia_Cuoc_FastMove.xlsx`)
* **Tuân thủ 5 Quy tắc Vàng Database:**
  1. Hàng 1 là **Header Row duy nhất**, không gộp ô (Merge Cell).
  2. Dữ liệu nguyên tử (Atomic Data): Mã vận đơn, Tên khách hàng, Trọng lượng, Khu vực, Cước phí và Trạng thái nằm trên các cột riêng biệt.
  3. Không chứa dòng/cột trống ngắt đoạn dữ liệu.
  4. Cột số tiền định dạng Number đồng nhất.
  5. Dòng tổng doanh thu dự kiến (`SUM`) được đưa ra ngoài vùng dữ liệu thô (Dòng 7).
* **Công thức tra cứu cố định dấu `$`: **
  Tại ô `E2` sheet `BangVanDon`, nhập công thức:
  ```excel
  =IFERROR(VLOOKUP(A2, BangGiaCuoc!A$2:C$500, 2, FALSE), "Mã lỗi/Cần kiểm tra")

2. Giải pháp bọc hàm IFERROR
Bọc toàn bộ hàm VLOOKUP bên trong hàm IFERROR với thông báo xử lý ngoại lệ:
=IFERROR(VLOOKUP(A2, BangGiaCuoc!A$2:C$500, 2, FALSE), "Mã lỗi/Cần kiểm tra")




  
