# Excel tutorial

Cảm ơn bạn đã ghé thăm Github Repo này của mình, một repo dành cho những bạn muốn phát triển kĩ năng quản lí dự án bằng Excel. Mình đang giả định rằng bạn đã có một chút kiến thức và trải nghiệm cơ bản về Excel (chẳng hạn, bạn đã quen thuộc với các hàm IF, SUM, v.v.). Chúng ta sẽ cùng nhau làm quen với các tính năng nâng cao hơn của Excel như Table, Pivot Table, Filter, Sort, Conditional Formatting, Mail Merge, và một số hàm quan trọng cho công việc quản lí dự án như các hàm tìm kiếm và hàm thao tác trên số liệu.

# Danh sách bài học và Điểm chính

1. [Phần 1 (1.1-1.3): Table, VLOOKUP, XLOOKUP](https://youtu.be/3D5UvIGPPhM)
  - **Table** tự động mở rộng khi thêm dữ liệu vào dưới và bên phải, tự động cập nhật miền giá trị cho các công thức.
  - `VLOOKUP(giá_trị_cần_tìm, bảng_tham_chiếu, cột_trả_về, cách_khớp)`
    * Đặt `cách_khớp` = 0 để khớp chính xác (= 1 thì Excel sẽ khớp gần đúng).
    * Trong `bảng_tham_chiếu`, cột chứa `giá_trị_cần_tìm` luôn là cột thứ nhất (1).
    * Thứ tự của `cột_trả_về` không được tự động cập nhật khi bạn thay đổi cấu trúc bảng.
  - `XLOOKUP(giá_trị_cần_tìm, dữ_liệu_tham_chiếu, dữ_liệu_trả_về, kết_quả_trả_về_nếu_không_tìm_thấy, cách_khớp, cách_tìm_kiếm)` - khắc phục những nhược điểm của `VLOOKUP()`.
  - Nên thiết kế cơ sở dữ liệu theo normal forms.

