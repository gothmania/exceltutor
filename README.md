# Excel tutorial

Cảm ơn bạn đã ghé thăm Github Repo này của mình, một repo dành cho những bạn muốn phát triển kĩ năng quản lí dự án bằng Excel. Mình đang giả định rằng bạn đã có một chút kiến thức và trải nghiệm cơ bản về Excel (chẳng hạn, bạn đã quen thuộc với các hàm IF, SUM, v.v.). Chúng ta sẽ cùng nhau làm quen với các tính năng nâng cao hơn của Excel như Table, Pivot Table, Filter, Sort, Conditional Formatting, và một số hàm quan trọng cho công việc quản lí dự án như các hàm tìm kiếm và hàm thao tác trên số liệu.

# Danh sách bài học và Điểm chính

Playlist bài học: [YouTube KonTrymNon](https://www.youtube.com/playlist?list=PLia_N2qlp_r-0dRuYOVxC7ggd0vtdNpkM)

1. Bài 1 (1.1-1.3): Table, VLOOKUP, XLOOKUP
  - **Table** tự động mở rộng khi thêm dữ liệu vào dưới và bên phải, tự động cập nhật miền giá trị cho các công thức.
  - `VLOOKUP(giá_trị_cần_tìm, bảng_tham_chiếu, cột_trả_về, cách_khớp)`
    * Đặt `cách_khớp` = 0 để khớp chính xác (= 1 thì Excel sẽ khớp gần đúng).
    * Trong `bảng_tham_chiếu`, cột chứa `giá_trị_cần_tìm` luôn là cột thứ nhất (1).
    * Thứ tự của `cột_trả_về` không được tự động cập nhật khi bạn thay đổi cấu trúc bảng.
  - `XLOOKUP(giá_trị_cần_tìm, dữ_liệu_tham_chiếu, dữ_liệu_trả_về, kết_quả_trả_về_nếu_không_tìm_thấy, cách_khớp, cách_tìm_kiếm)` - khắc phục những nhược điểm của `VLOOKUP()`.
  - Nên thiết kế cơ sở dữ liệu theo normal forms.

2. Bài 1 (1.4): IF lồng, IFS, XLOOKUP
  - Các cách để rẽ nhánh nhiều điều kiện:
    * `IF` lồng: Nhiều hàm `IF` lồng vào trong nhau. Ví dụ: `IF(điều_kiện_1, giá_trị_True1, IF(điều_kiện_2, giá_trị_True2, giá_trị_False))`.
    * `IFS`: không cần phải lồng nhiều hàm `IF` vào nhau nữa. Cú pháp: `IFS(điều_kiện_1, giá_trị_True1, điều_kiện_2, giá_trị_True2, ..., TRUE, giá_trị_False)`.
    * `XLOOKUP`: sử dụng chế độ match không chính xác (match giá trị nhỏ hơn hoặc lớn hơn gần nhất).
  - Hàm `ISBLANK` giúp kiểm tra một ô có phải là ô trống hay không.

3. Bài 1 (1.5): SUMIFS, Pivot Table
  - `SUMIFS(miền_tính_tổng, miền_kiểm_tra_điều_kiện1, điều_kiện_cần_kiểm_tra1, miền_kiểm_tra_điều_kiện2, điều_kiện_cần_kiểm_tra2, ...)`
    * Tính tổng nhưng giới hạn cho các dòng thỏa mãn một số điều kiện nhất định
  - Pivot Table: công cụ tổng hợp tự động số liệu dựa trên số liệu thô, cho phép hiển thị số liệu tổng hợp được nhóm lại theo dòng, cột, và có thể dùng bộ lộc (filter).

4. Bài 2 (2.1): MAXIFS, Conditional Formatting
  - `MAXIFS` tương tự `SUMIFS` nhưng dùng để tìm giá trị lớn nhất
  - Conditional Formatting là công cụ giúp định dạng các ô tự động dựa trên điều kiện. Bạn không nên định dạng các ô trong Excel bằng tay, đặc biệt là các ô trong Table.
    * Một trong các tính năng nâng cao của Conditional Formatting là sử dụng công thức (formula) để kiểm tra điều kiện.

5. Bài 2 (2.2): Nối trường, TEXT, YEAR
  - Bạn có thể tạo ra một chuỗi kí tự mới bằng cách nối dữ liệu trong các trường với nhau. Bạn có thể sử dụng toán tử `&` hoặc hàm `CONCAT`.
  - Hàm `TEXT` giúp định dạng lại giá trị số thành dạng mà bạn mong muốn, ví dụ thêm các chữ số 0 vào đằng trước cho đủ 3 chữ số: `TEXT(A1, "000")`.
  - Hàm `YEAR` trích xuất năm từ dữ liệu kiểu ngày tháng. Các hàm khác như `MONTH`, `DAY`, v.v. cũng thực hiện chức năng tương tự.

6. Bài 2 (2.3): RIGHT, FILTER, SEARCH, SORT
  - `RIGHT(chuỗi_kí_tự, n)` cho phép cắt ra `n` kí tự ở cuối một chuỗi kí tự. Ngoài ra có hàm `LEFT` và `MID` thực hiện chức năng tương tự.
  - Hàm `FILTER` cho phép lọc ra một số các hàng dữ liệu thỏa mãn một điều kiện nào đó.
  - Hàm `SEARCH` trả về vị trí đầu tiên mà một chuỗi kí tự xuất hiện trong một chuỗi kí tự khác. Hàm sẽ trả về lỗi nếu như không tìm thấy. Chúng ta có thể thay lỗi này bằng giá trị 0 với hàm `IFERROR`.
  - Hàm `SORT` sắp xếp các hàng dữ liệu theo ABC hoặc theo điều kiện.

7. Bài 3 (3.1): Name, Data Validation
  - Để thuận tiện cho việc quản lí các miền dữ liệu, bạn nên đặt tên thay vì gọi thẳng địa chỉ của chúng ra. Khi nhìn công thức, người khác sẽ dễ dàng hiểu được mục đích của công thức này.
  - Dữ liệu nhập bằng tay nên được kiểm tra về độ chính xác (gọi là validation hay phê chuẩn). Excel cho phép phê chuẩn bằng nhiều điều kiện khác nhau, đặc biệt có thể phê chuẩn từ một danh sách (sẽ hiện ra danh sách thả xuống ở các ô được kiểm tra).

8. Bài 3 (3.2): Power Query
  - Power Query cho phép bạn thu thập số liệu từ nhiều nguồn, xử lí, điều chỉnh, kết hợp, và tổng hợp mà không phải lưu trữ tất cả các số liệu đó vào các sheet của Excel. Bạn chỉ cần lưu trữ nội dung cuối cùng mà bạn muốn hiển thị trong Excel (bảng, Pivot Table, Pivot Chart).

9. Bài 3 (3.3-3.4): Consolidate, Subtotal
  - Consolidate cho phép "chồng" nhiều bảng có cùng cấu trúc lên nhau để tạo ra một bảng tổng hợp số liệu. Ví dụ, bạn có thể thu thập dữ liệu có cùng nội dung từ nhiều công ty, sau đó dùng Consolidate để tạo ra bảng tổng hợp cho tất cả các công ty này.
  - Subtotal cho phép thực hiện các phép tính trên một miền mà bỏ qua các ô có hàm `SUBTOTAL` khác. Như vậy bạn có thể hiển thị cả các dữ liệu chi tiết và dữ liệu tổng hợp trên cùng một cột mà không sợ bị gộp vào khi tính toán.
