# TV360 Schedule Formatter

Tool nhỏ để chuẩn hóa nhiều file text hoặc Excel lịch phát sóng TV360 thành một file Excel.

Đầu vào là một hoặc nhiều file `.txt` hoặc `.xlsx` có format chưa đồng nhất. Tool sẽ tự nhận các dạng giờ phổ biến như:

- `06:00 Chào buổi sáng`
- `06h30 - Phim truyện Việt Nam`
- `07.15: Bản tin`
- `08:00-08:30 Thiếu nhi`
- `09h` rồi tiêu đề ở dòng kế tiếp

Đầu ra mặc định là file `.xlsx` có các cột:

- `STT`
- `Thứ ngày`
- `Tiêu đề chương trình`
- `Thời gian phát sóng`
- `Nguồn file`
- `Dòng`
- `Dòng gốc`

## Cách dùng

### Chạy trên web local

```bash
cd /Users/mrk/tv360-schedule-formatter
python3 web_app.py
```

Sau đó mở:

```text
http://127.0.0.1:8765
```

Trên web app, chọn nhiều file `.txt` hoặc `.xlsx`, bấm `Xem trước` để kiểm tra kết quả parse, rồi bấm `Xuất Excel` để tải file `.xlsx`.

Nếu muốn đổi port:

```bash
python3 web_app.py --port 9000
```

### Chạy bằng command line

```bash
cd /Users/mrk/tv360-schedule-formatter
python3 format_tv360_schedule.py samples -o output/lich_tv360.xlsx
```

Chạy với nhiều file cụ thể, có thể trộn `.txt` và `.xlsx`:

```bash
python3 format_tv360_schedule.py lich_vtv1.txt lich_vtv3.xlsx -o output/lich_tv360.xlsx
```

Chỉ xuất các cột chính `STT`, `Thứ ngày`, `Tiêu đề chương trình`, `Thời gian phát sóng`:

```bash
python3 format_tv360_schedule.py samples -o output/lich_tv360_minimal.xlsx --minimal
```

Sắp xếp theo giờ phát sóng:

```bash
python3 format_tv360_schedule.py samples -o output/lich_tv360.xlsx --sort
```

Kiểm tra parser trước khi xuất Excel:

```bash
python3 format_tv360_schedule.py samples --dry-run
```

## Ghi chú

- Tool không tự dịch hoặc sửa chính tả tiêu đề chương trình, vì tiêu đề tiếng Việt nên được giữ đúng theo nguồn.
- Text nằm trong ngoặc như `(HD)`, `[Tập 12]`, `{Live}` sẽ bị loại khỏi tiêu đề khi xuất Excel.
- Tiêu đề chương trình trong output được chuẩn hóa kiểu chỉ viết hoa chữ cái đầu câu.
- Tên chương trình và số tập/số phát sóng được ngăn cách bằng dấu `-`, ví dụ `Phim tài liệu - tập 13`.
- Với input `.xlsx`, tool đọc các sheet trong workbook, ghép nội dung từng hàng thành một dòng để parse.
- Nếu ô giờ trong Excel là dạng time number, ví dụ `0.25`, tool sẽ đổi thành `06:00`.
- Với input `.xlsx` dạng bảng tuần, tool lấy `Thứ ngày` từ tiêu đề cột như `Thứ Hai (25/5)`.
- Nếu một dòng chỉ có giờ, tool sẽ lấy dòng text kế tiếp làm tiêu đề.
- Nếu file text có encoding khác UTF-8, tool sẽ thử thêm `utf-16`, `cp1258`, và `latin-1`.
