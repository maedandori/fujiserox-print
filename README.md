# Machine-Control: hai ứng dụng in Windows 10

Máy in: Fuji Xerox DocuCentre-VI C3371 (LAN). Excel: Office 2013. Không cần cài Node.js, .NET SDK hoặc .NET Runtime tại máy sử dụng; GitHub Actions đóng gói self-contained win-x64.

## Hai chương trình

- **Print-Settings.exe**: giao diện Windows để nhập URL Machine-Control, Agent Token, tên máy in, đường dẫn template .xlsx và chu kỳ lấy lệnh. Bấm **Lưu cấu hình** tạo/sửa **config.ini** cạnh hai file .exe.
- **Print-Agent.exe**: đọc **config.ini khi khởi động**, kết nối HTTPS đến Machine-Control để lấy lệnh, điền dữ liệu vào Excel và gửi in. Không ghi đè template. Phải khởi động lại Agent khi đổi cấu hình.

Hai EXE và config.ini phải đặt **cùng một thư mục**. Agent chạy trong phiên Windows đã đăng nhập, vì Microsoft Excel COM automation không phù hợp với Windows Service chạy ngầm dưới SYSTEM.

## Build từ iPhone bằng GitHub Actions

Vào **Actions > Build Machine-Control Print Apps > Run workflow**, chọn bản chạy thành công rồi tải artifact **Machine-Control-Two-Print-Apps-Win10-x64**. Artifact có hai EXE và sample-job.json.

Nếu workflow mới push vào main, GitHub Actions cũng tự khởi chạy.

## Sử dụng

1. Cài và kiểm tra driver Fuji Xerox C3371, in thử từ Windows. Khuyên đặt C3371 làm máy in mặc định.
2. Giải nén hai EXE vào ví dụ `C:\Machine-Control-Print`.
3. Chép `template.xlsx` vào thư mục đó hoặc chọn file XLSX từ đường dẫn khác.
4. Chạy `Print-Settings.exe`, chọn template và các cài đặt, nhấn **Lưu cấu hình**. Để trống tên máy in nếu muốn dùng máy in mặc định; Excel có thể cần tên máy in kèm cổng khi chỉ định trực tiếp.
5. Để in thử **không cần Cloudflare**, mở terminal trong thư mục và chạy `./Print-Agent.exe --test`. Điều chỉnh `sample-job.json` theo các ô thực trong template.
6. Khi backend Cloudflare đã có API và token hợp lệ, nhập API URL + Agent Token trong Settings, rồi chạy `Print-Agent.exe` để nhận lệnh liên tục. Nhấn Ctrl+C để dừng.

## config.ini

Ví dụ sau là cấu hình minh hoạ (không commit file config.ini chứa token):

```ini
[PrintAgent]
ApiBaseUrl=https://machine-control.maedandoris.workers.dev
AgentToken=
PrinterName=
TemplatePath=template.xlsx
PollSeconds=3
```

Tệp INI nằm cạnh EXE; bảo vệ quyền truy cập thư mục vì AgentToken lưu trong văn bản thuần. Nếu di chuyển EXE sang thư mục khác, hãy chuyển config.ini cùng nó.

## Hợp đồng API dự kiến (chưa xác minh đã triển khai)

- `GET /api/print-agent/jobs/next`: HTTP 204 khi không có job; hoặc JSON có `jobId`, `copies`, `cells` (địa chỉ `Sheet1!B3` -> nội dung).
- `POST /api/print-agent/jobs/{jobId}/status`: `{"status":"submitted","error":null}` hoặc `{"status":"failed","error":"..."}`.
- Header `Authorization: Bearer <AgentToken>` phải được xác thực tại Cloudflare. Backend cần bảo đảm chỉ một Agent claim một job để tránh in trùng.

**Trạng thái:** đã bổ sung mã nguồn cho hai EXE và workflow. Chưa khẳng định workflow build thành công hay đã thử trên Windows thật. Chưa tích hợp API phía Machine-Control, chưa có template XLSX thực tế để xác minh nội dung ô. Excel PrintOut xác nhận gửi vào spooler, không chứng minh giấy đã ra khỏi máy in.


## Template 部材カード 原紙 — ánh xạ 8 thẻ

Đã kiểm tra file mẫu `部材カード　原紙.xlsx` người dùng cung cấp: sheet tên **原紙**, các ô ghi là **ô đầu vùng gộp**. Mỗi lệnh có tối đa tám thẻ, theo thứ tự trái sang phải, hàng trên rồi hàng dưới:

| Thẻ | Tên | Màu | Số lượng |
| --- | --- | --- | --- |
| 1 | B8 | B12 | B16 |
| 2 | H8 | H12 | H16 |
| 3 | N8 | N12 | N16 |
| 4 | T8 | T12 | T16 |
| 5 | B32 | B36 | B40 |
| 6 | H32 | H36 | H40 |
| 7 | N32 | N36 | N40 |
| 8 | T32 | T36 | T40 |

Ghi chú: vị trí `T816` được hiểu thành `T16`. Khi job có `cards`, Agent tự ghi tên/màu/số lượng theo bảng và xoá ba trường của các thẻ còn trống trong bản in tạm. Job vẫn hỗ trợ `cells` kiểu cũ; nếu cùng một ô xuất hiện ở cả hai phần, giá trị `cells` ghi đè mapping tự động.

Lệnh mẫu: `{"jobId":"123","copies":1,"cards":[{"name":"部材 A","color":"白","quantity":"100"}],"cells":{}}`.

**Lưu ý:** Template XLSX đính kèm trong chat chưa được tự động đưa lên GitHub (repo chỉ chứa mã nguồn). Hãy sao chép file mẫu lên máy Windows và chọn đường dẫn bằng Print-Settings.exe. Cần kiểm tra in giấy thật để chắc chắn căn lề, kích thước trang và máy in phù hợp.

## Đánh số thứ tự thẻ và in nhiều trang

Print Agent tự đánh số lần lượt cho từng thẻ: `E20`, `K20`, `Q20`, `W20`, `E44`, `K44`, `Q44`, `W44`. Các ô tương ứng với thẻ 1–8. Nếu có 10 thẻ, trang 1 có số 1–8 và trang 2 có số 9–10, sáu vị trí còn lại trống. Mỗi trang mở template mới ở chế độ ReadOnly, không sửa file XLSX gốc.

Trong JSON, dùng `cards` là danh sách toàn bộ thẻ (tối đa 800), `startNumber` mặc định bằng 1, có thể đặt ví dụ 101 để bắt đầu từ số 101. `copies` là số bản sao của mỗi trang; số thứ tự **không tăng** giữa các bản sao cùng một trang. Nếu không cung cấp `cards`, Agent giữ chế độ cũ `cells` và không tự đánh số.

Lưu ý: Với các lệnh nhiều trang, nếu gặp lỗi giữa chừng, một số trang có thể đã được gửi tới spooler. Cần đối chiếu trước khi in lại để tránh in trùng.
