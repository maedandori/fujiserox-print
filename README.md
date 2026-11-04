# Machine-Control Print Agent — Fuji Xerox C3371

Ứng dụng C# dành cho **Windows 10 64-bit + Microsoft Excel 2013 + Fuji Xerox DocuCentre-VI C3371 qua LAN**.

## Trạng thái

- Đã thêm source C# và workflow GitHub Actions để build **EXE self-contained**.
- Có chế độ in thử bằng file XLSX cục bộ. **Chưa có template XLSX thực tế**, vì vậy chưa thể xác minh bố cục và in thử trên C3371.
- Chế độ nhận lệnh qua Cloudflare mới là **client theo hợp đồng API bên dưới**, chưa triển khai các API phía Machine-Control. Không được coi là đã tích hợp hoặc deploy.
- Chưa kiểm tra build thành công và chưa xác minh trên Windows thật.

## Xây dựng trên GitHub Actions

Trong repo này, vào **Actions → Build Machine-Control Print Agent → Run workflow**. Sau khi job thành công, mở lần chạy và tải artifact **Machine-Control-Print-Agent-Win10-x64**.

GitHub Actions tự build trên Windows runner bằng .NET SDK. Máy tính dùng để in **không cần cài Node.js hoặc .NET SDK/runtime riêng**.

## Cài và in thử tại Windows 10

1. Cài đúng driver Fuji Xerox/FUJIFILM C3371; in thử từ Windows. Đặt **C3371 làm máy in mặc định**.
2. Giải nén artifact vào thư mục riêng, ví dụ `C:\Machine-Control-Print-Agent`.
3. Sao chép `config.example.json` thành `config.json`.
4. Chép file mẫu Excel thật thành `template.xlsx` trong cùng thư mục (hoặc sửa `TemplatePath` trỏ tới đường dẫn tuyệt đối).
5. Sửa `sample-job.json`: khóa dạng `Sheet1!B3` phải khớp **tên sheet** và **địa chỉ ô có thật** trong template. Không được ghi vào ô gộp không phải ô đầu vùng.
6. Chạy **trong phiên Windows đã đăng nhập** (có thể mở PowerShell từ thư mục Agent):

   `./Machine-Control-Print-Agent.exe --test`

7. Agent mở một **bản Excel chỉ đọc**, điền dữ liệu trong bộ nhớ và gọi `Workbook.PrintOut`. **Không sửa file template gốc**. Kết quả 'đã gửi tới Windows Spooler' chưa có nghĩa giấy đã in xong.

Mặc định `PrinterName` để trống để Excel sử dụng máy in mặc định Windows; nếu chỉ định tên máy in, Excel có thể yêu cầu tên bao gồm cổng, ví dụ `... on Ne01:`.

Máy tính cần có **Excel 2013 cài sẵn**. Đây là ứng dụng điều khiển Excel trong phiên desktop; **không chạy bằng tài khoản SYSTEM hoặc dưới dạng dịch vụ Windows không tương tác**. Excel có thể hiển thị hộp thoại hoặc bị treo khi file mẫu gặp lỗi, cần thử nghiệm thực tế trước khi tự động hóa.

## API dự kiến cho Machine-Control (chưa triển khai)

Agent chủ động lấy lệnh, dùng `Authorization: Bearer <AgentToken>`:

- `GET /api/print-agent/jobs/next`: trả HTTP 204 nếu không có việc; nếu có, trả JSON:
  ```json
  {
    "jobId": "unique-job-id",
    "copies": 1,
    "cells": {"Sheet1!B3": "Tên khuôn", "Sheet1!E3": "15号機"}
  }
  ```
- `POST /api/print-agent/jobs/{jobId}/status`: JSON `{"status":"submitted","error":null}` hoặc `{"status":"failed","error":"..."}`.

Cloudflare phải xác thực token, **claim mỗi lệnh cho đúng một Agent**, kiểm tra trạng thái, ngăn in lại lệnh cũ và thiết kế cơ chế xử lý khi mất kết nối. Không tự động gửi lại lệnh `submitted` khi chưa đối chiếu với hàng đợi Windows, vì có nguy cơ in trùng.

**Chưa có khả năng chứng minh 'đã in xong' bằng phương thức Excel.PrintOut**; trước mắt chỉ báo trạng thái gửi thành công tới Windows.

## Cấu hình

`ApiBaseUrl`, `AgentToken` trong `config.json` để trống cho đến khi API thực sự triển khai. Không commit `config.json` có token.

`TemplatePath` là đường dẫn file mẫu trong máy Windows, `PollSeconds` mặc định 3 giây. Chế độ `--test` không gửi yêu cầu đến Cloudflare.

## Các việc tiếp theo

1. Cung cấp `template.xlsx` để xác minh tên sheet, các ô thay thế, vùng in, khổ giấy và hướng giấy.
2. Build workflow trên GitHub; kiểm tra lỗi nếu có.
3. In thử trực tiếp trên Windows 10.
4. Sau đó mới thêm API D1 và nút In trong repo `machine-control`, cấp token riêng cho Agent và thử end-to-end.
