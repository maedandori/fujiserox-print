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

## Chạy Print Agent ở khay hệ thống (System Tray)

Từ phiên bản này, `Print-Agent.exe` là ứng dụng WinForms **WinExe**, không mở cửa sổ CMD khi chạy bình thường. Biểu tượng nằm trong khay hệ thống cạnh đồng hồ Windows (có thể nằm trong phần biểu tượng ẩn, bấm mũi tên ^). Chỉ một Agent được phép chạy trong mỗi phiên người dùng Windows.

Nhấn chuột phải vào biểu tượng để sử dụng:
- **Xem trạng thái**: hiển thị trạng thái kết nối/lệnh in gần nhất.
- **Mở cài đặt**: khởi chạy `Print-Settings.exe` trong cùng thư mục.
- **Xem nhật ký**: mở `print-agent.log` bằng Notepad.
- **Thoát Print Agent**: dừng nhận lệnh, để việc in đang xử lý hoàn tất trước khi kết thúc.

Cần đăng nhập Windows để điều khiển Excel 2013 trong phiên làm việc. Chức năng chạy cùng Windows **chưa tự bật** trong ứng dụng: người dùng có thể tạo shortcut của `Print-Agent.exe` trong thư mục `shell:startup` để chạy khi đăng nhập. Việc sử dụng System Tray không làm mất yêu cầu xác thực Agent Token của Cloudflare.

Chế độ `Print-Agent.exe --test` vẫn dùng `sample-job.json` và in trực tiếp, nhưng ứng dụng WinExe không xuất hiện console; để xem lỗi, chạy từ source bằng `dotnet run` hoặc kiểm tra bản in/hàng đợi Windows. Cần thử nghiệm thực tế với Windows 10 và Excel 2013.


## Nhiều máy Windows / định danh Agent

Bản Print-Agent.exe mới tạo `agent-id.txt` chứa mã UUID riêng trong thư mục chương trình, kèm tên Windows (`Environment.MachineName`) và tên máy in khi gửi heartbeat. Agent tự báo trạng thái mỗi khoảng 60 giây và gửi header `X-Print-Agent-Id` trong các request lấy/xác nhận lệnh in.

**Mỗi máy phải có một agent-id.txt riêng.** Khi sao chép hai EXE và `config.ini` sang Windows khác, KHÔNG sao chép `agent-id.txt` từ máy trước. Khởi chạy Agent trên máy mới để tự tạo ID khác. Không xóa `agent-id.txt` của máy đang sử dụng vì các lệnh chờ đã được gán theo ID.

Machine-Control sẽ không cấp lệnh in cho Agent đời cũ không gửi Agent ID. Cần tải Print-Agent.exe từ artifact build mới và thay trên *tất cả* máy Windows muốn sử dụng.

Trong trang 部材印, chọn tên máy Windows và tên máy in trước khi gửi thử. API báo `submitted` nghĩa là đã gửi tới Windows spooler, không chứng minh giấy đã được in thành công. Nếu Agent đã nhận lệnh nhưng mất kết nối, lệnh sẽ để trạng thái `claimed` để tránh tự động in trùng.


## Chế độ xuất PDF tự động (Excel 2013)

Print-Settings.exe có lựa chọn **In ra máy in** hoặc **Xuất PDF tự động**, cùng nút chọn thư mục lưu PDF. Khi chọn PDF, chương trình ghi `PrintMode=pdf` và `PdfOutputPath=C:\\...\\PDF` vào `config.ini`. Chế độ cũ `PrintMode=printer` vẫn giữ nguyên; INI cũ không có hai khóa này sẽ mặc định in giấy.

Khi Agent nhận lệnh, nó dùng Excel 2013 COM để tạo **một file PDF nhiều trang**, 8 thẻ/trang theo mẫu `原紙`. Bản gốc XLSX chỉ được mở ReadOnly, không lưu thay đổi. Nếu yêu cầu `copies>1`, file PDF sẽ chứa các trang lặp lại tương ứng với số bản sao, giữ nguyên số thứ tự trên từng bản sao. File được đặt tên theo thời gian, mã job và chuỗi ngẫu nhiên để tránh ghi đè.

PDF được xuất trực tiếp bằng `Workbook.ExportAsFixedFormat`, **không sử dụng Microsoft Print to PDF**, vì vậy không hiện hộp thoại Save As, không phụ thuộc máy in PDF được chọn. Thư mục phải tồn tại hoặc Agent có quyền tạo và ghi. Chế độ này vẫn dùng Agent ID, Agent Token, heartbeat và hàng đợi in Machine-Control như trước; không cần sửa API web để xuất PDF tại Windows.

Biểu tượng khay hệ thống của Agent đã thay bằng **hình máy in**. Sau khi thay hai EXE, chạy Print-Settings để chọn chế độ, lưu cấu hình rồi thoát/mở lại Agent từ System Tray. **Giữ nguyên `config.ini` và `agent-id.txt` cũ; không xóa Agent ID**.

Chưa thử nghiệm PDF thực tế với Excel 2013: cần kiểm tra file PDF đầu tiên, số trang, vùng in, font và cách Excel nhân bản sheet trong template thật. Nếu file không sinh ra, xem `print-agent.log` và đường dẫn thư mục đã cài.

## Tên hiển thị máy in trên Machine-Control

Trong Print-Settings có trường **Tên hiển thị trên web** (ví dụ `Máy in phòng vật tư`). Giá trị lưu dưới khóa `PrinterDisplayName` của `config.ini`, hoàn toàn tách biệt với `PrinterName` là tên máy in/driver Windows dùng để in thực sự.

Agent gửi `printerDisplayName` cùng `agentId`, `machineName`, `printerName` qua heartbeat. Trên trang 部材印, tên tùy chỉnh được hiển thị nổi bật; tên Windows/driver vẫn nằm ở dòng thông tin phụ để đối chiếu. Không thay đổi tên PC, driver hoặc Agent ID. Nếu chưa đặt tên, website dùng tên driver/máy Windows cũ.

Sau khi đổi tên, bấm **Lưu cấu hình**, thoát và mở lại Print Agent để nhận cấu hình mới. Agent gửi heartbeat mỗi 60 giây; nhấn **接続確認** trên web sau khi Agent đã cập nhật trạng thái. Không sao chép `agent-id.txt` giữa các PC.


## Version và tự cập nhật qua GitHub Releases riêng tư

Hai EXE được GitHub Actions build chung với phiên bản \`1.0.<workflow_run_number>\`. Build thành công sẽ phát hành GitHub Release riêng tư với tag \`print-v1.0.N\`, gói \`Machine-Control-Print-Update.zip\` (chỉ gồm **Print-Agent.exe** và **Print-Settings.exe**) và \`update-manifest.json\` (phiên bản, kích thước, SHA-256). Vẫn có ZIP artifact thông thường để tải ban đầu.

**Lý do cần Cloudflare:** GitHub repository \`maedandori/fujiserox-print\` là private. Máy Windows chỉ dùng \`PRINT_AGENT_TOKEN\` hiện có để gọi hai API đã được thêm trên Worker Machine-Control:
- \`GET /api/print-agent/update/latest\`: trả số phiên bản mới nhất, kích thước gói và SHA-256.
- \`GET /api/print-agent/update/package?version=1.0.N\`: Worker chuyển tiếp bản ZIP từ GitHub Release, không đưa GitHub Token xuống Windows.

**Cấu hình duy nhất cần bổ sung trên Cloudflare:** Tạo fine-grained Personal Access Token tại GitHub, chỉ cho phép truy cập repository \`maedandori/fujiserox-print\`, quyền **Contents: Read-only**, rồi lưu token làm Cloudflare Worker Secret \`GITHUB_PRINT_RELEASE_TOKEN\` trên account chứa Worker \`machine-control\`. Không nhập token GitHub vào Print-Settings, không commit token vào repo. Không thay \`PRINT_AGENT_TOKEN\` hiện có.

**Tự cập nhật:** Agent kiểm tra lúc hàng đợi trống và mỗi 6 giờ sau đó. Khi có phiên bản cao hơn, nó tải ZIP qua HTTPS, so sánh dung lượng và SHA-256, kiểm tra chỉ có đúng hai EXE, rồi chạy tiến trình Windows PowerShell ẩn để chờ Agent thoát, thay hai EXE và khởi động lại Agent. Không tải/ghi khi Agent đang xử lý Excel. Nếu mạng lỗi, token proxy chưa cấu hình hoặc SHA không khớp, Agent vẫn chạy bản cũ và tiếp tục in. Tên máy, template và token lưu trong \`config.ini\`, mã định danh máy lưu \`agent-id.txt\`, **không bị thay đổi**.

**Giao diện:** Click chuột phải biểu tượng máy in System Tray để xem phiên bản hoặc chọn **Kiểm tra cập nhật**. Print-Settings hiển thị phiên bản ở tiêu đề; trang 部材印 hiển thị phiên bản báo từ từng Agent.

**Lưu ý bắt buộc:**
- Phải cài **một lần thủ công** bản EXE có AutoUpdate này; bản cũ không tự có cơ chế nâng cấp.
- Để hai EXE và \`config.ini\` ở thư mục Windows cho phép người dùng hiện tại ghi (ví dụ \`C:\Machine-Control-Print\`), tránh \`C:\Program Files\` nếu không có quyền ghi. Đóng Print-Settings trước khi cập nhật.
- Không hỗ trợ nâng cấp trong khi máy đang tắt, người dùng chưa đăng nhập hoặc mạng ngắt; Agent sẽ tự kiểm tra lại khi được mở.
- Bản đầu tiên vẫn cần xác minh thực tế trên Win10 và Excel 2013. Không tự nhận đã in giấy thành công khi chỉ gửi đến Windows Spooler.
