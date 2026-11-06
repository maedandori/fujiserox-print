using System.Net;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;

internal static class IniConfig
{
    public static AgentConfig Load(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("Thiếu config.ini. Chạy Print-Settings.exe trước.", path);
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var raw in File.ReadAllLines(path, System.Text.Encoding.UTF8))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith(';') || line.StartsWith('#') || line.StartsWith('[')) continue;
            int i = line.IndexOf('=');
            if (i > 0) values[line[..i].Trim()] = line[(i + 1)..].Trim();
        }
        string Get(string key, string fallback = "") => values.TryGetValue(key, out var v) ? v : fallback;
        if (!int.TryParse(Get("PollSeconds", "3"), out int seconds))
            throw new InvalidDataException("PollSeconds phải là số nguyên.");
        if (!int.TryParse(Get("MaxPdfPages", "96"), out int maxPdfPages) || maxPdfPages < 1 || maxPdfPages > 100)
            throw new InvalidDataException("MaxPdfPages phải nằm trong khoảng 1–100 trang.");
        return new AgentConfig { ApiBaseUrl = Get("ApiBaseUrl"), AgentToken = Get("AgentToken"),
            PrinterName = Get("PrinterName"), TemplatePath = Get("TemplatePath", "template.xlsx"), PollSeconds = seconds, MaxPdfPages = maxPdfPages,
            PrintMode = Get("PrintMode", "pdfprint"), PdfOutputPath = Get("PdfOutputPath"), PdfPrinterExe = Get("PdfPrinterExe", "SumatraPDF.exe"), PrinterDisplayName = Get("PrinterDisplayName"), AutoUpdateEnabled = Get("AutoUpdateEnabled","1")!="0" };
    }
}
internal sealed class AgentConfig
{
    public string ApiBaseUrl { get; set; } = "";
    public string AgentToken { get; set; } = "";
    public string PrinterName { get; set; } = "";
    public string PrinterDisplayName { get; set; } = "";
    public string TemplatePath { get; set; } = "template.xlsx";
    public int PollSeconds { get; set; } = 3;
    public int MaxPdfPages { get; set; } = 96;
    public string PrintMode { get; set; } = "pdfprint";
    public string PdfPrinterExe { get; set; } = "SumatraPDF.exe";
    public string PdfOutputPath { get; set; } = "";
    public bool AutoUpdateEnabled { get; set; } = true;
}
internal sealed class PrintCard
{
    public string Name { get; set; } = "";
    public string Color { get; set; } = "";
    public string Quantity { get; set; } = "";
    public bool IsFinal { get; set; } = false;
}
internal sealed class PrintJob
{
    public string JobId { get; set; } = "";
    public int Copies { get; set; } = 1;
    public List<PrintCard>? Cards { get; set; }
    public int StartNumber { get; set; } = 1;
    public Dictionary<string, string?> Cells { get; set; } = new();
}
internal static class Program
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private static readonly Regex AddressPattern = new(@"^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static string Root => AppContext.BaseDirectory;

    [STAThread]
    private static int Main(string[] args)
    {
        System.Windows.Forms.Application.EnableVisualStyles();
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
        using var instance = new System.Threading.Mutex(true, @"Local\MachineControlPrintAgent", out bool first);
        if (!first)
        {
            System.Windows.Forms.MessageBox.Show("Print Agent đang chạy trong khay hệ thống.", "Machine-Control");
            return 0;
        }
        if (args.Contains("--test", StringComparer.OrdinalIgnoreCase))
            return RunAgent(args, CancellationToken.None, _ => {});
        System.Windows.Forms.Application.Run(new TrayAgentContext());
        return 0;
    }

    internal static int RunAgent(string[] args, CancellationToken cancel, Action<string> status)
    {
        status("Đang khởi động");
        Console.WriteLine("Machine-Control Print Agent v"+AutoUpdate.Version+" - Windows 10 / Excel 2013");
        Console.WriteLine("Chạy trong phiên Windows có người dùng đăng nhập, không chạy dưới SYSTEM.");
        try
        {
            var config = IniConfig.Load(Path.Combine(Root, "config.ini"));
            if (args.Contains("--test", StringComparer.OrdinalIgnoreCase))
            {
                var testPath = Path.Combine(Root, "sample-job.json");
                var job = ReadJob(File.ReadAllText(testPath));
                ProcessPrintJob(config, job);
                Console.WriteLine(config.PrintMode.Equals("pdf", StringComparison.OrdinalIgnoreCase) ? "Đã tạo file PDF thử." : "Đã gửi PDF thử tới Windows Print Spooler.");
                return 0;
            }
            if (string.IsNullOrWhiteSpace(config.ApiBaseUrl) || string.IsNullOrWhiteSpace(config.AgentToken))
                throw new InvalidOperationException("Chưa cấu hình API Cloudflare và AgentToken. Dùng --test để in thử tại máy.");
            if (!Uri.TryCreate(config.ApiBaseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttps && !(baseUri.IsLoopback && baseUri.Scheme == Uri.UriSchemeHttp)))
                throw new InvalidOperationException("ApiBaseUrl phải là HTTPS (ngoại lệ HTTP localhost để kiểm thử).");
            if (config.PollSeconds < 2 || config.PollSeconds > 300)
                throw new InvalidOperationException("PollSeconds phải nằm trong khoảng 2–300.");

            using var client = new HttpClient { BaseAddress = new Uri(config.ApiBaseUrl.TrimEnd('/') + "/"), Timeout = TimeSpan.FromSeconds(30) };
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", config.AgentToken);
            var agentId = AgentIdentity.LoadOrCreate();
            client.DefaultRequestHeaders.Add("X-Print-Agent-Id", agentId);
            status("Agent: " + Environment.MachineName + " (" + agentId[..8] + ")");
            status("Đang kết nối với Machine-Control");
            var lastHeartbeat = DateTime.MinValue;
            var nextUpdateCheck = DateTime.UtcNow;
            var nextRemoteUpdateCheck = DateTime.UtcNow;
            while (!cancel.IsCancellationRequested)
            {
                if (DateTime.UtcNow - lastHeartbeat >= TimeSpan.FromSeconds(60))
                {
                    lastHeartbeat = DateTime.UtcNow;
                    try { PrinterHeartbeat.Send(client, config, agentId); }
                    catch (Exception ex) { Console.Error.WriteLine("Không gửi được trạng thái máy in: " + ex.Message); }
                }
                try
                {
                    using var response = client.GetAsync("api/print-agent/jobs/next").GetAwaiter().GetResult();
                    if (response.StatusCode == HttpStatusCode.NoContent)
                    {
                        // Check only while the queue is empty: never interrupt Excel/printing.
                        if (AutoUpdate.ConsumeRequest() || (config.AutoUpdateEnabled && DateTime.UtcNow >= nextUpdateCheck))
                        {
                            nextUpdateCheck = DateTime.UtcNow.AddHours(6);
                            if(AutoUpdate.TryStage(client,status,cancel))return 0;
                        }
                        // Lệnh Admin chỉ xử lý khi hàng đợi in trống; kiểm tra tối đa 30 giây/lần.
                        if (DateTime.UtcNow >= nextRemoteUpdateCheck)
                        {
                            nextRemoteUpdateCheck = DateTime.UtcNow.AddSeconds(30);
                            try
                            {
                                using var commandResponse=client.GetAsync("api/print-agent/update/command").GetAwaiter().GetResult();
                                if(commandResponse.IsSuccessStatusCode)
                                {
                                    using var commandJson=JsonDocument.Parse(commandResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult());
                                    if(commandJson.RootElement.TryGetProperty("update",out var needed)&&needed.ValueKind==JsonValueKind.True)
                                    {
                                        status("Admin yêu cầu cập nhật Print Agent...");
                                        var started=AutoUpdate.TryStage(client,status,cancel);
                                        ReportRemoteUpdate(client,started?"applying":"failed",started?null:"Không thể cập nhật. Kiểm tra log Agent và GitHub Release.");
                                        if(started)return 0;
                                    }
                                }
                                else Console.Error.WriteLine("Không kiểm tra được lệnh cập nhật web HTTP "+(int)commandResponse.StatusCode);
                            }
                            catch(Exception ex){Console.Error.WriteLine("Lỗi kiểm tra cập nhật từ Admin: "+ex.Message);}
                        }
                        status("Đã kết nối - chờ lệnh in (v"+AutoUpdate.Version+")");
                        if (cancel.WaitHandle.WaitOne(TimeSpan.FromSeconds(config.PollSeconds))) break;
                        continue;
                    }
                    response.EnsureSuccessStatusCode();
                    var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    var job = ReadJob(json);
                    status("Đang in: " + job.JobId);
                    Console.WriteLine($"Nhận yêu cầu in: {job.JobId}");
                    try
                    {
                        ProcessPrintJob(config, job);
                        Report(client, job.JobId, "submitted", null);
                        status((config.PrintMode.Equals("pdf", StringComparison.OrdinalIgnoreCase) ? "Đã tạo PDF: " : "Đã gửi PDF vào hàng đợi Windows: ") + job.JobId);
                        Console.WriteLine(config.PrintMode.Equals("pdf", StringComparison.OrdinalIgnoreCase)
                            ? $"Đã xuất PDF từ lệnh: {job.JobId}" : $"Đã gửi đến hàng đợi in Windows: {job.JobId}");
                    }
                    catch (Exception error)
                    {
                        status("Lỗi in: " + error.Message);
                        Console.Error.WriteLine("Lỗi in: " + error.Message);
                        Report(client, job.JobId, "failed", error.Message);
                    }
                }
                catch (Exception error)
                {
                    status("Mất kết nối: " + error.Message);
                    Console.Error.WriteLine("Không lấy được lệnh in: " + error.Message);
                }
                if (cancel.WaitHandle.WaitOne(TimeSpan.FromSeconds(config.PollSeconds))) break;
            }
            return 0;
        }
        catch (Exception error)
        {
            status("Lỗi cấu hình: " + error.Message);
            Console.Error.WriteLine("Lỗi cấu hình: " + error.Message);
            return 1;
        }
    }

    private static void ReportRemoteUpdate(HttpClient client,string state,string? error)
    {
        try
        {
            using var result=client.PostAsJsonAsync("api/print-agent/update/command/result",new {status=state,error}).GetAwaiter().GetResult();
            result.EnsureSuccessStatusCode();
        }
        catch(Exception ex){Console.Error.WriteLine("Không báo được trạng thái cập nhật: "+ex.Message);}
    }

    private static PrintJob ReadJob(string json)
    {
        var job = JsonSerializer.Deserialize<PrintJob>(json, JsonOptions)
            ?? throw new InvalidDataException("Dữ liệu lệnh in trống.");
        if (string.IsNullOrWhiteSpace(job.JobId) || job.JobId.Length > 128 || job.Copies < 1 || job.Copies > 100 ||
            job.Cells is null || job.Cells.Count > 100 || job.StartNumber < 1 || job.StartNumber > 1000000 || (job.Cards?.Count ?? 0) > 800)
            throw new InvalidDataException("JobId, Copies hoặc Cells không hợp lệ.");
        return job;
    }

    private static void Report(HttpClient client, string id, string status, string? error)
    {
        try
        {
            // Backend cần xác minh job đã được claim duy nhất trước khi cấp cho Agent.
            var path = "api/print-agent/jobs/" + Uri.EscapeDataString(id) + "/status";
            using var response = client.PostAsJsonAsync(path, new { status, error }).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine("Không báo được trạng thái lên Cloudflare: " + exception.Message);
            Console.Error.WriteLine("Không tự in lại lệnh này để tránh in trùng. Cần đối chiếu trên máy tính.");
        }
    }

    private static void ProcessPrintJob(AgentConfig config, PrintJob job)
    {
        // Mặc định mới: điền XLSX -> xuất PDF -> in PDF; không gọi Excel.PrintOut.
        var template = Path.IsPathRooted(config.TemplatePath) ? config.TemplatePath : Path.Combine(Root, config.TemplatePath);
        template = Path.GetFullPath(template);
        if (!File.Exists(template) || !template.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new FileNotFoundException("Không tìm thấy template XLSX: " + template);
        var mode = config.PrintMode.Trim().ToLowerInvariant();
        // Chế độ chỉ tạo PDF dành cho --test; job Cloudflare không được báo là "đã in".
        if (mode == "pdf" && !Environment.GetCommandLineArgs().Contains("--test", StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("PrintMode=pdf chỉ xuất file, không in. Chọn 'Xuất PDF rồi in' trong Print-Settings để nhận job Cloudflare.");
        if (mode == "printer")
        {
            // Tự chuyển cấu hình INI cũ từ in Excel trực tiếp sang quy trình PDF.
            Console.WriteLine("Đã chuyển PrintMode=printer cũ sang xuất PDF rồi in PDF.");
            mode = "pdfprint";
        }
        if (mode != "pdf" && mode != "pdfprint")
            throw new InvalidDataException("PrintMode phải là pdfprint hoặc pdf.");
        // Kiểm tra bộ in PDF trước khi xuất để tránh tạo file mà không thể in.
        string? pdfPrinter = mode == "pdfprint" ? ResolvePdfPrinter(config.PdfPrinterExe) : null;
        // Tách PDF theo MaxPdfPages trong config.ini (mặc định 96 trang, giới hạn 1–100).
        // Chạy tuần tự, không gửi các PDF cùng lúc. Nếu lỗi giữa chừng KHÔNG tự in lại.
        int cardsPerPdf = checked(config.MaxPdfPages * 8);
        int count = job.Cards?.Count ?? 0;
        int chunks = Math.Max(1, (count + cardsPerPdf - 1) / cardsPerPdf);
        for (int copy = 0; copy < job.Copies; copy++)
        {
            for (int chunk = 0; chunk < chunks; chunk++)
            {
                int offset = chunk * cardsPerPdf;
                var part = new PrintJob
                {
                    JobId = job.JobId + "_part_" + (chunk + 1) + "_copy_" + (copy + 1),
                    Copies = 1,
                    Cards = count == 0 ? null : job.Cards!.GetRange(offset, Math.Min(cardsPerPdf, count - offset)),
                    StartNumber = checked(job.StartNumber + offset),
                    Cells = job.Cells
                };
                var pdfPath = ExportPdf(config, part, template);
                if (mode == "pdfprint")
                {
                    PrintPdf(pdfPrinter!, config.PrinterName, pdfPath);
                    Console.WriteLine($"Đã gửi PDF {chunk + 1}/{chunks} (bản {copy + 1}/{job.Copies}) vào Windows Spooler.");
                }
            }
        }
    }

    private static string ResolvePdfPrinter(string configured)
    {
        if (string.IsNullOrWhiteSpace(configured)) configured = "SumatraPDF.exe";
        string[] candidates = Path.IsPathRooted(configured)
            ? new[] { configured }
            : new[] {
                Path.Combine(Root, configured),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "SumatraPDF", configured),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "SumatraPDF", configured),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SumatraPDF", configured)
            };
        foreach (var candidate in candidates)
            if (File.Exists(candidate) && candidate.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                return Path.GetFullPath(candidate);
        throw new FileNotFoundException("Không tìm thấy SumatraPDF.exe. Chọn đường dẫn EXE trong Print-Settings hoặc đặt SumatraPDF.exe cạnh Print-Agent.exe.");
    }

    private static void PrintPdf(string executable, string printerName, string pdfPath)
    {
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length < 5)
            throw new IOException("File PDF vừa xuất không hợp lệ: " + pdfPath);
        using var file = File.OpenRead(pdfPath);
        Span<byte> signature = stackalloc byte[5];
        if (file.Read(signature) != 5 || !signature.SequenceEqual("%PDF-"u8))
            throw new InvalidDataException("File xuất không có định dạng PDF.");
        file.Close();
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Root };
        if (string.IsNullOrWhiteSpace(printerName)) start.ArgumentList.Add("-print-to-default");
        else { start.ArgumentList.Add("-print-to"); start.ArgumentList.Add(printerName); }
        start.ArgumentList.Add("-print-settings");
        start.ArgumentList.Add("noscale,simplex");
        start.ArgumentList.Add("-silent");
        start.ArgumentList.Add(pdfPath);
        using var process = Process.Start(start) ?? throw new InvalidOperationException("Không thể mở bộ in PDF.");
        if (!process.WaitForExit(120000))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("Bộ in PDF quá thời gian 120 giây. Hãy đối chiếu hàng đợi in Windows trước khi thử lại để tránh trùng.");
        }
        if (process.ExitCode != 0)
            throw new IOException("Bộ in PDF trả mã lỗi " + process.ExitCode + ". Kiểm tra tên máy in, driver và hàng đợi Windows; không tự in lại.");
        Console.WriteLine("Đã gửi PDF tới Windows Print Spooler: " + pdfPath);
    }

    private static void PrintExcelLegacy(AgentConfig config, PrintJob job)
    {
        var template = Path.IsPathRooted(config.TemplatePath) ? config.TemplatePath : Path.Combine(Root, config.TemplatePath);
        template = Path.GetFullPath(template);
        if (!File.Exists(template) || !template.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new FileNotFoundException("Không tìm thấy template XLSX: " + template);
        if (config.PrintMode.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            ExportPdf(config, job, template);
            return;
        }
        if (!config.PrintMode.Equals("printer", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("PrintMode chỉ hỗ trợ printer hoặc pdf.");
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Không tìm thấy Microsoft Excel (Office 2013) trên máy.");
        object? excelObject = null;
        try
        {
            excelObject = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Không thể khởi chạy Excel.");
            dynamic excel = excelObject;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.AutomationSecurity = 3;

            string[] names = { "B8", "H8", "N8", "T8", "B32", "H32", "N32", "T32" };
            string[] colors = { "B12", "H12", "N12", "T12", "B36", "H36", "N36", "T36" };
            string[] quantities = { "B16", "H16", "N16", "T16", "B40", "H40", "N40", "T40" };
            string[] numbers = { "E20", "K20", "Q20", "W20", "E44", "K44", "Q44", "W44" };

            // Each batch uses a fresh read-only template; never changes the source XLSX.
            int total = job.Cards?.Count ?? 0;
            int pageCount = total == 0 ? 1 : (total + 7) / 8;
            for (int page = 0; page < pageCount; page++)
            {
                object? workbookObject = null;
                try
                {
                    workbookObject = excel.Workbooks.Open(template, ReadOnly: true, UpdateLinks: 0);
                    dynamic workbook = workbookObject;
                    var pageCells = new Dictionary<string, string?>(job.Cells, StringComparer.OrdinalIgnoreCase);
                    if (job.Cards is { Count: > 0 })
                    {
                        for (int i = 0; i < 8; i++)
                        {
                            int cardIndex = page * 8 + i;
                            var card = cardIndex < total ? job.Cards[cardIndex] : null;
                            pageCells["原紙!" + names[i]] = card?.Name ?? "";
                            pageCells["原紙!" + colors[i]] = card?.Color ?? "";
                            pageCells["原紙!" + quantities[i]] = card?.Quantity ?? "";
                            pageCells["原紙!" + numbers[i]] =
                                card is null ? "" : (checked(job.StartNumber + cardIndex).ToString(System.Globalization.CultureInfo.InvariantCulture) + (card?.IsFinal == true ? "(完)" : ""));
                        }
                    }
                    foreach (var (location, value) in pageCells)
                    {
                        var separator = location.LastIndexOf('!');
                        var sheetName = separator < 0 ? "" : location[..separator].Trim(' ', '\'');
                        var address = separator < 0 ? location : location[(separator + 1)..];
                        if (!AddressPattern.IsMatch(address) || sheetName.Length > 31)
                            throw new InvalidDataException("Vị trí ô không hợp lệ: " + location);
                        object? sheetObject = null, rangeObject = null;
                        try
                        {
                            sheetObject = string.IsNullOrEmpty(sheetName) ? workbook.Worksheets[1] : workbook.Worksheets[sheetName];
                            dynamic sheet = sheetObject!;
                            rangeObject = sheet.Range[address];
                            dynamic cell = rangeObject!;
                            if ((value?.Length ?? 0) > 1000) throw new InvalidDataException("Nội dung ô quá dài: " + location);
                            cell.Value2 = value ?? "";
                        }
                        finally
                        {
                            Release(rangeObject);
                            Release(sheetObject);
                        }
                    }
                    if (string.IsNullOrWhiteSpace(config.PrinterName))
                        workbook.PrintOut(Copies: job.Copies);
                    else
                        workbook.PrintOut(Copies: job.Copies, ActivePrinter: config.PrinterName);
                    Console.WriteLine($"Đã gửi trang {page + 1}/{pageCount} tới Windows Print Spooler.");
                }
                finally
                {
                    if (workbookObject != null)
                    {
                        try { ((dynamic)workbookObject).Close(SaveChanges: false); }
                        catch (Exception ex) { Console.Error.WriteLine("Không đóng được workbook: " + ex.Message); }
                        Release(workbookObject);
                    }
                }
            }
        }
        finally
        {
            if (excelObject != null)
            {
                try { ((dynamic)excelObject).Quit(); }
                catch (Exception ex) { Console.Error.WriteLine("Không thoát được Excel: " + ex.Message); }
                Release(excelObject);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static string ExportPdf(AgentConfig config, PrintJob job, string template)
    {
        var folder = config.PdfOutputPath;
        if (string.IsNullOrWhiteSpace(folder))
            folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Machine-Control-PDF");
        if (!Path.IsPathFullyQualified(folder))
            throw new InvalidDataException("PdfOutputPath phải là đường dẫn tuyệt đối.");
        folder = Path.GetFullPath(folder);
        Directory.CreateDirectory(folder);
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Không tìm thấy Microsoft Excel (Office 2013).");
        object? excelObject = null, workbookObject = null;
        try
        {
            excelObject = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Không thể mở Excel để xuất PDF.");
            dynamic excel = excelObject;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.AutomationSecurity = 3;
            workbookObject = excel.Workbooks.Open(template, ReadOnly: true, UpdateLinks: 0);
            dynamic workbook = workbookObject;
            string[] names = { "B8", "H8", "N8", "T8", "B32", "H32", "N32", "T32" };
            string[] colors = { "B12", "H12", "N12", "T12", "B36", "H36", "N36", "T36" };
            string[] quantities = { "B16", "H16", "N16", "T16", "B40", "H40", "N40", "T40" };
            string[] numbers = { "E20", "K20", "Q20", "W20", "E44", "K44", "Q44", "W44" };
            int total = job.Cards?.Count ?? 0;
            int pages = total == 0 ? 1 : (total + 7) / 8;
            // Create new sheets in the same workbook to export all cards in ONE PDF.
            // Multiple copies duplicate the pages, preserving the printed serial numbers.
            object? baseSheetObject = null;
            try
            {
                baseSheetObject = workbook.Worksheets["原紙"];
                dynamic baseSheet = baseSheetObject;
                for (int copy = 0; copy < job.Copies; copy++)
                {
                    for (int page = 0; page < pages; page++)
                    {
                        object? sheetObject = null;
                        bool releaseSheet = false;
                        try
                        {
                            // Always duplicate the original template BEFORE filling it.
                            // This prevents page 2 from inheriting data from page 1.
                            baseSheet.Copy(After: workbook.Worksheets[workbook.Worksheets.Count]);
                            sheetObject = workbook.Worksheets[workbook.Worksheets.Count];
                            releaseSheet = true;
                            dynamic sheet = sheetObject!;
                            if (job.Cards is { Count: > 0 })
                            {
                                for (int i = 0; i < 8; i++)
                                {
                                    int cardIndex = page * 8 + i;
                                    var card = cardIndex < total ? job.Cards[cardIndex] : null;
                                    SetPdfCell(sheet, names[i], card?.Name ?? "");
                                    SetPdfCell(sheet, colors[i], card?.Color ?? "");
                                    SetPdfCell(sheet, quantities[i], card?.Quantity ?? "");
                                    SetPdfCell(sheet, numbers[i],
                                        card is null ? "" : (checked(job.StartNumber + cardIndex).ToString(System.Globalization.CultureInfo.InvariantCulture) + (card?.IsFinal == true ? "(完)" : "")));
                                }
                            }
                            foreach (var (location, value) in job.Cells)
                            {
                                var separator = location.LastIndexOf('!');
                                var sheetName = separator < 0 ? "" : location[..separator].Trim(' ', '\'');
                                var address = separator < 0 ? location : location[(separator + 1)..];
                                if (!AddressPattern.IsMatch(address) || sheetName.Length > 31 ||
                                    (!string.IsNullOrEmpty(sheetName) && sheetName != "原紙"))
                                    throw new InvalidDataException("Địa chỉ ô không hợp lệ trên sheet 原紙: " + location);
                                SetPdfCell(sheet, address, value ?? "");
                            }
                        }
                        finally
                        {
                            if (releaseSheet) Release(sheetObject);
                        }
                    }
                }
                // Exclude the unfilled template sheet from the final PDF.
                baseSheet.Delete();
            }
            finally { Release(baseSheetObject); }
            var safeId = Regex.Replace(job.JobId, @"[^a-zA-Z0-9_-]", "_");
            var path = Path.Combine(folder, $"MaterialCards_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{safeId}_{Guid.NewGuid():N}.pdf");
            // xlTypePDF=0, xlQualityStandard=0; honor the template's print areas.
            workbook.ExportAsFixedFormat(0, path, 0, true, false);
            if (!File.Exists(path) || new FileInfo(path).Length == 0)
                throw new IOException("Excel không tạo được PDF hợp lệ.");
            Console.WriteLine("Đã xuất PDF: " + path);
            return path;
        }
        finally
        {
            if (workbookObject != null)
            {
                try { ((dynamic)workbookObject).Close(SaveChanges: false); }
                catch (Exception ex) { Console.Error.WriteLine("Không đóng được workbook PDF: " + ex.Message); }
                Release(workbookObject);
            }
            if (excelObject != null)
            {
                try { ((dynamic)excelObject).Quit(); }
                catch (Exception ex) { Console.Error.WriteLine("Không đóng được Excel PDF: " + ex.Message); }
                Release(excelObject);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void SetPdfCell(dynamic sheet, string address, string value)
    {
        if (!AddressPattern.IsMatch(address) || value.Length > 1000)
            throw new InvalidDataException("Địa chỉ hoặc nội dung PDF không hợp lệ: " + address);
        object? rangeObject = null;
        try
        {
            rangeObject = sheet.Range[address];
            ((dynamic)rangeObject).Value2 = value;
        }
        finally { Release(rangeObject); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }
}
