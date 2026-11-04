using System.Net;
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
        return new AgentConfig { ApiBaseUrl = Get("ApiBaseUrl"), AgentToken = Get("AgentToken"),
            PrinterName = Get("PrinterName"), TemplatePath = Get("TemplatePath", "template.xlsx"), PollSeconds = seconds };
    }
}
internal sealed class AgentConfig
{
    public string ApiBaseUrl { get; set; } = "";
    public string AgentToken { get; set; } = "";
    public string PrinterName { get; set; } = "";
    public string TemplatePath { get; set; } = "template.xlsx";
    public int PollSeconds { get; set; } = 3;
}
internal sealed class PrintCard
{
    public string Name { get; set; } = "";
    public string Color { get; set; } = "";
    public string Quantity { get; set; } = "";
}
internal sealed class PrintJob
{
    public string JobId { get; set; } = "";
    public int Copies { get; set; } = 1;
    public List<PrintCard>? Cards { get; set; }
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
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("Machine-Control Print Agent 1.0 - Windows 10 / Excel 2013");
        Console.WriteLine("Chạy trong phiên Windows có người dùng đăng nhập, không chạy dưới SYSTEM.");
        try
        {
            var config = IniConfig.Load(Path.Combine(Root, "config.ini"));
            if (args.Contains("--test", StringComparer.OrdinalIgnoreCase))
            {
                var testPath = Path.Combine(Root, "sample-job.json");
                var job = ReadJob(File.ReadAllText(testPath));
                PrintExcel(config, job);
                Console.WriteLine("Đã gửi bản in thử đến Windows Print Spooler (chưa xác nhận giấy đã in xong).");
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
            Console.WriteLine("Đã bật chế độ nhận lệnh. Bấm Ctrl+C để thoát.");
            while (true)
            {
                try
                {
                    using var response = client.GetAsync("api/print-agent/jobs/next").GetAwaiter().GetResult();
                    if (response.StatusCode == HttpStatusCode.NoContent)
                    {
                        Thread.Sleep(TimeSpan.FromSeconds(config.PollSeconds));
                        continue;
                    }
                    response.EnsureSuccessStatusCode();
                    var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    var job = ReadJob(json);
                    Console.WriteLine($"Nhận yêu cầu in: {job.JobId}");
                    try
                    {
                        PrintExcel(config, job);
                        Report(client, job.JobId, "submitted", null);
                        Console.WriteLine($"Đã gửi đến hàng đợi in Windows: {job.JobId}");
                    }
                    catch (Exception error)
                    {
                        Console.Error.WriteLine("Lỗi in: " + error.Message);
                        Report(client, job.JobId, "failed", error.Message);
                    }
                }
                catch (Exception error)
                {
                    Console.Error.WriteLine("Không lấy được lệnh in: " + error.Message);
                }
                Thread.Sleep(TimeSpan.FromSeconds(config.PollSeconds));
            }
        }
        catch (Exception error)
        {
            Console.Error.WriteLine("Lỗi cấu hình: " + error.Message);
            return 1;
        }
    }

    private static PrintJob ReadJob(string json)
    {
        var job = JsonSerializer.Deserialize<PrintJob>(json, JsonOptions)
            ?? throw new InvalidDataException("Dữ liệu lệnh in trống.");
        if (string.IsNullOrWhiteSpace(job.JobId) || job.JobId.Length > 128 || job.Copies < 1 || job.Copies > 10 ||
            job.Cells is null || job.Cells.Count > 100 || (job.Cards?.Count ?? 0) > 8)
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

    private static void PrintExcel(AgentConfig config, PrintJob job)
    {
        var template = Path.IsPathRooted(config.TemplatePath) ? config.TemplatePath : Path.Combine(Root, config.TemplatePath);
        template = Path.GetFullPath(template);
        if (!File.Exists(template) || !template.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new FileNotFoundException("Không tìm thấy template XLSX: " + template);
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Không tìm thấy Microsoft Excel (Office 2013) trên máy.");
        object? excelObject = null, workbookObject = null;
        try
        {
            excelObject = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException("Không thể khởi chạy Excel.");
            dynamic excel = excelObject;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.AutomationSecurity = 3; // ForceDisable: không chạy macro khi tự động mở file
            workbookObject = excel.Workbooks.Open(template, ReadOnly: true, UpdateLinks: 0);
            dynamic workbook = workbookObject;
            // Mau 部材カード 原紙: 8 the, thu tu tren-trai sang phai, duoi-trai sang phai.
            // Ghi vao o dau tien cua moi vung merged cell de Excel 2013 nhan du lieu.
            if (job.Cards is { Count: > 0 })
            {
                string[] names = { "B8", "H8", "N8", "T8", "B32", "H32", "N32", "T32" };
                string[] colors = { "B12", "H12", "N12", "T12", "B36", "H36", "N36", "T36" };
                string[] quantities = { "B16", "H16", "N16", "T16", "B40", "H40", "N40", "T40" };
                // Neu job co it hon 8 the, xoa noi dung o the con lai tren ban in tam.
                for (int i = 0; i < 8; i++)
                {
                    var card = i < job.Cards.Count ? job.Cards[i] : new PrintCard();
                    if (card is null) throw new InvalidDataException("Thong tin the khong hop le.");
                    job.Cells["原紙!" + names[i]] = card.Name ?? "";
                    job.Cells["原紙!" + colors[i]] = card.Color ?? "";
                    job.Cells["原紙!" + quantities[i]] = card.Quantity ?? "";
                }
            }
            foreach (var (location, value) in job.Cells)
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
            // Nếu PrinterName để trống: dùng máy in mặc định của Windows.
            // Excel.ActivePrinter có thể yêu cầu tên kèm cổng NeXX:, nên ưu tiên máy in mặc định.
            if (string.IsNullOrWhiteSpace(config.PrinterName))
                workbook.PrintOut(Copies: job.Copies);
            else
                workbook.PrintOut(Copies: job.Copies, ActivePrinter: config.PrinterName);
        }
        finally
        {
            if (workbookObject != null)
            {
                try { ((dynamic)workbookObject).Close(SaveChanges: false); }
                catch (Exception ex) { Console.Error.WriteLine("Không đóng được workbook: " + ex.Message); }
                Release(workbookObject);
            }
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

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }
}
