using System;
using System.Drawing;
using System.IO;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new SettingsForm());
    }
}

internal sealed class SettingsForm : Form
{
    private const string DefaultApiUrl = "https://machine-control.maedandoris.workers.dev";
    private readonly TextBox url = new() { Text = DefaultApiUrl }, token = new(), template = new();
    private readonly ComboBox printer = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown seconds = new();
    private readonly NumericUpDown maxPdfPages = new() { Minimum = 1, Maximum = 100, Value = 96, Width = 110 };
    private readonly ComboBox outputMode = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly TextBox pdfFolder = new();
    private readonly TextBox pdfPrinterExe = new() { Text = "SumatraPDF.exe" };
    private readonly TextBox printerDisplayName = new();
    private readonly CheckBox autoUpdate = new() { Text = "Tự động cập nhật khi Agent không in", AutoSize = true, Checked = true };
    private readonly CheckBox startWithWindows = new() { Text = "Khởi động Print-Agent cùng Windows (khi đăng nhập)", AutoSize = true };
    private readonly Button languageButton = new() { AutoSize = true, Height = 36 };
    private static readonly System.Collections.Generic.Dictionary<string,string> Japanese = new()
    {
        ["Hiện token"]="トークン表示",
        ["Tên máy in (để trống = mặc định)"]="Windowsプリンター",
        ["Làm mới"]="更新",
        ["Tên hiển thị trên web"]="Web表示名",
        ["File mẫu XLSX"]="Excelテンプレート",
        ["Chọn file..."]="ファイル選択...",
        ["Chu kỳ kiểm tra (giây)"]="確認間隔（秒）",
        ["Chế độ xuất"]="出力モード",
        ["Thư mục lưu PDF"]="PDF保存先",
        ["Chọn thư mục..."]="フォルダー選択...",
        ["Tự động cập nhật khi Agent không in"]="待機中に自動更新",
        ["Khởi động Print-Agent cùng Windows (khi đăng nhập)"]="Windowsサインイン時にPrint-Agentを自動起動",
        ["Lưu cấu hình"]="設定を保存",
        ["Đóng"]="閉じる",
        ["Lưu config.ini cùng thư mục với hai ứng dụng. Agent cần chạy khi người dùng Windows đang đăng nhập và Excel 2013 đã được cài."]="設定はEXEと同じフォルダーに保存します。Windowsにサインインし、Excel 2013が必要です。",
        ["Xuất PDF rồi in"]="PDFを作成して印刷",
        ["Bộ in PDF (SumatraPDF)"]="PDF印刷ツール (SumatraPDF)",
        ["Số trang tối đa mỗi PDF"]="PDFごとの最大ページ数",
        ["Chọn EXE..."]="EXEを選択...",
        ["Xuất PDF tự động"]="PDFを自動保存",
        ["(Máy in mặc định Windows)"]="（Windowsの既定プリンター）"
    };
    private string L(string vi,string ja)=>UiLanguage.T(vi,ja);
    private void ApplyLanguage()
    {
        Text = UiLanguage.T("Machine-Control - Cài đặt máy in (v","Machine-Control - 印刷設定 (v")
             + (typeof(Program).Assembly.GetName().Version?.ToString(3)??"1.0.0") + ")";
        languageButton.Text=UiLanguage.Current=="ja"?"🇯🇵 日本語 → 🇻🇳 Tiếng Việt":"🇻🇳 Tiếng Việt → 🇯🇵 日本語";
        void Translate(Control control)
        {
            if(control.Tag is not string && Japanese.ContainsKey(control.Text))control.Tag=control.Text;
            if(control.Tag is string original && Japanese.TryGetValue(original,out var japan))
                control.Text=UiLanguage.Current=="ja"?japan:original;
            foreach(Control child in control.Controls)Translate(child);
        }
        Translate(this);
        var mode=outputMode.SelectedIndex;
        outputMode.Items.Clear();
        outputMode.Items.AddRange(new object[]{UiLanguage.T("Xuất PDF rồi in","PDFを作成して印刷"),UiLanguage.T("Xuất PDF tự động","PDFを自動保存")});
        outputMode.SelectedIndex=mode<0?0:mode;
        var chosen=printer.SelectedIndex>0?printer.Text:"";
        RefreshPrinters(chosen);
        printerDisplayName.PlaceholderText=UiLanguage.T("Ví dụ: Máy in phòng vật tư / 1号機","例: 資材室プリンター / 1号機");
    }
    private readonly string iniPath = Path.Combine(AppContext.BaseDirectory, "config.ini");

    public SettingsForm()
    {
        Text = "Machine-Control - Cài đặt máy in (v"+(typeof(Program).Assembly.GetName().Version?.ToString(3)??"1.0.0")+")";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(690, 730);
        Size = new Size(800, 810);
        Font = new Font("Segoe UI", 10);
        var grid = new TableLayoutPanel {
            Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 3, RowCount = 14,
            AutoSize = false
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        for (int i = 0; i < 13; i++) grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 47));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        url.Dock = token.Dock = printer.Dock = template.Dock = DockStyle.Fill;
        token.UseSystemPasswordChar = true;
        seconds.Minimum = 2; seconds.Maximum = 300; seconds.Value = 3; seconds.Dock = DockStyle.Left; seconds.Width = 110;
        AddLabel(grid, "API Machine-Control", 0); grid.Controls.Add(url, 1, 0); grid.SetColumnSpan(url, 2);
        AddLabel(grid, "Agent Token", 1); grid.Controls.Add(token, 1, 1);
        var showToken = new CheckBox { Text = "Hiện token", AutoSize = true, Dock = DockStyle.Fill };
        showToken.CheckedChanged += (_, _) => token.UseSystemPasswordChar = !showToken.Checked;
        grid.Controls.Add(showToken, 2, 1);
        AddLabel(grid, "Tên máy in (để trống = mặc định)", 2); grid.Controls.Add(printer, 1, 2);
        var refreshPrinters = new Button { Text = "Làm mới", Dock = DockStyle.Fill };
        refreshPrinters.Click += (_, _) => RefreshPrinters(printer.Text);
        grid.Controls.Add(refreshPrinters, 2, 2);
        AddLabel(grid, "Tên hiển thị trên web", 3);
        printerDisplayName.Dock = DockStyle.Fill;
        printerDisplayName.MaxLength = 100;
        printerDisplayName.PlaceholderText = "Ví dụ: Máy in phòng vật tư / 1号機";
        grid.Controls.Add(printerDisplayName, 1, 3);
        grid.SetColumnSpan(printerDisplayName, 2);
        AddLabel(grid, "File mẫu XLSX", 4); grid.Controls.Add(template, 1, 4);
        var browse = new Button { Text = "Chọn file...", Dock = DockStyle.Fill };
        browse.Click += (_, _) => {
            using var dlg = new OpenFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", CheckFileExists = true, Title=UiLanguage.T("Chọn mẫu Excel","Excelテンプレートを選択") };
            if (dlg.ShowDialog(this) == DialogResult.OK) template.Text = dlg.FileName;
        };
        grid.Controls.Add(browse, 2, 4);
        AddLabel(grid, "Chu kỳ kiểm tra (giây)", 5); grid.Controls.Add(seconds, 1, 5);
        AddLabel(grid, "Chế độ xuất", 6);
        outputMode.Items.AddRange(new object[] { "Xuất PDF rồi in", "Xuất PDF tự động" });
        outputMode.SelectedIndex = 0;
        outputMode.Dock = DockStyle.Fill;
        outputMode.SelectedIndexChanged += (_, _) => UpdatePdfInputs();
        grid.Controls.Add(outputMode, 1, 6);
        grid.SetColumnSpan(outputMode, 2);
        AddLabel(grid, "Thư mục lưu PDF", 7);
        pdfFolder.Dock = DockStyle.Fill;
        pdfFolder.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Machine-Control-PDF");
        grid.Controls.Add(pdfFolder, 1, 7);
        var browsePdf = new Button { Text = "Chọn thư mục...", Dock = DockStyle.Fill };
        browsePdf.Click += (_, _) => {
            using var folder = new FolderBrowserDialog { Description = UiLanguage.T("Chọn nơi lưu PDF","PDF保存先を選択"), ShowNewFolderButton = true };
            if (Directory.Exists(pdfFolder.Text)) folder.SelectedPath = pdfFolder.Text;
            if (folder.ShowDialog(this) == DialogResult.OK) pdfFolder.Text = folder.SelectedPath;
        };
        grid.Controls.Add(browsePdf, 2, 7);
        AddLabel(grid, "Bộ in PDF (SumatraPDF)", 8);
        pdfPrinterExe.Dock = DockStyle.Fill;
        grid.Controls.Add(pdfPrinterExe, 1, 8);
        var choosePdfPrinter = new Button { Text = "Chọn EXE...", Dock = DockStyle.Fill };
        choosePdfPrinter.Click += (_, _) => {
            using var dialog = new OpenFileDialog { Filter = "Executable (*.exe)|*.exe", CheckFileExists = true,
                Title = UiLanguage.T("Chọn SumatraPDF.exe","SumatraPDF.exeを選択") };
            if (dialog.ShowDialog(this) == DialogResult.OK) pdfPrinterExe.Text = dialog.FileName;
        };
        grid.Controls.Add(choosePdfPrinter, 2, 8);
        AddLabel(grid, "Số trang tối đa mỗi PDF", 9);
        grid.Controls.Add(maxPdfPages, 1, 9);
        var note = new Label { Text = "Lưu config.ini cùng thư mục với hai ứng dụng. Agent cần chạy khi người dùng Windows đang đăng nhập và Excel 2013 đã được cài.", AutoSize = true, Dock = DockStyle.Fill };
        grid.Controls.Add(autoUpdate, 0, 10); grid.SetColumnSpan(autoUpdate, 3);
        grid.Controls.Add(startWithWindows, 0, 11); grid.SetColumnSpan(startWithWindows, 3);
        grid.Controls.Add(note, 0, 12); grid.SetColumnSpan(note, 3);
        var controls = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
        var save = new Button { Text = "Lưu cấu hình", AutoSize = true, Height = 36 };
        var close = new Button { Text = "Đóng", AutoSize = true, Height = 36 };
        save.Click += (_, _) => SaveConfig();
        close.Click += (_, _) => Close();
        languageButton.Click += (_,_) => { if (UiLanguage.Set(UiLanguage.Current=="ja"?"vi":"ja")) ApplyLanguage(); else MessageBox.Show(L("Không lưu được ngôn ngữ.","言語設定を保存できません。")); };
        controls.Controls.Add(save); controls.Controls.Add(close); controls.Controls.Add(languageButton);
        grid.Controls.Add(controls, 0, 13); grid.SetColumnSpan(controls, 3);
        Controls.Add(grid);
        AcceptButton = save;
        RefreshPrinters("");
        LoadConfig();
        startWithWindows.Checked = IsStartupEnabled();
        UpdatePdfInputs();
        ApplyLanguage();
    }
    private static void AddLabel(TableLayoutPanel grid, string name, int row) =>
        grid.Controls.Add(new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);

    private void RefreshPrinters(string preferred)
    {
        try
        {
            printer.Items.Clear();
            printer.Items.Add(L("(Máy in mặc định Windows)","（Windowsの既定プリンター）"));
            foreach (string installed in PrinterSettings.InstalledPrinters)
                printer.Items.Add(installed);
            if (!string.IsNullOrWhiteSpace(preferred) && !printer.Items.Contains(preferred))
                printer.Items.Add(preferred + " (không tìm thấy)");
            if (!string.IsNullOrWhiteSpace(preferred))
                printer.SelectedItem = printer.Items.Contains(preferred) ? preferred : preferred + " (không tìm thấy)";
            else printer.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, L("Không đọc được danh sách máy in: ","プリンター一覧を読み取れません: ") + ex.Message, L("Lỗi máy in","プリンターエラー"));
        }
    }

    private void UpdatePdfInputs()
    {
        bool isPdf = true;
        pdfFolder.Enabled = isPdf;
        pdfPrinterExe.Enabled = outputMode.SelectedIndex == 0;
        foreach (Control control in Controls[0].Controls)
            if (control is Button button && (button.Tag as string ?? button.Text) == "Chọn thư mục...") button.Enabled = isPdf;
        foreach (Control control in Controls[0].Controls)
            if (control is Button button && (button.Tag as string ?? button.Text) == "Chọn EXE...") button.Enabled = outputMode.SelectedIndex == 0;
    }

    private void LoadConfig()
    {
        if (!File.Exists(iniPath)) { template.Text = "template.xlsx"; return; }
        foreach (var raw in File.ReadAllLines(iniPath, Encoding.UTF8))
        {
            var line = raw.Trim(); if (line.Length == 0 || line[0] == '[' || line[0] == ';' || line[0] == '#') continue;
            var idx = line.IndexOf('='); if (idx < 1) continue;
            var key = line[..idx].Trim(); var value = line[(idx + 1)..].Trim();
            switch (key.ToLowerInvariant())
            {
                case "apibaseurl": if (!string.IsNullOrWhiteSpace(value)) url.Text = value; break;
                case "agenttoken": token.Text = value; break;
                case "printername": RefreshPrinters(value); break;
                case "printerdisplayname": printerDisplayName.Text = value; break;
                case "autoupdateenabled": autoUpdate.Checked = value != "0"; break;
                case "templatepath": template.Text = value; break;
                case "printmode": outputMode.SelectedIndex = value.Equals("pdf", StringComparison.OrdinalIgnoreCase) ? 1 : 0; break;
                case "pdfprinterexe": if (!string.IsNullOrWhiteSpace(value)) pdfPrinterExe.Text = value; break;
                case "pdfoutputpath": if (!string.IsNullOrWhiteSpace(value)) pdfFolder.Text = value; break;
                case "pollseconds": if (int.TryParse(value, out int n)) seconds.Value = Math.Clamp(n, 2, 300); break;
                case "maxpdfpages": if (int.TryParse(value, out int pages)) maxPdfPages.Value = Math.Clamp(pages, 1, 100); break;
            }
        }
    }

    private const string StartupEntry = "MachineControlPrintAgent";
    private static string AgentExePath => Path.Combine(AppContext.BaseDirectory, "Print-Agent.exe");
    private static bool IsStartupEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
        var current = key?.GetValue(StartupEntry) as string;
        return string.Equals(current, "\"" + AgentExePath + "\"", StringComparison.OrdinalIgnoreCase);
    }
    private static void ConfigureStartup(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true)
            ?? throw new InvalidOperationException("Không mở được cấu hình Windows Startup.");
        if (enabled)
        {
            if (!File.Exists(AgentExePath))
                throw new FileNotFoundException("Không tìm thấy Print-Agent.exe cạnh Print-Settings.exe.", AgentExePath);
            key.SetValue(StartupEntry, "\"" + AgentExePath + "\"", RegistryValueKind.String);
        }
        else key.DeleteValue(StartupEntry, false);
    }

    private void SaveConfig()
    {
        try
        {
            var api = url.Text.Trim();
            if (api.Length > 0 && (!Uri.TryCreate(api, UriKind.Absolute, out var u) ||
                (u.Scheme != Uri.UriSchemeHttps && !(u.IsLoopback && u.Scheme == Uri.UriSchemeHttp))))
                throw new InvalidOperationException("Địa chỉ API phải là HTTPS (HTTP chỉ cho localhost).");
            var path = template.Text.Trim();
            if (path.Length == 0 || !path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Hãy chọn file mẫu .xlsx.");
            if (string.IsNullOrWhiteSpace(pdfFolder.Text) || !Path.IsPathFullyQualified(pdfFolder.Text.Trim()))
                throw new InvalidOperationException("Thư mục lưu PDF phải là đường dẫn tuyệt đối.");
            if (outputMode.SelectedIndex == 0 && !pdfPrinterExe.Text.Trim().EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Hãy chọn SumatraPDF.exe để in PDF.");
            if (new[] { api, token.Text, printer.Text, printerDisplayName.Text, path, pdfFolder.Text, pdfPrinterExe.Text }.Any(x => x.Contains('\r') || x.Contains('\n')))
                throw new InvalidOperationException("Giá trị cấu hình không được chứa ký tự xuống dòng.");
            var selectedPrinter = printer.SelectedIndex <= 0 ? "" : printer.Text;
            if (selectedPrinter.EndsWith(" (không tìm thấy)", StringComparison.Ordinal))
                throw new InvalidOperationException("Máy in này không còn được cài đặt. Hãy chọn máy in khác.");
            var text = "[PrintAgent]\r\n"
                + "ApiBaseUrl=" + api + "\r\n"
                + "AgentToken=" + token.Text.Trim() + "\r\n"
                + "PrinterName=" + selectedPrinter + "\r\n"
                + "PrinterDisplayName=" + printerDisplayName.Text.Trim() + "\r\n"
                + "TemplatePath=" + path + "\r\n"
                + "PollSeconds=" + seconds.Value + "\r\n"
                + "MaxPdfPages=" + maxPdfPages.Value + "\r\n"
                + "PrintMode=" + (outputMode.SelectedIndex == 1 ? "pdf" : "pdfprint") + "\r\n"
                + "PdfOutputPath=" + pdfFolder.Text.Trim() + "\r\n"
                + "PdfPrinterExe=" + pdfPrinterExe.Text.Trim() + "\r\n"
                + "AutoUpdateEnabled=" + (autoUpdate.Checked ? "1" : "0") + "\r\n";
            var temp = iniPath + ".tmp";
            File.WriteAllText(temp, text, new UTF8Encoding(false));
            File.Move(temp, iniPath, true);
            ConfigureStartup(startWithWindows.Checked);
            MessageBox.Show(this, L("Đã lưu: ","保存しました: ") + iniPath + L("\nHãy khởi động lại Print-Agent.exe để áp dụng thay đổi.","\n変更を適用するには Print-Agent.exe を再起動してください。"), L("Thành công","保存完了"));
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, L("Không lưu được","保存エラー"), MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }
}
