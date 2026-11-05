using System;
using System.Drawing;
using System.IO;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
    private readonly ComboBox outputMode = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly TextBox pdfFolder = new();
    private readonly string iniPath = Path.Combine(AppContext.BaseDirectory, "config.ini");

    public SettingsForm()
    {
        Text = "Machine-Control - Cài đặt máy in";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(690, 490);
        Size = new Size(780, 550);
        Font = new Font("Segoe UI", 10);
        var grid = new TableLayoutPanel {
            Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 3, RowCount = 9,
            AutoSize = false
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        for (int i = 0; i < 8; i++) grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 47));
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
        AddLabel(grid, "File mẫu XLSX", 3); grid.Controls.Add(template, 1, 3);
        var browse = new Button { Text = "Chọn file...", Dock = DockStyle.Fill };
        browse.Click += (_, _) => {
            using var dlg = new OpenFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", CheckFileExists = true };
            if (dlg.ShowDialog(this) == DialogResult.OK) template.Text = dlg.FileName;
        };
        grid.Controls.Add(browse, 2, 3);
        AddLabel(grid, "Chu kỳ kiểm tra (giây)", 4); grid.Controls.Add(seconds, 1, 4);
        AddLabel(grid, "Chế độ xuất", 5);
        outputMode.Items.AddRange(new object[] { "In ra máy in", "Xuất PDF tự động" });
        outputMode.SelectedIndex = 0;
        outputMode.Dock = DockStyle.Fill;
        outputMode.SelectedIndexChanged += (_, _) => UpdatePdfInputs();
        grid.Controls.Add(outputMode, 1, 5);
        grid.SetColumnSpan(outputMode, 2);
        AddLabel(grid, "Thư mục lưu PDF", 6);
        pdfFolder.Dock = DockStyle.Fill;
        pdfFolder.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Machine-Control-PDF");
        grid.Controls.Add(pdfFolder, 1, 6);
        var browsePdf = new Button { Text = "Chọn thư mục...", Dock = DockStyle.Fill };
        browsePdf.Click += (_, _) => {
            using var folder = new FolderBrowserDialog { Description = "Chọn nơi lưu PDF", ShowNewFolderButton = true };
            if (Directory.Exists(pdfFolder.Text)) folder.SelectedPath = pdfFolder.Text;
            if (folder.ShowDialog(this) == DialogResult.OK) pdfFolder.Text = folder.SelectedPath;
        };
        grid.Controls.Add(browsePdf, 2, 6);
        var note = new Label { Text = "Lưu config.ini cùng thư mục với hai ứng dụng. Agent cần chạy khi người dùng Windows đang đăng nhập và Excel 2013 đã được cài.", AutoSize = true, Dock = DockStyle.Fill };
        grid.Controls.Add(note, 0, 7); grid.SetColumnSpan(note, 3);
        var controls = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
        var save = new Button { Text = "Lưu cấu hình", AutoSize = true, Height = 36 };
        var close = new Button { Text = "Đóng", AutoSize = true, Height = 36 };
        save.Click += (_, _) => SaveConfig();
        close.Click += (_, _) => Close();
        controls.Controls.Add(save); controls.Controls.Add(close);
        grid.Controls.Add(controls, 0, 8); grid.SetColumnSpan(controls, 3);
        Controls.Add(grid);
        AcceptButton = save;
        RefreshPrinters("");
        LoadConfig();
        UpdatePdfInputs();
    }
    private static void AddLabel(TableLayoutPanel grid, string name, int row) =>
        grid.Controls.Add(new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);

    private void RefreshPrinters(string preferred)
    {
        try
        {
            printer.Items.Clear();
            printer.Items.Add("(Máy in mặc định Windows)");
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
            MessageBox.Show(this, "Không đọc được danh sách máy in: " + ex.Message, "Lỗi máy in");
        }
    }

    private void UpdatePdfInputs()
    {
        bool isPdf = outputMode.SelectedIndex == 1;
        pdfFolder.Enabled = isPdf;
        foreach (Control control in Controls[0].Controls)
            if (control is Button button && button.Text == "Chọn thư mục...") button.Enabled = isPdf;
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
                case "templatepath": template.Text = value; break;
                case "printmode": outputMode.SelectedIndex = value.Equals("pdf", StringComparison.OrdinalIgnoreCase) ? 1 : 0; break;
                case "pdfoutputpath": if (!string.IsNullOrWhiteSpace(value)) pdfFolder.Text = value; break;
                case "pollseconds": if (int.TryParse(value, out int n)) seconds.Value = Math.Clamp(n, 2, 300); break;
            }
        }
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
            if (outputMode.SelectedIndex == 1 && (string.IsNullOrWhiteSpace(pdfFolder.Text) || !Path.IsPathFullyQualified(pdfFolder.Text.Trim())))
                throw new InvalidOperationException("Thư mục lưu PDF phải là đường dẫn tuyệt đối.");
            if (new[] { api, token.Text, printer.Text, path, pdfFolder.Text }.Any(x => x.Contains('\r') || x.Contains('\n')))
                throw new InvalidOperationException("Giá trị cấu hình không được chứa ký tự xuống dòng.");
            var selectedPrinter = printer.SelectedIndex <= 0 ? "" : printer.Text;
            if (selectedPrinter.EndsWith(" (không tìm thấy)", StringComparison.Ordinal))
                throw new InvalidOperationException("Máy in này không còn được cài đặt. Hãy chọn máy in khác.");
            var text = "[PrintAgent]\r\n"
                + "ApiBaseUrl=" + api + "\r\n"
                + "AgentToken=" + token.Text.Trim() + "\r\n"
                + "PrinterName=" + selectedPrinter + "\r\n"
                + "TemplatePath=" + path + "\r\n"
                + "PollSeconds=" + seconds.Value + "\r\n"
                + "PrintMode=" + (outputMode.SelectedIndex == 1 ? "pdf" : "printer") + "\r\n"
                + "PdfOutputPath=" + pdfFolder.Text.Trim() + "\r\n";
            var temp = iniPath + ".tmp";
            File.WriteAllText(temp, text, new UTF8Encoding(false));
            File.Move(temp, iniPath, true);
            MessageBox.Show(this, "Đã lưu: " + iniPath + "\nHãy khởi động lại Print-Agent.exe để áp dụng thay đổi.", "Thành công");
        }
        catch (Exception ex) { MessageBox.Show(this, ex.Message, "Không lưu được", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }
}
