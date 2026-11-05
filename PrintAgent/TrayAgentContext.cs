using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

internal sealed class TrayAgentContext : ApplicationContext
{
    private readonly NotifyIcon tray;
    private readonly Icon printerIcon;
    private readonly Control dispatcher = new();
    private readonly CancellationTokenSource stopping = new();
    private readonly Thread worker;
    private readonly object logLock = new();
    private string lastStatus = "Đang khởi động";
    private bool exiting;
    private readonly ToolStripMenuItem vietnameseItem;
    private readonly ToolStripMenuItem japaneseItem;
    private readonly ToolStripMenuItem statusItem;
    private readonly ToolStripMenuItem settingsItem;
    private readonly ToolStripMenuItem logItem;
    private readonly ToolStripMenuItem updateItem;
    private readonly ToolStripMenuItem exitItem;
    private readonly ToolStripItem versionItem;

    public TrayAgentContext()
    {
        dispatcher.CreateControl();
        var menu = new ContextMenuStrip();
        versionItem=menu.Items.Add("");
        versionItem.Enabled=false;
        var languageMenu=new ToolStripMenuItem();
        vietnameseItem=new ToolStripMenuItem("🇻🇳 Tiếng Việt");
        japaneseItem=new ToolStripMenuItem("🇯🇵 日本語");
        vietnameseItem.Click+=(_,_)=>ChangeLanguage("vi");
        japaneseItem.Click+=(_,_)=>ChangeLanguage("ja");
        languageMenu.DropDownItems.Add(vietnameseItem);
        languageMenu.DropDownItems.Add(japaneseItem);
        menu.Items.Add(languageMenu);
        updateItem=new ToolStripMenuItem();
        updateItem.Click+=(_,_)=>{
            AutoUpdate.RequestCheck();
            UpdateStatus(UiLanguage.T("Đã yêu cầu kiểm tra phiên bản khi hàng đợi rảnh.","印刷待機時に更新を確認します。"));
        };
        menu.Items.Add(updateItem);
        statusItem=new ToolStripMenuItem();
        statusItem.Click+=(_,_)=>MessageBox.Show(lastStatus, "Machine-Control Print Agent",MessageBoxButtons.OK,MessageBoxIcon.Information);
        menu.Items.Add(statusItem);
        settingsItem=new ToolStripMenuItem();
        settingsItem.Click+=(_,_)=>OpenSettings();
        menu.Items.Add(settingsItem);
        logItem=new ToolStripMenuItem();
        logItem.Click+=(_,_)=>OpenLog();
        menu.Items.Add(logItem);
        menu.Items.Add(new ToolStripSeparator());
        exitItem=new ToolStripMenuItem();
        exitItem.Click+=(_,_)=>StopAgent();
        menu.Items.Add(exitItem);
        ApplyLanguage(languageMenu);
        printerIcon = PrinterTrayIcon.Create();
        tray = new NotifyIcon
        {
            Text = "Machine-Control Print Agent",
            Icon = printerIcon,
            ContextMenuStrip = menu,
            Visible = true
        };
        tray.DoubleClick += (_, _) =>
            MessageBox.Show(lastStatus, "Machine-Control Print Agent");

        worker = new Thread(() =>
        {
            try
            {
                Program.RunAgent(Array.Empty<string>(), stopping.Token, UpdateStatus);
            }
            catch (Exception ex) { UpdateStatus("Lỗi ứng dụng: " + ex.Message); }
        })
        {
            IsBackground = false, // Finish current Excel operation cleanly on exit.
            Name = "Machine-Control Excel Printer"
        };
        worker.SetApartmentState(ApartmentState.STA);
        worker.Start();
    }

    private void ChangeLanguage(string language)
    {
        if (!UiLanguage.Set(language))
        {
            MessageBox.Show(UiLanguage.T("Không lưu được ngôn ngữ.","言語設定を保存できません。"),"Machine-Control");
            return;
        }
        // The other EXE reads the same language file on its next launch.
        var languageMenu=vietnameseItem.OwnerItem as ToolStripMenuItem;
        if(languageMenu!=null)ApplyLanguage(languageMenu);
    }
    private void ApplyLanguage(ToolStripMenuItem languageMenu)
    {
        versionItem.Text=UiLanguage.T("Phiên bản ","バージョン ")+AutoUpdate.Version;
        languageMenu.Text=UiLanguage.T("Ngôn ngữ / 言語","言語 / Ngôn ngữ");
        vietnameseItem.Checked=UiLanguage.Current=="vi";
        japaneseItem.Checked=UiLanguage.Current=="ja";
        updateItem.Text=UiLanguage.T("Kiểm tra cập nhật","更新を確認");
        statusItem.Text=UiLanguage.T("Xem trạng thái","状態を表示");
        settingsItem.Text=UiLanguage.T("Mở cài đặt","設定を開く");
        logItem.Text=UiLanguage.T("Xem nhật ký","ログを表示");
        exitItem.Text=UiLanguage.T("Thoát Print Agent","Print Agent を終了");
    }

    private void UpdateStatus(string message)
    {
        if(message=="AUTOMATIC_UPDATE_RESTART")
        {
            try {dispatcher.BeginInvoke(new Action(StopAgent));}
            catch { /* Process exits without replacing a job currently printing. */ }
            return;
        }
        lastStatus = message;
        try
        {
            lock (logLock)
            {
                var logPath = Path.Combine(AppContext.BaseDirectory, "print-agent.log");
                File.AppendAllText(logPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message + Environment.NewLine);
            }
        }
        catch { /* Lack of log write access must not stop printing. */ }
    }

    private static void OpenSettings()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Print-Settings.exe");
        if (!File.Exists(path))
        {
            MessageBox.Show(UiLanguage.T("Không tìm thấy Print-Settings.exe cùng thư mục Agent.","Agentと同じフォルダーに Print-Settings.exe がありません。"), "Machine-Control");
            return;
        }
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(ex.Message,UiLanguage.T("Lỗi mở cài đặt","設定を開けません")); }
    }

    private void OpenLog()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "print-agent.log");
        try
        {
            if (!File.Exists(path)) File.WriteAllText(path,UiLanguage.T("Chưa có nhật ký.","ログがありません。"));
            Process.Start(new ProcessStartInfo("notepad.exe")
            {
                UseShellExecute = true, Arguments = "\"" + path + "\""
            });
        }
        catch (Exception ex) { MessageBox.Show(ex.Message,UiLanguage.T("Lỗi mở nhật ký","ログを開けません")); }
    }

    private void StopAgent()
    {
        if (exiting) return;
        exiting = true;
        stopping.Cancel();
        tray.Visible = false;
        tray.Dispose();
        printerIcon.Dispose();
        dispatcher.Dispose();
        // The non-background printing worker completes any Excel COM operation
        // before the process ends; do not kill Excel in the middle of a job.
        ExitThread();
    }
}
