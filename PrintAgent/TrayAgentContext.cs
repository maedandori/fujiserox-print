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
    private readonly CancellationTokenSource stopping = new();
    private readonly Thread worker;
    private readonly object logLock = new();
    private string lastStatus = "Đang khởi động";
    private bool exiting;

    public TrayAgentContext()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("Xem trạng thái", null, (_, _) =>
            MessageBox.Show(lastStatus, "Machine-Control Print Agent",
                MessageBoxButtons.OK, MessageBoxIcon.Information));
        menu.Items.Add("Mở cài đặt", null, (_, _) => OpenSettings());
        menu.Items.Add("Xem nhật ký", null, (_, _) => OpenLog());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Thoát Print Agent", null, (_, _) => StopAgent());

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

    private void UpdateStatus(string message)
    {
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
            MessageBox.Show("Không tìm thấy Print-Settings.exe cùng thư mục Agent.", "Machine-Control");
            return;
        }
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi mở cài đặt"); }
    }

    private void OpenLog()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "print-agent.log");
        try
        {
            if (!File.Exists(path)) File.WriteAllText(path, "Chưa có nhật ký.");
            Process.Start(new ProcessStartInfo("notepad.exe")
            {
                UseShellExecute = true, Arguments = "\"" + path + "\""
            });
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Lỗi mở nhật ký"); }
    }

    private void StopAgent()
    {
        if (exiting) return;
        exiting = true;
        stopping.Cancel();
        tray.Visible = false;
        tray.Dispose();
        printerIcon.Dispose();
        // The non-background printing worker completes any Excel COM operation
        // before the process ends; do not kill Excel in the middle of a job.
        ExitThread();
    }
}
