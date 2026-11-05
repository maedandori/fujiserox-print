using System;
using System.Drawing.Printing;
using System.Net.Http;
using System.Net.Http.Json;
using System.Runtime.InteropServices;

internal static class PrinterHeartbeat
{
    // Win32 spooler reports status flags, NOT guaranteed tray-level paper inventory.
    private const int PaperOut = 0x10;
    private const int PaperProblem = 0x40;
    private const int PaperJam = 0x8;

    [DllImport("winspool.drv", EntryPoint="OpenPrinterW", CharSet=CharSet.Unicode, SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool OpenPrinter(string name, out IntPtr handle, IntPtr defaults);

    [DllImport("winspool.drv", SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetPrinter(IntPtr handle, uint level, IntPtr buffer, uint bufferSize, out uint needed);

    [DllImport("winspool.drv", SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ClosePrinter(IntPtr handle);

    public static void Send(HttpClient client, AgentConfig config, string agentId)
    {
        var name=config.PrinterName;
        if (string.IsNullOrWhiteSpace(name))
            name=new PrinterSettings().PrinterName;
        if (string.IsNullOrWhiteSpace(name)) name="(no Windows printer)";
        bool valid=false;
        try { valid=new PrinterSettings { PrinterName=name }.IsValid; }
        catch { /* Driver may be unavailable. */ }
        int status=0;
        string paper="unknown";
        if (valid && OpenPrinter(name, out var handle, IntPtr.Zero))
        {
            try
            {
                // PRINTER_INFO_6 is a DWORD status (level 6), no structures or spooler jobs required.
                var buffer=Marshal.AllocHGlobal(sizeof(int));
                try
                {
                    if (GetPrinter(handle, 6, buffer, sizeof(int), out _))
                    {
                        status=Marshal.ReadInt32(buffer);
                        if ((status & PaperOut)!=0) paper="paper_out";
                        else if ((status & (PaperProblem|PaperJam))!=0) paper="paper_problem";
                    }
                }
                finally { Marshal.FreeHGlobal(buffer); }
            }
            finally { ClosePrinter(handle); }
        }
        // Driver reporting no error is NOT proof that paper is present.
        using var response=client.PostAsJsonAsync("api/print-agent/status/heartbeat",
            new {agentId,machineName=Environment.MachineName,printerName=name,printerValid=valid,paperStatus=paper,windowsStatus=status})
            .GetAwaiter().GetResult();
        response.EnsureSuccessStatusCode();
    }
}
