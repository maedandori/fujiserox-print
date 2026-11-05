using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

internal static class PrinterTrayIcon
{
    [DllImport("user32.dll", SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr handle);

    public static Icon Create()
    {
        using var image = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(image))
        using (var white = new SolidBrush(Color.White))
        using (var dark = new SolidBrush(Color.FromArgb(35, 65, 100)))
        using (var accent = new SolidBrush(Color.FromArgb(24, 140, 200)))
        using (var outline = new Pen(Color.FromArgb(19, 44, 78), 1.5f))
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.Clear(Color.Transparent);
            graphics.FillRectangle(white, 8, 2, 16, 9);
            graphics.DrawRectangle(outline, 8, 2, 16, 9);
            graphics.FillRectangle(dark, 3, 11, 26, 14);
            graphics.DrawRectangle(outline, 3, 11, 26, 14);
            graphics.FillEllipse(accent, 24, 14, 3, 3);
            graphics.FillRectangle(white, 8, 19, 16, 11);
            graphics.DrawRectangle(outline, 8, 19, 16, 11);
            graphics.DrawLine(outline, 11, 24, 21, 24);
            graphics.DrawLine(outline, 11, 27, 20, 27);
        }
        var handle = image.GetHicon();
        try
        {
            using var temporary = Icon.FromHandle(handle);
            return (Icon)temporary.Clone();
        }
        finally { DestroyIcon(handle); }
    }
}
