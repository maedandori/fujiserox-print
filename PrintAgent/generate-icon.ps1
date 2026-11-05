# Tạo icon máy in Win32 cho file EXE bằng công cụ có sẵn trên Windows runner.
Add-Type -AssemblyName System.Drawing
$bitmap = [System.Drawing.Bitmap]::new(32,32)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::Transparent)
$dark = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(35,65,100))
$light = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
$accent = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(24,140,200))
$border = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(19,44,78),2)
try {
  $graphics.FillRectangle($light,8,2,16,12)
  $graphics.DrawRectangle($border,8,2,16,12)
  $graphics.FillRectangle($dark,3,12,26,15)
  $graphics.DrawRectangle($border,3,12,26,15)
  $graphics.FillEllipse($accent,24,15,3,3)
  $graphics.FillRectangle($light,8,20,16,10)
  $graphics.DrawRectangle($border,8,20,16,10)
  $graphics.DrawLine($border,11,24,21,24)
  $handle=$bitmap.GetHicon()
  $icon=[System.Drawing.Icon]::FromHandle($handle)
  try {
    $stream=[System.IO.File]::Create((Join-Path $PSScriptRoot "printer.ico"))
    try { $icon.Save($stream) } finally { $stream.Dispose() }
  } finally {
    $icon.Dispose()
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class NativeIcon { [DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr handle); }' -ErrorAction SilentlyContinue
    [NativeIcon]::DestroyIcon($handle) | Out-Null
  }
} finally {
  $border.Dispose();$accent.Dispose();$light.Dispose();$dark.Dispose()
  $graphics.Dispose();$bitmap.Dispose()
}
