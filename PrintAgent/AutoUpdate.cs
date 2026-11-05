using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

// Self-update only while idle. All other files (INI, Agent ID, XLSX, logs) stay untouched.
// Private GitHub assets are downloaded exclusively via the authenticated Machine-Control Worker.
internal static class AutoUpdate
{
    private sealed class Manifest
    {
        public string Version { get; set; } = "";
        public string Sha256 { get; set; } = "";
        public long Size { get; set; }
    }
    private static int checkRequested;
    public static string Version => typeof(Program).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";
    public static void RequestCheck() => Interlocked.Exchange(ref checkRequested,1);
    public static bool ConsumeRequest() => Interlocked.Exchange(ref checkRequested,0)==1;

    public static bool TryStage(HttpClient client, Action<string> status, CancellationToken cancel)
    {
        try
        {
            using var request = client.GetAsync("api/print-agent/update/latest",HttpCompletionOption.ResponseHeadersRead,cancel)
                .GetAwaiter().GetResult();
            if (!request.IsSuccessStatusCode)
            {
                status("Cập nhật: máy chủ chưa sẵn sàng (HTTP " + (int)request.StatusCode + ")");
                return false;
            }
            var json=request.Content.ReadAsStringAsync(cancel).GetAwaiter().GetResult();
            var manifest=JsonSerializer.Deserialize<Manifest>(json,new JsonSerializerOptions {PropertyNameCaseInsensitive=true});
            if(manifest==null || !System.Version.TryParse(manifest.Version,out var latest) ||
               !System.Version.TryParse(Version,out var current) ||
               manifest.Size<1024 || manifest.Size>250*1024*1024 ||
               manifest.Sha256.Length!=64 || !System.Text.RegularExpressions.Regex.IsMatch(manifest.Sha256,"^[a-f0-9]{64}$",System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                throw new InvalidDataException("Manifest cập nhật không hợp lệ.");
            if(latest<=current) return false;
            status("Tìm thấy bản cập nhật v"+manifest.Version+" (hiện tại v"+Version+")");
            // Never write into the program directory until the package is fully verified.
            var staging=Path.Combine(Path.GetTempPath(),"MachineControlPrintUpdate-"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(staging);
            var zipPath=Path.Combine(staging,"update.zip");
            try
            {
                using(var download=client.GetAsync("api/print-agent/update/package?version="+Uri.EscapeDataString(manifest.Version),HttpCompletionOption.ResponseHeadersRead,cancel).GetAwaiter().GetResult())
                {
                    download.EnsureSuccessStatusCode();
                    using var content=download.Content.ReadAsStreamAsync(cancel).GetAwaiter().GetResult();
                    using var destination=new FileStream(zipPath,FileMode.CreateNew,FileAccess.Write,FileShare.None);
                    using var deadline=CancellationTokenSource.CreateLinkedTokenSource(cancel);
                    deadline.CancelAfter(TimeSpan.FromMinutes(6));
                    content.CopyToAsync(destination,81920,deadline.Token).GetAwaiter().GetResult();
                }
                var info=new FileInfo(zipPath);
                if(info.Length!=manifest.Size)throw new InvalidDataException("Kích thước ZIP không khớp manifest.");
                using(var stream=File.OpenRead(zipPath))
                {
                    var digest=Convert.ToHexString(SHA256.HashData(stream));
                    if(!digest.Equals(manifest.Sha256,StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("SHA-256 không khớp, đã hủy cập nhật.");
                }
                using(var archive=ZipFile.OpenRead(zipPath))
                {
                    if(archive.Entries.Count!=2)throw new InvalidDataException("Gói cập nhật phải chứa đúng 2 EXE.");
                    foreach(var name in new[] {"Print-Agent.exe","Print-Settings.exe"})
                    {
                        var entry=archive.GetEntry(name) ?? throw new InvalidDataException("Thiếu "+name);
                        if(entry.Length<1024||entry.Length>180*1024*1024)
                            throw new InvalidDataException("Kích thước EXE không hợp lệ.");
                        entry.ExtractToFile(Path.Combine(staging,name));
                    }
                }
                // Make sure program location supports overwriting; if not, abort rather than lose Agent.
                var probe=Path.Combine(AppContext.BaseDirectory,".update-write-test-"+Guid.NewGuid().ToString("N"));
                File.WriteAllText(probe,"test");File.Delete(probe);
                LaunchReplacement(staging,Environment.ProcessId);
                status("AUTOMATIC_UPDATE_RESTART");
                return true;
            }
            catch { try {Directory.Delete(staging,true);}catch{} throw; }
        }
        catch(OperationCanceledException){return false;}
        catch(Exception ex)
        {
            status("Không thể cập nhật: "+ex.Message);
            return false; // Printing must continue if updates are unavailable.
        }
    }

    private static void LaunchReplacement(string staging,int processId)
    {
        static string Quoted(string value)=>"'"+value.Replace("'","''")+"'";
        var dir=AppContext.BaseDirectory;
        var agent=Path.Combine(dir,"Print-Agent.exe");
        var settings=Path.Combine(dir,"Print-Settings.exe");
        var srcAgent=Path.Combine(staging,"Print-Agent.exe");
        var srcSettings=Path.Combine(staging,"Print-Settings.exe");
        var script=String.Join("\n",new[]{
          "$ErrorActionPreference='Stop'",
          "$oldPid="+processId,
          "$agent="+Quoted(agent),
          "$settings="+Quoted(settings),
          "$srcAgent="+Quoted(srcAgent),
          "$srcSettings="+Quoted(srcSettings),
          "$staging="+Quoted(staging),
          "try { Wait-Process -Id $oldPid -Timeout 180 -ErrorAction SilentlyContinue } catch {}",
          "$updated=$false",
          "for($i=0;$i -lt 90;$i++){",
          "  try {",
          "    if(Get-Process -Name 'Print-Settings' -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 2; continue }",
          "    Copy-Item -LiteralPath $srcAgent -Destination $agent -Force -ErrorAction Stop",
          "    Copy-Item -LiteralPath $srcSettings -Destination $settings -Force -ErrorAction Stop",
          "    $updated=$true; break",
          "  } catch { Start-Sleep -Seconds 2 }",
          "}",
          "if($updated){ Start-Process -FilePath $agent -WorkingDirectory "+Quoted(dir)+"; Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }"
        });
        var encoded=Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
        var process=new ProcessStartInfo("powershell.exe")
        {
            UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden,
            Arguments="-NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand "+encoded
        };
        Process.Start(process) ?? throw new InvalidOperationException("Không thể chạy Windows Updater.");
    }
}
