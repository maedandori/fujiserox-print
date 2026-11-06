using System;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;

internal static class PortablePdf
{
    private sealed class Manifest
    {
        public string? SumatraSha256 {get;set;}
        public long SumatraSize {get;set;}
    }
    // Gói cập nhật Print-Agent 1.0.67 vẫn chứa 2 EXE. Bản 1.0.68 tự tải
    // SumatraPDF.exe riêng sau khi khởi động, bằng token Worker và SHA-256 Release.
    public static void Ensure(HttpClient client,AgentConfig config,Action<string> status,CancellationToken cancel)
    {
        if(config.PrintMode.Equals("pdf",StringComparison.OrdinalIgnoreCase))return;
        if(!string.Equals(config.PdfPrinterExe,"SumatraPDF.exe",StringComparison.OrdinalIgnoreCase))return;
        var target=Path.Combine(AppContext.BaseDirectory,"SumatraPDF.exe");
        if(File.Exists(target)&&new FileInfo(target).Length>1024)return;
        string? temporary=null;
        try
        {
            using var reply=client.GetAsync("api/print-agent/update/latest",cancel).GetAwaiter().GetResult();
            if(!reply.IsSuccessStatusCode)return;
            var manifest=JsonSerializer.Deserialize<Manifest>(reply.Content.ReadAsStringAsync(cancel).GetAwaiter().GetResult(),new JsonSerializerOptions{PropertyNameCaseInsensitive=true});
            if(manifest==null||manifest.SumatraSize<1024||manifest.SumatraSize>40*1024*1024||
               string.IsNullOrWhiteSpace(manifest.SumatraSha256)||manifest.SumatraSha256.Length!=64||
               !System.Text.RegularExpressions.Regex.IsMatch(manifest.SumatraSha256,"^[0-9a-f]{64}$",System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                return; // Release cũ, không có SumatraPDF.
            status("Đang tải SumatraPDF Portable từ gói phát hành đã xác minh...");
            temporary=target+"."+Guid.NewGuid().ToString("N")+".download";
            using(var response=client.GetAsync("api/print-agent/update/sumatra",HttpCompletionOption.ResponseHeadersRead,cancel).GetAwaiter().GetResult())
            {
                response.EnsureSuccessStatusCode();
                using var input=response.Content.ReadAsStreamAsync(cancel).GetAwaiter().GetResult();
                using var output=new FileStream(temporary,FileMode.CreateNew,FileAccess.Write,FileShare.None);
                var buffer=new byte[81920];int n;long total=0;
                while((n=input.Read(buffer,0,buffer.Length))>0)
                {
                    cancel.ThrowIfCancellationRequested();
                    total+=n;if(total>manifest.SumatraSize)throw new InvalidDataException("SumatraPDF lớn hơn manifest");
                    output.Write(buffer,0,n);
                }
                if(total!=manifest.SumatraSize)throw new InvalidDataException("SumatraPDF chưa tải đủ");
            }
            using(var stream=File.OpenRead(temporary))
            {
                var hash=Convert.ToHexString(SHA256.HashData(stream));
                if(!hash.Equals(manifest.SumatraSha256,StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("SHA-256 SumatraPDF không khớp");
            }
            // Không được ghi đè bộ in Portable đã được quản trị viên cài thủ công.
            if(!File.Exists(target))File.Move(temporary,target);
            status("Đã có SumatraPDF Portable; không cần cài phần mềm.");
        }
        catch(Exception e) {status("Không tự tải được SumatraPDF: "+e.Message);}
        finally {if(temporary!=null)try{File.Delete(temporary);}catch{}}
    }
}
