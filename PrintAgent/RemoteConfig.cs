using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;

internal static class RemoteConfig
{
    private static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true };
    private static readonly string[] Fields = {"PrinterName","PrinterDisplayName","TemplatePath","PrintMode","PdfOutputPath","PdfPrinterExe","PollSeconds","MaxPdfPages","AutoUpdateEnabled"};
    private sealed class Command { public bool Update {get;set;} public string Revision {get;set;} = ""; public Dictionary<string,JsonElement>? Config {get;set;} }
    private static Dictionary<string,object> Snapshot(AgentConfig cfg) => new()
    {
        ["PrinterName"]=cfg.PrinterName, ["PrinterDisplayName"]=cfg.PrinterDisplayName,
        ["TemplatePath"]=cfg.TemplatePath, ["PrintMode"]=cfg.PrintMode=="printer"?"pdfprint":cfg.PrintMode,
        ["PdfOutputPath"]=cfg.PdfOutputPath, ["PdfPrinterExe"]=cfg.PdfPrinterExe,
        ["PollSeconds"]=cfg.PollSeconds, ["MaxPdfPages"]=cfg.MaxPdfPages,
        ["AutoUpdateEnabled"]=cfg.AutoUpdateEnabled
    };
    public static void SendSnapshot(HttpClient client,AgentConfig config)
    {
        try {
            using var response=client.PostAsJsonAsync("api/print-agent/config/snapshot",new {config=Snapshot(config)}).GetAwaiter().GetResult();
            if(!response.IsSuccessStatusCode)Console.Error.WriteLine("Không đồng bộ cấu hình: HTTP "+(int)response.StatusCode);
        } catch(Exception e) {Console.Error.WriteLine("Không đồng bộ cấu hình: "+e.Message);}
    }
    private static string ReadString(Dictionary<string,JsonElement> cfg,string key,int max=512)
    {
        if(!cfg.TryGetValue(key,out var value)||value.ValueKind!=JsonValueKind.String)throw new InvalidDataException(key+" phải là chuỗi.");
        var v=value.GetString()??"";
        if(v.Length>max||v.Any(c=>char.IsControl(c)))throw new InvalidDataException(key+" có ký tự không hợp lệ.");
        return v.Trim();
    }
    private static int ReadInt(Dictionary<string,JsonElement> cfg,string key,int min,int max)
    {
        if(!cfg.TryGetValue(key,out var value)||!value.TryGetInt32(out var n)||n<min||n>max)throw new InvalidDataException(key+" không hợp lệ.");
        return n;
    }
    private static Dictionary<string,string> Validate(Dictionary<string,JsonElement>? cfg)
    {
        if(cfg==null||cfg.Count!=Fields.Length||cfg.Keys.Any(k=>!Fields.Contains(k,StringComparer.Ordinal)))throw new InvalidDataException("Danh sách thông số không hợp lệ.");
        var result=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        foreach(var key in Fields.Where(k=>k!="PollSeconds"&&k!="MaxPdfPages"&&k!="AutoUpdateEnabled"))
            result[key]=ReadString(cfg,key);
        result["PollSeconds"]=ReadInt(cfg,"PollSeconds",2,300).ToString();
        result["MaxPdfPages"]=ReadInt(cfg,"MaxPdfPages",1,100).ToString();
        if(!cfg.TryGetValue("AutoUpdateEnabled",out var update)||update.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new InvalidDataException("AutoUpdateEnabled không hợp lệ.");
        result["AutoUpdateEnabled"]=update.GetBoolean()?"1":"0";
        if(!new[]{"pdf","pdfprint"}.Contains(result["PrintMode"],StringComparer.Ordinal)||string.IsNullOrEmpty(result["TemplatePath"])||string.IsNullOrEmpty(result["PrinterName"])||!result["PdfPrinterExe"].EndsWith(".exe",StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Chế độ, tên máy in, template hoặc đường dẫn Sumatra không hợp lệ.");
        return result;
    }
    public static AgentConfig CheckAndApply(HttpClient client,AgentConfig active,Action<string> status)
    {
        Command? cmd;
        try {
            using var response=client.GetAsync("api/print-agent/config/command").GetAwaiter().GetResult();
            if(!response.IsSuccessStatusCode)return active;
            cmd=JsonSerializer.Deserialize<Command>(response.Content.ReadAsStringAsync().GetAwaiter().GetResult(),Json);
            if(cmd?.Update!=true)return active;
        } catch(Exception e) {Console.Error.WriteLine("Không đọc được lệnh INI: "+e.Message);return active;}
        var path=Path.Combine(AppContext.BaseDirectory,"config.ini");
        string? temp=null;
        try {
            if(!Guid.TryParse(cmd.Revision,out _))throw new InvalidDataException("Revision không hợp lệ.");
            var values=Validate(cmd.Config);
            var lines=File.ReadAllLines(path,Encoding.UTF8).ToList();
            var pending=new HashSet<string>(values.Keys,StringComparer.OrdinalIgnoreCase);
            for(int i=0;i<lines.Count;i++) {
                var line=lines[i].Trim();
                if(line.StartsWith(";")||line.StartsWith("#")||line.StartsWith("[")||!line.Contains('='))continue;
                var key=line[..line.IndexOf('=')].Trim();
                if(values.TryGetValue(key,out var value)) {lines[i]=key+"="+value;pending.Remove(key);}
            }
            foreach(var key in pending)lines.Add(key+"="+values[key]);
            temp=path+".remote-"+Guid.NewGuid().ToString("N")+".tmp";
            File.WriteAllLines(temp,lines, new UTF8Encoding(false));
            // Validate file with the exact parser used by the running Agent.
            var candidate=IniConfig.Load(temp);
            var backup=path+".bak";
            File.Replace(temp,path,backup);
            temp=null;
            active=candidate;
            status("Đã áp dụng config.ini từ web; bản cũ lưu tại config.ini.bak.");
            SendSnapshot(client,active);
            Report(client,cmd.Revision,"completed",null,Snapshot(active));
        } catch(Exception e) {
            status("Không lưu được config.ini từ web: "+e.Message);
            Report(client,cmd.Revision,"failed",e.Message,null);
        } finally {if(temp!=null)try{File.Delete(temp);}catch{}}
        return active;
    }
    private static void Report(HttpClient client,string revision,string state,string? error,Dictionary<string,object>? config)
    {
        try {
            using var resp=client.PostAsJsonAsync("api/print-agent/config/result",new{revision,status=state,error,config}).GetAwaiter().GetResult();
            if(!resp.IsSuccessStatusCode)Console.Error.WriteLine("Không thể báo trạng thái INI: HTTP "+(int)resp.StatusCode);
        } catch(Exception e) {Console.Error.WriteLine("Không báo được kết quả INI: "+e.Message);}
    }
}
