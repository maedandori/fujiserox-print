using System;
using System.IO;
using System.Text;

internal static class AgentIdentity
{
    public static string LoadOrCreate()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "agent-id.txt");
        if (File.Exists(path))
        {
            var existing = File.ReadAllText(path, Encoding.UTF8).Trim();
            if (Guid.TryParse(existing, out var id)) return id.ToString("D");
            throw new InvalidDataException("agent-id.txt không hợp lệ; không tự đổi mã máy khi đã đăng ký.");
        }
        var created = Guid.NewGuid().ToString("D");
        // Avoid accidental ID churn; only generate once.
        using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        using (var writer = new StreamWriter(file, new UTF8Encoding(false)))
            writer.WriteLine(created);
        return created;
    }
}
