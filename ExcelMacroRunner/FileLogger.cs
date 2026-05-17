using System.Globalization;
using System.Text;

namespace ExcelMacroRunner;

public sealed class FileLogger
{
    private readonly object _sync = new();
    private string _logFolder;

    public FileLogger(string logFolder)
    {
        _logFolder = NormalizeFolder(logFolder);
        Directory.CreateDirectory(_logFolder);
    }

    public event EventHandler<string>? LogWritten;

    public string CurrentLogFolder
    {
        get
        {
            lock (_sync)
            {
                return _logFolder;
            }
        }
    }

    public void UpdateLogFolder(string logFolder)
    {
        var normalized = NormalizeFolder(logFolder);

        lock (_sync)
        {
            _logFolder = normalized;
            Directory.CreateDirectory(_logFolder);
        }
    }

    public void Info(string message) => Write("INFO", message);

    public void Warning(string message) => Write("WARN", message);

    public void Error(string message) => Write("ERROR", message);

    public void Exception(string context, Exception exception)
    {
        Write("ERROR", $"{context} | {exception.GetType().Name}: {exception.Message}");
    }

    private void Write(string level, string message)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        var line = $"{timestamp} [{level}] {message}";
        var logPath = GetCurrentLogPath();

        lock (_sync)
        {
            File.AppendAllText(logPath, line + Environment.NewLine, new UTF8Encoding(false));
        }

        LogWritten?.Invoke(this, line);
    }

    private string GetCurrentLogPath()
    {
        lock (_sync)
        {
            Directory.CreateDirectory(_logFolder);
            var fileName = $"ExcelMacroRunner_{DateTime.Now:yyyyMMdd}.log";
            return Path.Combine(_logFolder, fileName);
        }
    }

    private static string NormalizeFolder(string logFolder)
    {
        if (string.IsNullOrWhiteSpace(logFolder))
        {
            return Path.Combine(AppContext.BaseDirectory, "Logs", "ExcelMacroRunner");
        }

        return Path.GetFullPath(Environment.ExpandEnvironmentVariables(logFolder.Trim()));
    }
}
