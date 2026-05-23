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
        Write("ERROR", $"{context} | {FormatException(exception)}");
    }

    private static string FormatException(Exception exception)
    {
        var builder = new StringBuilder();
        var current = exception;
        var depth = 0;

        while (current is not null)
        {
            if (depth > 0)
            {
                builder.Append(" | Inner: ");
            }

            builder.Append(current.GetType().Name);

            if (!string.IsNullOrWhiteSpace(current.Message))
            {
                builder.Append(": ");
                builder.Append(current.Message);
            }

            if (current is FileNotFoundException fileNotFound &&
                !string.IsNullOrWhiteSpace(fileNotFound.FileName))
            {
                builder.Append(" | FileName: ");
                builder.Append(fileNotFound.FileName);
            }

            if (current is DirectoryNotFoundException)
            {
                builder.Append(" | Verifique se a pasta existe e esta acessivel pelo usuario que executa o app.");
            }

            if (!string.IsNullOrWhiteSpace(current.StackTrace))
            {
                var firstStackLine = current.StackTrace
                    .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault();

                if (!string.IsNullOrWhiteSpace(firstStackLine))
                {
                    builder.Append(" | Stack: ");
                    builder.Append(firstStackLine.Trim());
                }
            }

            current = current.InnerException;
            depth++;
        }

        return builder.ToString();
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
