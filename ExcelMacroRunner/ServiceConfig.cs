using System.Text.Json;

namespace ExcelMacroRunner;

public sealed class ServiceConfigRoot
{
    public ServiceConfig ServiceConfig { get; set; } = new();
}

public sealed class ServiceConfig
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    public string BatPath { get; set; } = string.Empty;
    public List<ExcelMacroJobConfig> ExcelJobs { get; set; } = new();
    public string HtmlPath { get; set; } = string.Empty;
    public int LoopDelaySeconds { get; set; } = 5;
    public int HtmlMaxAgeMinutes { get; set; } = 10;
    public bool KillExcelOnTimeout { get; set; } = true;
    public bool ShowExcel { get; set; }
    public string LogFolder { get; set; } = Path.Combine(AppContext.BaseDirectory, "Logs", "ExcelMacroRunner");

    public static ServiceConfig Load(string settingsPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(settingsPath);

        if (!File.Exists(settingsPath))
        {
            throw new FileNotFoundException("O arquivo appsettings.json não foi encontrado.", settingsPath);
        }

        using var stream = File.OpenRead(settingsPath);
        var root = JsonSerializer.Deserialize<ServiceConfigRoot>(stream, JsonOptions)
            ?? throw new InvalidOperationException("Não foi possível ler o arquivo appsettings.json.");

        root.ServiceConfig.Normalize();
        root.ServiceConfig.Validate();
        return root.ServiceConfig;
    }

    public static bool TryLoad(string settingsPath, out ServiceConfig? config, out string? errorMessage)
    {
        try
        {
            config = Load(settingsPath);
            errorMessage = null;
            return true;
        }
        catch (Exception ex)
        {
            config = null;
            errorMessage = ex.Message;
            return false;
        }
    }

    public static ServiceConfig CreateFallback(string baseDirectory)
    {
        var safeBaseDirectory = string.IsNullOrWhiteSpace(baseDirectory)
            ? AppContext.BaseDirectory
            : baseDirectory;

        return new ServiceConfig
        {
            HtmlPath = Path.Combine(safeBaseDirectory, "dashboard.html"),
            LogFolder = Path.Combine(safeBaseDirectory, "Logs", "ExcelMacroRunner"),
            ExcelJobs =
            {
                new ExcelMacroJobConfig
                {
                    Name = "Job 1",
                    ExcelFilePath = Path.Combine(safeBaseDirectory, "Atualizador1.xlsm"),
                    MacroName = "Modulo1.AtualizarBase1"
                },
                new ExcelMacroJobConfig
                {
                    Name = "Job 2",
                    ExcelFilePath = Path.Combine(safeBaseDirectory, "Atualizador2.xlsm"),
                    MacroName = "Modulo1.AtualizarBase2"
                },
                new ExcelMacroJobConfig
                {
                    Name = "Job 3",
                    ExcelFilePath = Path.Combine(safeBaseDirectory, "Consolidador.xlsm"),
                    MacroName = "Modulo1.ConsolidarTudo"
                }
            }
        };
    }

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(BatPath))
        {
            throw new InvalidOperationException("ServiceConfig.BatPath não foi configurado.");
        }

        if (string.IsNullOrWhiteSpace(HtmlPath))
        {
            throw new InvalidOperationException("ServiceConfig.HtmlPath não foi configurado.");
        }

        if (string.IsNullOrWhiteSpace(LogFolder))
        {
            throw new InvalidOperationException("ServiceConfig.LogFolder não foi configurado.");
        }

        if (LoopDelaySeconds <= 0)
        {
            throw new InvalidOperationException("ServiceConfig.LoopDelaySeconds deve ser maior que zero.");
        }

        if (HtmlMaxAgeMinutes <= 0)
        {
            throw new InvalidOperationException("ServiceConfig.HtmlMaxAgeMinutes deve ser maior que zero.");
        }

        if (ExcelJobs.Count == 0)
        {
            throw new InvalidOperationException("Configure pelo menos um item em ServiceConfig.ExcelJobs.");
        }

        foreach (var job in ExcelJobs)
        {
            job.Validate();
        }
    }

    private void Normalize()
    {
        BatPath = NormalizePath(BatPath);
        HtmlPath = NormalizePath(HtmlPath);
        LogFolder = NormalizePath(LogFolder);

        foreach (var job in ExcelJobs)
        {
            job.Normalize();
        }
    }

    private static string NormalizePath(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return Path.GetFullPath(Environment.ExpandEnvironmentVariables(value.Trim()));
    }
}

public sealed class ExcelMacroJobConfig
{
    public string Name { get; set; } = string.Empty;
    public string ExcelFilePath { get; set; } = string.Empty;
    public string MacroName { get; set; } = string.Empty;

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(Name))
        {
            throw new InvalidOperationException("Cada item de ServiceConfig.ExcelJobs precisa de um Name.");
        }

        if (string.IsNullOrWhiteSpace(ExcelFilePath))
        {
            throw new InvalidOperationException($"O job '{Name}' precisa de ExcelFilePath.");
        }

        if (string.IsNullOrWhiteSpace(MacroName))
        {
            throw new InvalidOperationException($"O job '{Name}' precisa de MacroName.");
        }
    }

    public void Normalize()
    {
        Name = Name.Trim();
        MacroName = MacroName.Trim();

        if (!string.IsNullOrWhiteSpace(ExcelFilePath))
        {
            ExcelFilePath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(ExcelFilePath.Trim()));
        }
    }
}
