using System.Diagnostics;

namespace ExcelMacroRunner;

public partial class MainForm : Form
{
    private const int MaxLogLines = 200;

    private readonly MacroRunnerService _runnerService;
    private readonly FileLogger _logger;
    private readonly Queue<string> _recentLogs = new();

    public MainForm()
    {
        InitializeComponent();

        var settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
        _logger = new FileLogger(Path.Combine(AppContext.BaseDirectory, "Logs", "ExcelMacroRunner"));
        _runnerService = new MacroRunnerService(settingsPath, _logger);

        _logger.LogWritten += Logger_LogWritten;
        _runnerService.StatusChanged += RunnerService_StatusChanged;

        var executableIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (executableIcon is not null)
        {
            Icon = executableIcon;
        }

        lblConfigPathValue.Text = settingsPath;
        ApplySnapshot(_runnerService.GetSnapshot());
        RefreshPathsFromConfig();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _runnerService.Dispose();
        base.OnFormClosing(e);
    }

    private async void btnStart_Click(object sender, EventArgs e)
    {
        await RunUiCommandAsync(() => _runnerService.StartAsync());
    }

    private async void btnStop_Click(object sender, EventArgs e)
    {
        await RunUiCommandAsync(() => _runnerService.StopAsync(false, "Parada solicitada pelo usuário."));
    }

    private async void btnRestartNow_Click(object sender, EventArgs e)
    {
        await RunUiCommandAsync(() => _runnerService.RestartNowAsync());
    }

    private async void btnKillExcel_Click(object sender, EventArgs e)
    {
        await RunUiCommandAsync(() => _runnerService.KillExcelProcessesAsync());
    }

    private void btnOpenLogFolder_Click(object sender, EventArgs e)
    {
        try
        {
            var logFolder = _logger.CurrentLogFolder;
            Directory.CreateDirectory(logFolder);

            Process.Start(new ProcessStartInfo
            {
                FileName = logFolder,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Não foi possível abrir a pasta de log.{Environment.NewLine}{ex.Message}",
                "Erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void RunnerService_StatusChanged(object? sender, RunnerStatusSnapshot snapshot)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => ApplySnapshot(snapshot)));
            return;
        }

        ApplySnapshot(snapshot);
    }

    private void Logger_LogWritten(object? sender, string line)
    {
        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => AppendLogLine(line)));
            return;
        }

        AppendLogLine(line);
    }

    private async Task RunUiCommandAsync(Func<Task> action)
    {
        SetButtonsEnabled(false);

        try
        {
            await action();
            RefreshPathsFromConfig();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                "Erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            SetButtonsEnabled(true);
        }
    }

    private void ApplySnapshot(RunnerStatusSnapshot snapshot)
    {
        lblStatusValue.Text = snapshot.State switch
        {
            RunnerState.Stopped => "Parado",
            RunnerState.Running => "Rodando",
            RunnerState.Stopping => "Parando",
            RunnerState.Restarting => "Reiniciando",
            RunnerState.Error => "Erro",
            _ => snapshot.State.ToString()
        };

        lblLastExecutionValue.Text = FormatDateTime(snapshot.LastExecution);
        lblHtmlLastUpdateValue.Text = FormatDateTime(snapshot.LastHtmlUpdate);
        lblHtmlPathValue.Text = snapshot.HtmlPath;
        lblNextActionValue.Text = snapshot.NextAction;
        txtCurrentMessage.Text = snapshot.Message;
    }

    private void RefreshPathsFromConfig()
    {
        var config = _runnerService.CurrentConfig;
        lblLogFolderValue.Text = _logger.CurrentLogFolder;
        lblHtmlPathValue.Text = config.HtmlPath;
    }

    private void AppendLogLine(string line)
    {
        _recentLogs.Enqueue(line);
        while (_recentLogs.Count > MaxLogLines)
        {
            _recentLogs.Dequeue();
        }

        txtLogs.Text = string.Join(Environment.NewLine, _recentLogs);
        txtLogs.SelectionStart = txtLogs.TextLength;
        txtLogs.ScrollToCaret();
    }

    private void SetButtonsEnabled(bool enabled)
    {
        btnStart.Enabled = enabled;
        btnStop.Enabled = enabled;
        btnRestartNow.Enabled = enabled;
        btnKillExcel.Enabled = enabled;
        btnOpenLogFolder.Enabled = enabled;
    }

    private static string FormatDateTime(DateTime? value)
    {
        return value.HasValue
            ? value.Value.ToString("yyyy-MM-dd HH:mm:ss")
            : "Sem informação";
    }
}
