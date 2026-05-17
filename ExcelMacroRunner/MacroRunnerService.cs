using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacroRunner;

public sealed class MacroRunnerService : IDisposable
{
    private readonly string _settingsPath;
    private readonly FileLogger _logger;
    private readonly SemaphoreSlim _lifecycleLock = new(1, 1);
    private readonly object _snapshotLock = new();

    private CancellationTokenSource? _cancellationTokenSource;
    private Thread? _workerThread;
    private ServiceConfig _config;
    private RunnerStatusSnapshot _snapshot;

    public MacroRunnerService(string settingsPath, FileLogger logger)
    {
        _settingsPath = settingsPath;
        _logger = logger;

        if (ServiceConfig.TryLoad(settingsPath, out var loadedConfig, out var loadError))
        {
            _config = loadedConfig!;
            _logger.UpdateLogFolder(_config.LogFolder);
        }
        else
        {
            _config = ServiceConfig.CreateFallback(AppContext.BaseDirectory);
            _logger.Warning($"Falha ao carregar configuração na inicialização: {loadError}");
        }

        _snapshot = new RunnerStatusSnapshot(
            RunnerState.Stopped,
            null,
            TryGetHtmlLastWrite(_config.HtmlPath),
            "Aguardando início",
            "Aplicação pronta.",
            _config.HtmlPath);
    }

    public event EventHandler<RunnerStatusSnapshot>? StatusChanged;

    public bool IsRunning => _workerThread is { IsAlive: true };

    public ServiceConfig CurrentConfig
    {
        get
        {
            lock (_snapshotLock)
            {
                return _config;
            }
        }
    }

    public RunnerStatusSnapshot GetSnapshot()
    {
        lock (_snapshotLock)
        {
            return _snapshot;
        }
    }

    public async Task StartAsync()
    {
        await _lifecycleLock.WaitAsync().ConfigureAwait(false);

        try
        {
            if (IsRunning)
            {
                _logger.Info("A solicitação de iniciar foi ignorada porque o loop já está rodando.");
                return;
            }

            if (!TryReloadConfig(out var errorMessage))
            {
                UpdateSnapshot(
                    RunnerState.Error,
                    nextAction: "Corrigir appsettings.json",
                    message: errorMessage);

                _logger.Error(errorMessage);
                return;
            }

            _cancellationTokenSource = new CancellationTokenSource();
            _workerThread = new Thread(() => RunLoop(_cancellationTokenSource.Token))
            {
                IsBackground = true,
                Name = "ExcelMacroRunner.Worker"
            };
            _workerThread.SetApartmentState(ApartmentState.STA);

            UpdateSnapshot(
                RunnerState.Running,
                nextAction: "Iniciando loop contínuo",
                message: "Loop iniciado em background.");

            _workerThread.Start();
            _logger.Info("Loop contínuo iniciado.");
        }
        finally
        {
            _lifecycleLock.Release();
        }
    }

    public async Task StopAsync(bool killExcel, string reason)
    {
        await _lifecycleLock.WaitAsync().ConfigureAwait(false);

        try
        {
            await StopCoreAsync(killExcel, reason).ConfigureAwait(false);
        }
        finally
        {
            _lifecycleLock.Release();
        }
    }

    public async Task RestartNowAsync()
    {
        await _lifecycleLock.WaitAsync().ConfigureAwait(false);

        try
        {
            UpdateSnapshot(
                RunnerState.Restarting,
                nextAction: "Parando loop atual",
                message: "Reinício manual solicitado.");

            var killExcel = CurrentConfig.KillExcelOnTimeout;
            var stopped = await StopCoreAsync(killExcel, "Reinício manual solicitado pelo usuário.").ConfigureAwait(false);
            if (!stopped)
            {
                return;
            }

            StartWorkerThread();
        }
        finally
        {
            _lifecycleLock.Release();
        }
    }

    public Task KillExcelProcessesAsync()
    {
        return Task.Run(() =>
        {
            var killed = KillAllExcelProcesses("Solicitação manual do usuário.");
            if (killed == 0)
            {
                _logger.Info("Nenhum processo EXCEL.EXE foi encontrado para encerrar.");
            }
        });
    }

    public void Dispose()
    {
        try
        {
            StopAsync(false, "Encerramento da aplicação.").GetAwaiter().GetResult();
        }
        catch
        {
            // Evita falha no fechamento da aplicação.
        }

        _cancellationTokenSource?.Dispose();
        _lifecycleLock.Dispose();
    }

    private async Task<bool> StopCoreAsync(bool killExcel, string reason)
    {
        if (!IsRunning)
        {
            UpdateSnapshot(
                RunnerState.Stopped,
                nextAction: "Aguardando início",
                message: "O loop já estava parado.");

            return true;
        }

        _logger.Info(reason);
        UpdateSnapshot(
            RunnerState.Stopping,
            nextAction: "Encerrando thread de execução",
            message: reason);

        var thread = _workerThread;
        _cancellationTokenSource?.Cancel();

        if (killExcel)
        {
            KillAllExcelProcesses("Parada solicitada com encerramento de EXCEL.EXE.");
        }

        if (thread is not null)
        {
            var stopped = await Task.Run(() => thread.Join(TimeSpan.FromSeconds(30))).ConfigureAwait(false);
            if (!stopped)
            {
                _logger.Warning("A thread de execução não encerrou em 30 segundos.");
                UpdateSnapshot(
                    RunnerState.Error,
                    nextAction: "Loop ainda finalizando",
                    message: "A thread de execução ainda está ativa. Use 'Matar Excel' se necessário.");

                return false;
            }
        }

        _cancellationTokenSource?.Dispose();
        _cancellationTokenSource = null;
        _workerThread = null;

        UpdateSnapshot(
            RunnerState.Stopped,
            nextAction: "Aguardando início",
            message: "Loop parado.");

        _logger.Info("Loop contínuo parado.");
        return true;
    }

    private void StartWorkerThread()
    {
        if (!TryReloadConfig(out var errorMessage))
        {
            UpdateSnapshot(
                RunnerState.Error,
                nextAction: "Corrigir appsettings.json",
                message: errorMessage);

            _logger.Error(errorMessage);
            return;
        }

        _cancellationTokenSource = new CancellationTokenSource();
        _workerThread = new Thread(() => RunLoop(_cancellationTokenSource.Token))
        {
            IsBackground = true,
            Name = "ExcelMacroRunner.Worker"
        };
        _workerThread.SetApartmentState(ApartmentState.STA);

        UpdateSnapshot(
            RunnerState.Running,
            nextAction: "Executando novo loop",
            message: "Loop iniciado em background.");

        _workerThread.Start();
        _logger.Info("Loop contínuo iniciado.");
    }

    private void RunLoop(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                ExecuteCycle(cancellationToken);

                UpdateSnapshot(
                    RunnerState.Running,
                    nextAction: $"Aguardando {_config.LoopDelaySeconds} segundo(s) para o próximo ciclo",
                    message: "Ciclo concluído com sucesso.");

                if (cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(_config.LoopDelaySeconds)))
                {
                    break;
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.Exception("Falha no ciclo principal", ex);

                UpdateSnapshot(
                    RunnerState.Error,
                    nextAction: $"Nova tentativa em {_config.LoopDelaySeconds} segundo(s)",
                    message: ex.Message);

                if (cancellationToken.WaitHandle.WaitOne(TimeSpan.FromSeconds(_config.LoopDelaySeconds)))
                {
                    break;
                }

                UpdateSnapshot(
                    RunnerState.Restarting,
                    nextAction: "Preparando nova tentativa automática",
                    message: "Reiniciando ciclo após erro.");
            }
        }
    }

    private void ExecuteCycle(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateConfiguredPaths();

        var cycleStart = DateTime.Now;
        var watchdogState = new HtmlWatchdogState(cycleStart);
        using var watchdogTimer = new System.Threading.Timer(
            _ => CheckHtmlTimeout(watchdogState),
            null,
            TimeSpan.FromSeconds(15),
            TimeSpan.FromSeconds(15));

        UpdateSnapshot(
            RunnerState.Running,
            nextAction: "Validando caminhos configurados",
            message: "Iniciando novo ciclo.");

        ExecuteBatchFile(cancellationToken);

        foreach (var job in _config.ExcelJobs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ExecuteExcelJob(job, cancellationToken, watchdogState);
        }

        var lastExecution = DateTime.Now;
        var htmlLastWrite = TryGetHtmlLastWrite(_config.HtmlPath);

        UpdateSnapshot(
            RunnerState.Running,
            lastExecution: lastExecution,
            lastHtmlUpdate: htmlLastWrite,
            nextAction: "Verificando atualização do HTML",
            message: "Macros concluídas; validando arquivo HTML.");

        EnsureHtmlIsFresh(cycleStart, htmlLastWrite);
    }

    private void ExecuteBatchFile(CancellationToken cancellationToken)
    {
        UpdateSnapshot(
            RunnerState.Running,
            nextAction: $"Executando BAT: {_config.BatPath}",
            message: "Iniciando etapa BAT.");

        var commandShell = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
        var oemEncoding = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = commandShell,
                Arguments = $"/c \"{_config.BatPath}\"",
                WorkingDirectory = Path.GetDirectoryName(_config.BatPath) ?? AppContext.BaseDirectory,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = oemEncoding,
                StandardErrorEncoding = oemEncoding,
                CreateNoWindow = true
            }
        };

        process.OutputDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                _logger.Info($"BAT> {args.Data}");
            }
        };

        process.ErrorDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                _logger.Warning($"BAT! {args.Data}");
            }
        };

        if (!process.Start())
        {
            throw new InvalidOperationException("Não foi possível iniciar o arquivo BAT.");
        }

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            while (!process.WaitForExit(500))
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            process.WaitForExit();
            cancellationToken.ThrowIfCancellationRequested();

            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException($"O BAT terminou com código de saída {process.ExitCode}.");
            }

            _logger.Info("Arquivo BAT executado com sucesso.");
        }
        catch
        {
            TryKillProcess(process, "Encerrando BAT após falha ou cancelamento.");
            throw;
        }
    }

    private void ExecuteExcelJob(
        ExcelMacroJobConfig job,
        CancellationToken cancellationToken,
        HtmlWatchdogState watchdogState)
    {
        Excel.Application? excelApplication = null;
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? workbook = null;

        UpdateSnapshot(
            RunnerState.Running,
            nextAction: $"Abrindo workbook: {job.Name}",
            message: $"Preparando Excel para '{job.ExcelFilePath}'.");

        try
        {
            excelApplication = new Excel.Application
            {
                Visible = _config.ShowExcel,
                DisplayAlerts = false,
                ScreenUpdating = _config.ShowExcel,
                EnableEvents = false,
                AskToUpdateLinks = false
            };

            workbooks = excelApplication.Workbooks;
            cancellationToken.ThrowIfCancellationRequested();

            workbook = workbooks.Open(
                job.ExcelFilePath,
                UpdateLinks: 0,
                ReadOnly: false,
                IgnoreReadOnlyRecommended: true,
                AddToMru: false);

            UpdateSnapshot(
                RunnerState.Running,
                nextAction: $"Executando macro: {job.MacroName}",
                message: $"Workbook aberto para o job '{job.Name}'.");

            excelApplication.Run(job.MacroName);
            cancellationToken.ThrowIfCancellationRequested();

            if (watchdogState.IsTriggered)
            {
                throw new InvalidOperationException("O watchdog do HTML disparou durante a execução da macro.");
            }

            _logger.Info($"Macro executada com sucesso: {job.Name} -> {job.MacroName}");
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            if (watchdogState.IsTriggered)
            {
                throw new InvalidOperationException("Execução do Excel interrompida por timeout de atualização do HTML.", ex);
            }

            _logger.Exception($"Falha ao executar o job '{job.Name}'", ex);
            KillAllExcelProcesses($"Falha no job '{job.Name}'.");
            throw new InvalidOperationException($"Falha no job '{job.Name}': {ex.Message}", ex);
        }
        finally
        {
            TryCloseWorkbook(workbook);
            TryQuitExcel(excelApplication);

            ReleaseComObject(workbook);
            ReleaseComObject(workbooks);
            ReleaseComObject(excelApplication);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private void ValidateConfiguredPaths()
    {
        if (!File.Exists(_config.BatPath))
        {
            throw new FileNotFoundException("Arquivo BAT não encontrado.", _config.BatPath);
        }

        foreach (var job in _config.ExcelJobs)
        {
            if (!File.Exists(job.ExcelFilePath))
            {
                throw new FileNotFoundException($"Workbook do job '{job.Name}' não encontrado.", job.ExcelFilePath);
            }
        }

        var htmlDirectory = Path.GetDirectoryName(_config.HtmlPath);
        if (string.IsNullOrWhiteSpace(htmlDirectory) || !Directory.Exists(htmlDirectory))
        {
            _logger.Warning($"A pasta do HTML ainda não existe: {_config.HtmlPath}");
        }
    }

    private void CheckHtmlTimeout(HtmlWatchdogState watchdogState)
    {
        if (watchdogState.IsTriggered)
        {
            return;
        }

        DateTime? htmlLastWrite = null;
        if (File.Exists(_config.HtmlPath))
        {
            htmlLastWrite = File.GetLastWriteTime(_config.HtmlPath);
            UpdateSnapshot(lastHtmlUpdate: htmlLastWrite);
        }

        var freshnessReference = htmlLastWrite.HasValue && htmlLastWrite.Value > watchdogState.CycleStart
            ? htmlLastWrite.Value
            : watchdogState.CycleStart;

        if (DateTime.Now - freshnessReference <= TimeSpan.FromMinutes(_config.HtmlMaxAgeMinutes))
        {
            return;
        }

        if (!watchdogState.TryTrigger())
        {
            return;
        }

        var message = $"HTML sem atualização há mais de {_config.HtmlMaxAgeMinutes} minuto(s).";
        _logger.Error(message);

        UpdateSnapshot(
            RunnerState.Restarting,
            nextAction: "Matando EXCEL.EXE por timeout do HTML",
            message: message);

        if (_config.KillExcelOnTimeout)
        {
            KillAllExcelProcesses("Timeout detectado no HTML.");
        }
    }

    private void EnsureHtmlIsFresh(DateTime cycleStart, DateTime? htmlLastWrite)
    {
        if (!htmlLastWrite.HasValue)
        {
            _logger.Warning("O arquivo HTML ainda não existe. O loop continuará tentando.");
            return;
        }

        var freshnessReference = htmlLastWrite.Value > cycleStart
            ? htmlLastWrite.Value
            : cycleStart;

        if (DateTime.Now - freshnessReference <= TimeSpan.FromMinutes(_config.HtmlMaxAgeMinutes))
        {
            return;
        }

        var message = $"O HTML está desatualizado há mais de {_config.HtmlMaxAgeMinutes} minuto(s).";
        _logger.Error(message);

        if (_config.KillExcelOnTimeout)
        {
            KillAllExcelProcesses("HTML desatualizado após o ciclo.");
        }

        throw new InvalidOperationException(message);
    }

    private bool TryReloadConfig(out string errorMessage)
    {
        try
        {
            var loadedConfig = ServiceConfig.Load(_settingsPath);

            lock (_snapshotLock)
            {
                _config = loadedConfig;
            }

            _logger.UpdateLogFolder(loadedConfig.LogFolder);
            UpdateSnapshot(
                lastHtmlUpdate: TryGetHtmlLastWrite(loadedConfig.HtmlPath),
                nextAction: "Configuração carregada",
                message: "appsettings.json lido com sucesso.");

            errorMessage = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = $"Erro ao carregar appsettings.json: {ex.Message}";
            return false;
        }
    }

    private void UpdateSnapshot(
        RunnerState? state = null,
        DateTime? lastExecution = null,
        DateTime? lastHtmlUpdate = null,
        string? nextAction = null,
        string? message = null)
    {
        RunnerStatusSnapshot snapshot;

        lock (_snapshotLock)
        {
            snapshot = _snapshot with
            {
                State = state ?? _snapshot.State,
                LastExecution = lastExecution ?? _snapshot.LastExecution,
                LastHtmlUpdate = lastHtmlUpdate ?? _snapshot.LastHtmlUpdate,
                NextAction = nextAction ?? _snapshot.NextAction,
                Message = message ?? _snapshot.Message,
                HtmlPath = _config.HtmlPath
            };

            _snapshot = snapshot;
        }

        StatusChanged?.Invoke(this, snapshot);
    }

    private DateTime? TryGetHtmlLastWrite(string htmlPath)
    {
        if (string.IsNullOrWhiteSpace(htmlPath) || !File.Exists(htmlPath))
        {
            return null;
        }

        return File.GetLastWriteTime(htmlPath);
    }

    private int KillAllExcelProcesses(string reason)
    {
        var killed = 0;
        var processes = Process.GetProcessesByName("EXCEL");

        foreach (var process in processes)
        {
            try
            {
                process.Kill();
                process.WaitForExit(5000);
                killed++;
            }
            catch (Exception ex)
            {
                _logger.Exception($"Falha ao encerrar o processo EXCEL.EXE PID {process.Id}", ex);
            }
            finally
            {
                process.Dispose();
            }
        }

        if (killed > 0)
        {
            _logger.Warning($"{killed} processo(s) EXCEL.EXE encerrado(s). Motivo: {reason}");
        }

        return killed;
    }

    private static void TryCloseWorkbook(Excel.Workbook? workbook)
    {
        if (workbook is null)
        {
            return;
        }

        try
        {
            workbook.Close(SaveChanges: false);
        }
        catch
        {
            // A limpeza continua mesmo se o Excel já estiver em estado inconsistente.
        }
    }

    private static void TryQuitExcel(Excel.Application? excelApplication)
    {
        if (excelApplication is null)
        {
            return;
        }

        try
        {
            excelApplication.Quit();
        }
        catch
        {
            // A limpeza continua mesmo se o Excel já estiver em estado inconsistente.
        }
    }

    private static void ReleaseComObject(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        try
        {
            Marshal.ReleaseComObject(comObject);
        }
        catch
        {
            // Ignora falhas de liberação para evitar quebrar o finally.
        }
    }

    private void TryKillProcess(Process process, string reason)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill();
                process.WaitForExit(5000);
                _logger.Warning(reason);
            }
        }
        catch (Exception ex)
        {
            _logger.Exception("Falha ao encerrar processo externo", ex);
        }
    }

    private sealed class HtmlWatchdogState
    {
        private int _triggered;

        public HtmlWatchdogState(DateTime cycleStart)
        {
            CycleStart = cycleStart;
        }

        public DateTime CycleStart { get; }

        public bool IsTriggered => Volatile.Read(ref _triggered) == 1;

        public bool TryTrigger()
        {
            return Interlocked.Exchange(ref _triggered, 1) == 0;
        }
    }
}

public enum RunnerState
{
    Stopped,
    Running,
    Stopping,
    Restarting,
    Error
}

public sealed record RunnerStatusSnapshot(
    RunnerState State,
    DateTime? LastExecution,
    DateTime? LastHtmlUpdate,
    string NextAction,
    string Message,
    string HtmlPath);
