using System;
using System.Threading.Tasks;
using System.IO;
using System.Net.Sockets;
using System.Threading;
namespace MedControl;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // Single-instance: cria mutex nomeado. Se já existir, sinaliza restauração e sai.
        bool created;
        using var singleMutex = new Mutex(true, "MedControlSingleInstanceMutex", out created);
        var restoreEvent = new EventWaitHandle(false, EventResetMode.AutoReset, "MedControl_Restore_Event");
        if (!created)
        {
            try { restoreEvent.Set(); } catch { }
            return; // evita segunda instância
        }
        // Handlers globais para capturar exceções que poderiam encerrar o processo silenciosamente
        try
        {
            Application.ThreadException += (s, e) => LogAndShowCrash("ThreadException", e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                LogAndShowCrash("UnhandledException", ex);
            };
                TaskScheduler.UnobservedTaskException += (s, e) =>
                {
                    // Apenas log, não mostrar popup para exceções não observadas (normal em shutdown)
                    SafeLog("UnobservedTaskException", e.Exception?.ToString() ?? "(sem detalhes)");
                    try { e.SetObserved(); } catch { }
                };
            Application.ApplicationExit += (s, e) =>
            {
                _shuttingDown = true;
                SafeLog("ApplicationExit", "normal shutdown");
            };
        }
        catch { }
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        // Ajusta Visual Styles de acordo com o tema salvo (clássico desliga estilos visuais)
        try
        {
            AppConfig.Load();
            MedControl.Database.Setup();
            try { MedControl.SyncService.Start(); } catch { }
            try { if (MedControl.GroupConfig.Mode == MedControl.GroupMode.Host) MedControl.GroupHost.Start(); else MedControl.GroupHost.Stop(); } catch { }
            // Inicia coordenação de papéis (host/client) com eleição automática e operação offline-first
            try { MedControl.GroupCoordinator.Start(); } catch { }
            // Splash + pré-download de dados para uso offline
            try { MedControl.InitialSync.RunWithUiSplash(); } catch { }
            // Auto-start: garante registro de inicialização com o Windows (escopo usuário)
            try { MedControl.StartupHelper.EnsureAutoStart(); } catch { }
            try
            {
                if (MedControl.GroupConfig.Mode == MedControl.GroupMode.Client)
                {
                    // Tenta conectar ao Host em background na inicialização
                        // Usa helper para evitar popup em exceções de tarefas em background
                        SafeRun(() => Task.Run(() => { try { MedControl.GroupClient.Ping(out _); } catch { } }));
                }
            }
            catch { }
            var themeRaw = (MedControl.Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
            var theme = themeRaw switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => themeRaw };
            if (theme == "classico" || theme == "alto_contraste" || theme == "terminal")
                Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled;
            else
                Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.ClientAndNonClientAreasEnabled;
        }
        catch { }
        // Checagem de atualização (silenciosa) em background
        try { MedControl.UpdateService.CheckSilentlyInBackground(); } catch { }
    Application.Run(new Form1());
    try { MedControl.GroupCoordinator.Stop(); } catch { }
    try { MedControl.GroupHost.Stop(); } catch { }
    try { MedControl.SyncService.Stop(); } catch { }
    }    

    private static void LogAndShowCrash(string kind, Exception? ex)
    {
        try
        {
            var msg = ex == null ? "(sem detalhes)" : ex.ToString();
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {kind}: {msg}\n";
            var logPath = System.IO.Path.Combine(AppContext.BaseDirectory, "AppErrors.log");
            System.IO.File.AppendAllText(logPath, line);
            if (!IsShuttingDown())
            {
                if (!IsTransientNetwork(ex) && AllowCrashPopup())
                {
                    MessageBox.Show($"Erro inesperado: {ex?.Message}\n\nDetalhes salvos em AppErrors.log", "Falha", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        catch { }
    }

    private static bool _shuttingDown;
    internal static bool IsShuttingDown() => _shuttingDown;
    internal static void MarkShuttingDown(string reason)
    {
        _shuttingDown = true;
        SafeLog("MarkShuttingDown", reason);
    }
    private static void SafeLog(string kind, string message)
    {
        try
        {
            var logPath = System.IO.Path.Combine(AppContext.BaseDirectory, "AppErrors.log");
            System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {kind}: {message}\n");
        }
        catch { }
    }
    private static DateTime _lastCrashPopupUtc = DateTime.MinValue;
    private static int _popupCount = 0;
    private static bool AllowCrashPopup()
    {
        var now = DateTime.UtcNow;
        // Permite no máximo 1 popup a cada 60 segundos e até 3 por sessão
        if (_popupCount >= 3) return false;
        if ((now - _lastCrashPopupUtc).TotalSeconds < 60) return false;
        _lastCrashPopupUtc = now;
        _popupCount++;
        return true;
    }
    private static bool IsTransientNetwork(Exception? ex)
    {
        if (ex == null) return false;
        // Verifica cadeia de InnerException por erros de socket / transporte
        Exception? cur = ex;
        while (cur != null)
        {
            if (cur is SocketException) return true;
            if (cur is IOException && cur.Message.Contains("conexão estabelecida", StringComparison.OrdinalIgnoreCase)) return true;
            if (cur.Message.Contains("unable to write data to the transport connection", StringComparison.OrdinalIgnoreCase)) return true;
            if (cur.Message.Contains("foi anulada pelo software", StringComparison.OrdinalIgnoreCase)) return true;
            cur = cur.InnerException;
        }
        return false;
    }
        // Helper para anexar continuação que observa exceções evitando UnobservedTaskException popups
        private static void SafeRun(Func<Task> start)
        {
            try
            {
                var t = start();
                if (t == null) return;
                t.ContinueWith(ct =>
                {
                    SafeLog("BackgroundTaskFault", ct.Exception?.ToString() ?? "(sem detalhes)");
                }, TaskContinuationOptions.OnlyOnFaulted);
            }
            catch (Exception ex)
            {
                SafeLog("SafeRunSetupFault", ex.ToString());
            }
        }
}