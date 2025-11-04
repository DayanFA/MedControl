using System;
using System.Threading.Tasks;
namespace MedControl;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
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
            try
            {
                if (MedControl.GroupConfig.Mode == MedControl.GroupMode.Client)
                {
                    // Tenta conectar ao Host em background na inicialização
                    Task.Run(() =>
                    {
                        try { MedControl.GroupClient.Ping(out _); } catch { }
                    });
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
}