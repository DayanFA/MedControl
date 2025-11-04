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
            var themeRaw = (MedControl.Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
            var theme = themeRaw switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => themeRaw };
            if (theme == "classico" || theme == "alto_contraste" || theme == "terminal")
                Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled;
            else
                Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.ClientAndNonClientAreasEnabled;
        }
        catch { }
        Application.Run(new Form1());
    }    
}