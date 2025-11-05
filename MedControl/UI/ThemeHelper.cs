using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;

namespace MedControl.UI
{
    public static class ThemeHelper
    {
        private static Icon? _appIcon;
        
        // Helpers to detect tagged controls
        private static bool HasKeepBackcolorTag(Control c)
        {
            try { return (c.Tag as string)?.IndexOf("keep-backcolor", StringComparison.OrdinalIgnoreCase) >= 0; } catch { return false; }
        }

        private static bool IsAccentButton(Control c)
        {
            return c is Button b && (b.Tag as string) == "accent";
        }

        private static bool IsKeepBackcolorControl(Control c)
        {
            if (HasKeepBackcolorTag(c)) return true;
            try { return c is MedControl.UI.SquareCardPanel; } catch { return false; }
        }

        // Tenta carregar o ícone do app a partir de Assets/app.ico (preferido) ou Assets/app.png
        private static Icon? TryLoadAppIcon()
        {
            if (_appIcon != null) return _appIcon;
            try
            {
                var baseDir = AppContext.BaseDirectory;
                // Procura primeiro por ICO
                var icoPath = Path.Combine(baseDir, "Assets", "app.ico");
                if (File.Exists(icoPath))
                {
                    _appIcon = new Icon(icoPath);
                    return _appIcon;
                }

                // Como fallback, tenta PNG -> Icon em runtime
                var pngPath = Path.Combine(baseDir, "Assets", "app.png");
                if (File.Exists(pngPath))
                {
                    using var bmp = new Bitmap(pngPath);
                    // reduz pra 256 ou 32 se necessário
                    var size = bmp.Width > 256 || bmp.Height > 256 ? 256 : Math.Max(32, Math.Min(bmp.Width, bmp.Height));
                    using var resized = new Bitmap(bmp, new Size(size, size));
                    var hIcon = resized.GetHicon();
                    _appIcon = Icon.FromHandle(hIcon);
                    return _appIcon;
                }
            }
            catch { }
            return null;
        }

        private static void ApplyAppIcon(Form form)
        {
            try
            {
                var ic = TryLoadAppIcon();
                if (ic != null)
                {
                    form.Icon = ic;
                }
            }
            catch { }
        }

        public static void ApplyCurrentTheme(Form form)
        {
            try
            {
                var themeRaw = (MedControl.Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                var theme = NormalizeThemeKey(themeRaw);

                // Define o ícone do formulário, se disponível
                try { ApplyAppIcon(form); } catch { }

                // Primeiro: remover efeitos Mica se não for tema Mica
                if (theme != "mica")
                {
                    try { MedControl.UI.FluentEffects.ResetWin11Backdrop(form); } catch { }
                }

                // Resetar estilos para evitar resíduos entre trocas de tema
                // Preservamos apenas o que estiver marcado por Tag (accent/keep-backcolor)
                try { ResetStyles(form); } catch { }

                switch (theme)
                {
                    case "classico":
                        try { Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled; } catch { }
                        ApplyClassicStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                    case "claro":
                        ApplyLightStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                    case "escuro":
                        ApplyDarkStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                    case "mica":
                        try { MedControl.UI.FluentEffects.ApplyWin11Mica(form); } catch { }
                        ApplyMicaStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                    case "alto_contraste":
                        try { Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled; } catch { }
                        ApplyHighContrastStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                    case "terminal":
                        try { Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled; } catch { }
                        ApplyTerminalStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                    default:
                        // Tema 'padrao' e legados: aplicar estilo padrão (fontes do sistema) e menus
                        ApplyDefaultStyle(form);
                        try { StyleMenusForTheme(form, theme); } catch { }
                        break;
                }
            }
            catch { }
        }

        public static void ApplyClassicIfNeeded(Form form)
        {
            try
            {
                var theme = NormalizeThemeKey((MedControl.Database.GetConfig("theme") ?? "padrao").ToLowerInvariant());
                if (theme == "classico")
                {
                    try { Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled; } catch { }
                    ApplyClassicStyle(form);
                }
            }
            catch { }
        }

        public static void ApplyClassicStyle(Control root)
        {
            try
            {
                var classicFont = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular);
                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton) c.Font = classicFont;
                    }
                    catch { }

                    if (c is TextBox || c is ComboBox || c is ListBox)
                    {
                        c.BackColor = SystemColors.Window;
                        c.ForeColor = SystemColors.WindowText;
                    }
                    else if (c is DataGridView dgv)
                    {
                        c.BackColor = SystemColors.Control;
                        c.ForeColor = SystemColors.ControlText;
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = SystemColors.Control;
                        dgv.GridColor = SystemColors.ControlDark;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
                        dgv.ColumnHeadersDefaultCellStyle.Font = classicFont;
                        dgv.DefaultCellStyle.BackColor = SystemColors.Window;
                        dgv.DefaultCellStyle.ForeColor = SystemColors.WindowText;
                        dgv.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
                        dgv.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
                    }
                    else if (c is Button btn)
                    {
                        btn.FlatStyle = FlatStyle.Standard;
                        c.BackColor = SystemColors.Control;
                        c.ForeColor = SystemColors.ControlText;
                    }
                    else if (IsKeepBackcolorControl(c))
                    {
                        // Preserva o fundo (badges/status/cartões); só ajusta a cor do texto
                        c.ForeColor = SystemColors.ControlText;
                    }
                    else
                    {
                        c.BackColor = SystemColors.Control;
                        c.ForeColor = SystemColors.ControlText;
                    }

                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }

        private static void ApplyMicaStyle(Control root)
        {
            try
            {
                var font = new Font("Segoe UI", 10f, FontStyle.Regular);
                var bg = Color.FromArgb(245, 246, 250); // leve cinza azulado
                var header = Color.FromArgb(238, 240, 244);
                var accent = Color.FromArgb(0, 120, 215);

                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton) c.Font = font;
                    }
                    catch { }

                    if (c is DataGridView dgv)
                    {
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = bg;
                        dgv.GridColor = Color.FromArgb(223, 226, 231);
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = header;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
                        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 10f);
                        dgv.DefaultCellStyle.BackColor = Color.White;
                        dgv.DefaultCellStyle.ForeColor = Color.Black;
                        dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(220, 235, 255);
                        dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                    }
                    else if (c is TextBox or ComboBox or ListBox)
                    {
                        c.BackColor = Color.White;
                        c.ForeColor = Color.Black;
                    }
                    else if (c is Button btn && (btn.Tag as string) != "accent")
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderSize = 0;
                        btn.BackColor = accent;
                        btn.ForeColor = Color.White;
                    }
                    else if (IsKeepBackcolorControl(c))
                    {
                        // preserva cor de fundo do badge/status, mas ajusta a cor do texto para o tema claro
                        c.ForeColor = Color.Black;
                    }
                    else
                    {
                        c.BackColor = bg;
                        c.ForeColor = Color.Black;
                    }

                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }

        private static void ApplyHighContrastStyle(Control root)
        {
            try
            {
                var font = new Font("Segoe UI", 11f, FontStyle.Bold);
                var bg = Color.Black; var fg = Color.White; var border = Color.White;
                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton) c.Font = font;
                    }
                    catch { }
                    if (c is DataGridView dgv)
                    {
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = bg;
                        dgv.GridColor = border;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = bg;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = fg;
                        dgv.ColumnHeadersDefaultCellStyle.Font = font;
                        dgv.DefaultCellStyle.BackColor = bg;
                        dgv.DefaultCellStyle.ForeColor = fg;
                        dgv.DefaultCellStyle.SelectionBackColor = Color.Yellow;
                        dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                    }
                    else if (c is Button btn && (btn.Tag as string) != "accent")
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderSize = 2;
                        btn.FlatAppearance.BorderColor = border;
                        btn.BackColor = bg; btn.ForeColor = fg;
                    }
                    else if (c is TextBox or ComboBox or ListBox)
                    {
                        c.BackColor = bg; c.ForeColor = fg;
                    }
                    else if (IsKeepBackcolorControl(c))
                    {
                        // preserva badges, texto branco no alto contraste
                        c.ForeColor = fg;
                    }
                    else
                    {
                        c.BackColor = bg; c.ForeColor = fg;
                    }
                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }

        private static void ApplyTerminalStyle(Control root)
        {
            try
            {
                var font = new Font("Consolas", 10f, FontStyle.Regular);
                var bg = Color.Black; var fg = Color.FromArgb(0, 255, 128);
                var border = Color.FromArgb(0, 180, 90);

                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton) c.Font = font;
                    }
                    catch { }
                    if (c is DataGridView dgv)
                    {
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = bg; dgv.GridColor = border;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = bg;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = fg;
                        dgv.ColumnHeadersDefaultCellStyle.Font = font;
                        dgv.DefaultCellStyle.BackColor = bg;
                        dgv.DefaultCellStyle.ForeColor = fg;
                        dgv.DefaultCellStyle.SelectionBackColor = border;
                        dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                    }
                    else if (c is Button btn && (btn.Tag as string) != "accent")
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderSize = 1;
                        btn.FlatAppearance.BorderColor = border;
                        btn.BackColor = bg; btn.ForeColor = fg;
                    }
                    else if (c is TextBox or ComboBox or ListBox)
                    {
                        c.BackColor = bg; c.ForeColor = fg;
                    }
                    else if (IsKeepBackcolorControl(c))
                    {
                        // preserva badges, texto em verde CRT sobre fundo existente
                        c.ForeColor = fg;
                    }
                    else
                    {
                        c.BackColor = bg; c.ForeColor = fg;
                    }
                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }

        private static void ResetStyles(Control root)
        {
            void Walk(Control c)
            {
                // Reset genérico de cores para um baseline neutro
                if (IsKeepBackcolorControl(c))
                {
                    // preserva a cor de fundo (badge/status/cartões), mas normaliza a cor do texto
                    c.ForeColor = SystemColors.ControlText;
                }
                else if (IsAccentButton(c))
                {
                    // preserva o estilo/acento do botão
                }
                else if (c is TextBox or ComboBox or ListBox)
                {
                    c.BackColor = SystemColors.Window;
                    c.ForeColor = SystemColors.WindowText;
                }
                else if (c is DataGridView dgv)
                {
                    dgv.EnableHeadersVisualStyles = true;
                    dgv.BackgroundColor = SystemColors.Window;
                    dgv.GridColor = SystemColors.ControlDark;
                    dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle();
                    dgv.DefaultCellStyle = new DataGridViewCellStyle();
                }
                else if (c is Button btn2)
                {
                    btn2.FlatStyle = FlatStyle.Standard;
                    btn2.FlatAppearance.BorderSize = 1;
                    btn2.BackColor = SystemColors.Control;
                    btn2.ForeColor = SystemColors.ControlText;
                }
                else
                {
                    c.BackColor = SystemColors.Control;
                    c.ForeColor = SystemColors.ControlText;
                }
                foreach (Control child in c.Controls) Walk(child);
            }
            Walk(root);
        }

        public static void ApplyToAllOpenForms()
        {
            try
            {
                foreach (Form f in Application.OpenForms)
                {
                    try
                    {
                        ApplyAppIcon(f);
                        ApplyCurrentTheme(f);
                        // Opcional: chamar método ApplyTheme específico do formulário, se existir
                        try
                        {
                            var mi = f.GetType().GetMethod("ApplyTheme", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
                            mi?.Invoke(f, null);
                        }
                        catch { }
                        f.Invalidate(true); f.Refresh();
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static void StyleMenusForTheme(Form form, string theme)
        {
            Color bg, fg, sel, selBorder;
            switch (theme)
            {
                case "alto_contraste":
                case "terminal":
                case "escuro":
                    bg = theme == "alto_contraste" || theme == "terminal" ? Color.Black : Color.FromArgb(30, 30, 30);
                    fg = theme == "alto_contraste" ? Color.White : Color.FromArgb(235, 235, 235);
                    sel = theme == "alto_contraste" ? Color.Yellow : Color.FromArgb(55, 85, 120);
                    selBorder = theme == "alto_contraste" ? Color.White : Color.FromArgb(80, 110, 145);
                    break;
                case "claro":
                case "mica":
                case "classico":
                default:
                    bg = theme == "classico" ? SystemColors.Control : Color.FromArgb(246, 247, 250);
                    fg = Color.Black;
                    sel = Color.FromArgb(220, 232, 249);
                    selBorder = Color.FromArgb(180, 200, 230);
                    break;
            }

            void Walk(Control c)
            {
                if (c is MenuStrip ms)
                {
                    ms.BackColor = bg;
                    ms.ForeColor = fg;
                    ms.RenderMode = ToolStripRenderMode.Professional;
                    ms.Renderer = new ModernMenuRenderer(bg, fg, sel, selBorder);
                    try { ms.Font = (NormalizeThemeKey((MedControl.Database.GetConfig("theme") ?? "padrao").ToLowerInvariant()) == "padrao") ? SystemFonts.MenuFont : ms.Font; } catch { }
                    foreach (ToolStripItem it in ms.Items) it.ForeColor = fg;
                }
                foreach (Control child in c.Controls) Walk(child);
            }
            Walk(form);
        }

        private static void ApplyDefaultStyle(Control root)
        {
            try
            {
                var defaultFont = SystemFonts.DefaultFont;
                var menuFont = SystemFonts.MenuFont;
                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton)
                        {
                            if (c is MenuStrip)
                                c.Font = menuFont;
                            else
                                c.Font = defaultFont;
                        }
                    }
                    catch { }

                    // Cores já foram resetadas em ResetStyles; não forçar aqui para respeitar tags.

                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }

        public static void ApplyVisualStyleStateForCurrentTheme()
        {
            try
            {
                var theme = NormalizeThemeKey((MedControl.Database.GetConfig("theme") ?? "padrao").ToLowerInvariant());
                if (theme == "classico" || theme == "alto_contraste" || theme == "terminal")
                    Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled;
                else
                    Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.ClientAndNonClientAreasEnabled;
            }
            catch { }
        }

        private static string NormalizeThemeKey(string raw)
        {
            return raw switch
            {
                "marrom" => "padrao",
                "branco" => "claro",
                "preto" => "escuro",
                "azul" => "padrao",
                _ => raw
            };
        }

        private static void ApplyLightStyle(Control root)
        {
            try
            {
                var font = new Font("Segoe UI", 10f, FontStyle.Regular);
                var bg = Color.FromArgb(246, 247, 250); // claro suave
                var header = Color.FromArgb(235, 238, 243);
                var grid = Color.FromArgb(220, 224, 230);
                var text = Color.FromArgb(24, 24, 24);
                var sel = Color.FromArgb(220, 232, 249); // azul suave

                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton) c.Font = font;
                    }
                    catch { }
                    if (c is DataGridView dgv)
                    {
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = bg;
                        dgv.GridColor = grid;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = header;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = text;
                        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 10f);
                        dgv.DefaultCellStyle.BackColor = Color.White;
                        dgv.DefaultCellStyle.ForeColor = text;
                        dgv.DefaultCellStyle.SelectionBackColor = sel;
                        dgv.DefaultCellStyle.SelectionForeColor = text;
                    }
                    else if (c is TextBox or ComboBox or ListBox)
                    {
                        c.BackColor = Color.White; c.ForeColor = text;
                    }
                    else if (c is Button btn && (btn.Tag as string) != "accent")
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderSize = 1;
                        btn.FlatAppearance.BorderColor = Color.FromArgb(210, 214, 220);
                        btn.BackColor = Color.White;
                        btn.ForeColor = text;
                    }
                    else if (IsKeepBackcolorControl(c))
                    {
                        // preserva badges, texto escuro
                        c.ForeColor = text;
                    }
                    else
                    {
                        c.BackColor = bg; c.ForeColor = text;
                    }
                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }

        private static void ApplyDarkStyle(Control root)
        {
            try
            {
                var font = new Font("Segoe UI", 10f, FontStyle.Regular);
                var bg = Color.FromArgb(30, 30, 30);
                var header = Color.FromArgb(42, 42, 42);
                var grid = Color.FromArgb(58, 58, 58);
                var text = Color.FromArgb(235, 235, 235);
                var sel = Color.FromArgb(55, 85, 120); // seleção suave

                void Walk(Control c)
                {
                    try
                    {
                        var tag = c.Tag as string;
                        bool keepFont = tag != null && tag.IndexOf("keep-font", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isAccentButton = IsAccentButton(c);
                        if (!keepFont && !isAccentButton) c.Font = font;
                    }
                    catch { }
                    if (c is DataGridView dgv)
                    {
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = bg;
                        dgv.GridColor = grid;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = header;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = text;
                        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 10f);
                        dgv.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dgv.DefaultCellStyle.ForeColor = text;
                        dgv.DefaultCellStyle.SelectionBackColor = sel;
                        dgv.DefaultCellStyle.SelectionForeColor = Color.White;
                    }
                    else if (c is TextBox or ComboBox or ListBox)
                    {
                        c.BackColor = Color.FromArgb(37, 37, 38); c.ForeColor = text;
                    }
                    else if (c is Button btn && (btn.Tag as string) != "accent")
                    {
                        btn.FlatStyle = FlatStyle.Flat;
                        btn.FlatAppearance.BorderSize = 1;
                        btn.FlatAppearance.BorderColor = grid;
                        btn.BackColor = Color.FromArgb(45, 45, 48);
                        btn.ForeColor = text;
                    }
                    else if (IsKeepBackcolorControl(c))
                    {
                        // preserva badges, texto claro
                        c.ForeColor = text;
                    }
                    else
                    {
                        c.BackColor = bg; c.ForeColor = text;
                    }
                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
        }
    }
}
