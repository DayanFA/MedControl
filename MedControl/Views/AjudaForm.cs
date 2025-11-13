using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class AjudaForm : Form
    {
        public AjudaForm()
        {
            Text = "Sobre";
            StartPosition = FormStartPosition.CenterParent;
            Width = 520;
            Height = 300;

            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new Padding(18, 18, 18, 12),
                AutoScroll = true // permite scroll se janela menor
            };
            // Cada linha autoajustada para garantir visibilidade do botão
            for (int i = 0; i < 7; i++) main.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var title = new Label
            {
                Text = "MedControl Chat",
                AutoSize = true,
                Font = new Font("Segoe UI", 18F, FontStyle.Bold, GraphicsUnit.Point),
                Margin = new Padding(0, 0, 0, 8),
                Tag = "keep-font"
            };

            var version = new Label
            {
                Text = $"Versão: {GetVersionString()}",
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point),
                Margin = new Padding(0, 0, 0, 12),
                Tag = "keep-font"
            };

            var desc = new Label
            {
                Text = "Aplicação de gerenciamento de chaves.",
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                Margin = new Padding(0, 0, 0, 12)
            };

            var lic = new Label
            {
                Text = "Licença: MIT",
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                Margin = new Padding(0, 0, 0, 12)
            };

            var linkIntro = new Label
            {
                Text = "Para atualizações e informações, clique aqui:",
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                Margin = new Padding(0, 0, 0, 4)
            };

            var link = new LinkLabel
            {
                Text = "https://github.com/DayanFA/MedControl",
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                LinkColor = Color.FromArgb(0, 123, 255),
                ActiveLinkColor = Color.FromArgb(0, 90, 200),
                Margin = new Padding(0, 0, 0, 12)
            };
            link.LinkClicked += (_, __) =>
            {
                try { Process.Start(new ProcessStartInfo { FileName = "https://github.com/DayanFA/MedControl", UseShellExecute = true }); } catch { }
            };

            var btnUpdate = new Button
            {
                Text = "Procurar atualização",
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 0)
            };
            btnUpdate.Click += async (_, __) =>
            {
                try
                {
                    btnUpdate.Enabled = false;
                    var prev = btnUpdate.Text;
                    btnUpdate.Text = "Procurando...";
                    await MedControl.UpdateService.CheckNowAsync(this);
                    btnUpdate.Text = prev;
                    btnUpdate.Enabled = true;
                }
                catch { btnUpdate.Enabled = true; }
            };

            main.Controls.Add(title, 0, 0);
            main.Controls.Add(version, 0, 1);
            main.Controls.Add(desc, 0, 2);
            main.Controls.Add(lic, 0, 3);
            main.Controls.Add(linkIntro, 0, 4);
            main.Controls.Add(link, 0, 5);
            main.Controls.Add(btnUpdate, 0, 6);
            Controls.Add(main);

            // Aplicar tema atual
            try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
        }

        private static string GetVersionString()
        {
            try
            {
                var v = Application.ProductVersion;
                // Normaliza para Major.Minor.Build
                // Remove qualquer sufixo (+hash ou -beta) e manter Major.Minor.Patch
                v = v.Trim();
                var metaIdx = v.IndexOfAny(new[] { '+', '-' });
                if (metaIdx > 0) v = v.Substring(0, metaIdx);
                var parts = v.Split('.');
                if (parts.Length >= 3)
                    return string.Join('.', parts[0], parts[1], parts[2]);
                // Garante sempre pelo menos 3 componentes
                if (parts.Length == 2) return parts[0] + "." + parts[1] + ".0";
                if (parts.Length == 1) return parts[0] + ".0.0";
                return "1.0.0";
            }
            catch { return "1.0.0"; }
        }
    }
}
