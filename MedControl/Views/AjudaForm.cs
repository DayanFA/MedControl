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
                RowCount = 6,
                Padding = new Padding(18, 18, 18, 12),
                AutoSize = true,
            };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

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
                Text = $"v {GetVersionString()}",
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                Margin = new Padding(0, 0, 0, 12)
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

            main.Controls.Add(title);
            main.Controls.Add(version);
            main.Controls.Add(desc);
            main.Controls.Add(lic);
            main.Controls.Add(linkIntro);
            main.Controls.Add(link);
            main.Controls.Add(btnUpdate);
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
                var parts = v.Split('.');
                if (parts.Length >= 3)
                    return string.Join('.', parts[0], parts[1], parts[2]);
                return v;
            }
            catch { return "1.0.0"; }
        }
    }
}
