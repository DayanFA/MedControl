using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class PreloadForm : Form
    {
        private Label _status;
        private ProgressBar _bar;

        public PreloadForm()
        {
            Text = "Preparando...";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Size = new Size(420, 140);
            TopMost = true;

            _status = new Label
            {
                Text = "Baixando dados do grupo...",
                Dock = DockStyle.Top,
                Height = 36,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10, FontStyle.Regular)
            };
            _bar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Height = 20,
                Margin = new Padding(12)
            };
            var tip = new Label
            {
                Text = "Dica: funciona offline – dados são salvos localmente.",
                Dock = DockStyle.Bottom,
                Height = 24,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8, FontStyle.Italic)
            };

            Padding = new Padding(16);
            Controls.Add(tip);
            Controls.Add(_bar);
            Controls.Add(_status);
        }

        public void SetStatus(string text)
        {
            try { _status.Text = text; } catch { }
        }
    }
}
