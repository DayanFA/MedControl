using System;
using System.Drawing;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class SyncProgressForm : Form
    {
        private readonly Label _status;
        private readonly ProgressBar _bar;

        public SyncProgressForm()
        {
            Text = "Sincronizando...";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Size = new Size(460, 150);
            TopMost = true;
            ShowInTaskbar = false;

            _status = new Label
            {
                Text = "Sincronizando 0%",
                Dock = DockStyle.Top,
                Height = 36,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10, FontStyle.Regular)
            };
            _bar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Style = ProgressBarStyle.Continuous,
                Height = 20,
                Margin = new Padding(12)
            };
            // Dica
            var tip = new Label
            {
                Text = "Os dados serão salvos localmente para uso offline.",
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

        public void SetProgress(int percent, string text)
        {
            try
            {
                if (percent < 0) percent = 0; if (percent > 100) percent = 100;
                _bar.Value = percent;
                _status.Text = string.IsNullOrWhiteSpace(text) ? $"Sincronizando {percent}%" : text;
            }
            catch { }
        }
    }
}
