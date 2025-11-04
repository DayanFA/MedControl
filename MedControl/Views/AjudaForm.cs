using System.Diagnostics;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class AjudaForm : Form
    {
        public AjudaForm()
        {
            Text = "Sobre";
            Width = 500;
            Height = 420;
            var box = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Text = "Sistema de Gerenciamento de Chaves\r\n\r\nVersão 1.0\r\nLicença: a definir\r\n\r\nContato:\r\nAndré Ferreira - ferreira.andre@sou.ufac.br\r\nDayan FA - contatodayanfa@gmail.com\r\n"
            };
            Controls.Add(box);

            // Aplicar tema atual
            try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
        }
    }
}
