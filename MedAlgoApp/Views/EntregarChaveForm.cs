using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class EntregarChaveForm : Form
    {
        private ComboBox _chave = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private TextBox _aluno = new TextBox();
        private TextBox _prof = new TextBox();
        private CheckBox _itemEmprestado = new CheckBox { Text = "Algum item está sendo emprestado" };
        private TextBox _termo = new TextBox { ReadOnly = true };
        private Button _btnTermo = new Button { Text = "Importar Termo" };
    private Button _btn = new Button { Text = "Entregar" };
    private Button _verEntregas = new Button { Text = "Ver Entregas" };

        public EntregarChaveForm()
        {
            Text = "Entregar Chave";
            Width = 500;
            Height = 320;

            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            table.Controls.Add(new Label { Text = "Chave:" }, 0, 0);
            table.Controls.Add(_chave, 1, 0);
            table.Controls.Add(new Label { Text = "Aluno:" }, 0, 1);
            table.Controls.Add(_aluno, 1, 1);
            table.Controls.Add(new Label { Text = "Professor:" }, 0, 2);
            table.Controls.Add(_prof, 1, 2);

            table.Controls.Add(_itemEmprestado, 1, 3);
            table.Controls.Add(new Label { Text = "Termo de Compromisso (PDF):" }, 0, 4);
            var termoPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill };
            termoPanel.Controls.Add(_termo);
            termoPanel.Controls.Add(_btnTermo);
            _termo.Width = 250;
            table.Controls.Add(termoPanel, 1, 4);
            var panelButtons = new FlowLayoutPanel { Dock = DockStyle.Fill };
            panelButtons.Controls.Add(_btn);
            panelButtons.Controls.Add(_verEntregas);
            table.Controls.Add(panelButtons, 1, 5);

            Controls.Add(table);

            Load += (_, __) =>
            {
                _btnTermo.Enabled = false;
                _chave.Items.Clear();
                foreach (var c in Database.GetChaves()) _chave.Items.Add(c.Nome);
                if (_chave.Items.Count > 0) _chave.SelectedIndex = 0;
            };

            _itemEmprestado.CheckedChanged += (_, __) =>
            {
                _btnTermo.Enabled = _itemEmprestado.Checked;
                if (!_itemEmprestado.Checked) _termo.Text = string.Empty;
            };

            _btnTermo.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf" };
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    _termo.Text = ofd.FileName;
                }
            };

            _btn.Click += (_, __) =>
            {
                if (_chave.SelectedItem is null || (string.IsNullOrWhiteSpace(_aluno.Text) && string.IsNullOrWhiteSpace(_prof.Text)))
                {
                    MessageBox.Show("Selecione uma chave e informe aluno ou professor.");
                    return;
                }
                var now = DateTime.Now;
                var r = new Reserva
                {
                    Chave = _chave.SelectedItem!.ToString()!,
                    Aluno = _aluno.Text.Trim(),
                    Professor = _prof.Text.Trim(),
                    DataHora = now,
                    EmUso = true,
                    Termo = _termo.Text,
                    Devolvido = false,
                    DataDevolucao = null
                };
                Database.InsertReserva(r);
                MessageBox.Show($"Chave {r.Chave} entregue por {(string.IsNullOrWhiteSpace(r.Aluno) ? r.Professor : r.Aluno)} em {r.DataHora:dd/MM/yyyy HH:mm:ss}");
                Close();
            };

            _verEntregas.Click += (_, __) => new EntregasForm().ShowDialog(this);
        }
    }
}
