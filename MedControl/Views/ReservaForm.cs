using System;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class ReservaForm : Form
    {
        private ComboBox _chave = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
    private ComboBox _aluno = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown }; // permite digitar
    private ComboBox _prof = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
        private DateTimePicker _dataHora = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm:ss", ShowUpDown = true };
        private Button _btn = new Button { Text = "Reservar" };

        public ReservaForm()
        {
            Text = "Fazer Reserva";
            Width = 420;
            Height = 260;
            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
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
            table.Controls.Add(new Label { Text = "Data e Hora:" }, 0, 3);
            table.Controls.Add(_dataHora, 1, 3);
            table.Controls.Add(_btn, 1, 4);

            Controls.Add(table);

            Load += (_, __) =>
            {
                _chave.Items.Clear();
                foreach (var c in Database.GetChaves()) _chave.Items.Add(c.Nome);
                if (_chave.Items.Count > 0) _chave.SelectedIndex = 0;

                var alunosPath = Database.GetConfig("caminho_alunos");
                var profsPath = Database.GetConfig("caminho_professores");
                _aluno.Items.Clear();
                _prof.Items.Clear();
                foreach (var a in ExcelHelper.LoadFirstColumn(alunosPath ?? string.Empty)) _aluno.Items.Add(a);
                foreach (var p in ExcelHelper.LoadFirstColumn(profsPath ?? string.Empty)) _prof.Items.Add(p);
            };

            _btn.Click += (_, __) =>
            {
                if (_chave.SelectedItem is null || (string.IsNullOrWhiteSpace(_aluno.Text) && string.IsNullOrWhiteSpace(_prof.Text)))
                {
                    MessageBox.Show("Selecione uma chave e informe aluno ou professor.");
                    return;
                }
                var r = new Reserva
                {
                    Chave = _chave.SelectedItem!.ToString()!,
                    Aluno = _aluno.Text.Trim(),
                    Professor = _prof.Text.Trim(),
                    DataHora = _dataHora.Value,
                    EmUso = false,
                    Termo = string.Empty,
                    Devolvido = false,
                    DataDevolucao = null
                };
                Database.InsertReserva(r);
                MessageBox.Show($"Chave {r.Chave} reservada por {(string.IsNullOrWhiteSpace(r.Aluno) ? r.Professor : r.Aluno)} em {r.DataHora:dd/MM/yyyy HH:mm:ss}");
                Close();
            };
        }
    }
}
