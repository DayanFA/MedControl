using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class EntregasForm : Form
    {
        private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
        private BindingSource _bs = new BindingSource();

        public EntregasForm()
        {
            Text = "Entregas";
            Width = 1000; Height = 600;

            var panelTop = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40 };
            var btnDevolvido = new Button { Text = "Marcar como Devolvido" };
            var btnEditar = new Button { Text = "Editar" };
            var btnExcluir = new Button { Text = "Excluir" };
            var btnMaisInfo = new Button { Text = "Mais Info" };
            panelTop.Controls.AddRange(new Control[] { btnDevolvido, btnEditar, btnExcluir, btnMaisInfo });

            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Reserva.Chave), Width = 120 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Reserva.Aluno), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Reserva.Professor), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Data e Hora", DataPropertyName = nameof(Reserva.DataHora), Width = 160 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 100 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Item Atribuído", Width = 120 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Termo de Compromisso", DataPropertyName = nameof(Reserva.Termo), Width = 220 });
            _grid.CellFormatting += (_, e) =>
            {
                if (_grid.Columns[e.ColumnIndex].HeaderText == "Status" && _grid.Rows[e.RowIndex].DataBoundItem is Reserva r)
                {
                    e.Value = r.EmUso ? "Em Uso" : "Reservado";
                }
                if (_grid.Columns[e.ColumnIndex].HeaderText == "Item Atribuído" && _grid.Rows[e.RowIndex].DataBoundItem is Reserva r2)
                {
                    e.Value = string.IsNullOrWhiteSpace(r2.Termo) ? "Não" : "Sim";
                }
            };

            Controls.Add(_grid);
            Controls.Add(panelTop);

            Load += (_, __) => RefreshGrid();
            _grid.CellDoubleClick += (_, e) => OpenTermo();
            btnDevolvido.Click += (_, __) => MarcarDevolvido();
            btnEditar.Click += (_, __) => EditarEntrega();
            btnExcluir.Click += (_, __) => ExcluirEntrega();
            btnMaisInfo.Click += (_, __) => MaisInfo();
        }

        private void RefreshGrid()
        {
            _bs.DataSource = Database.GetReservas();
            _grid.DataSource = _bs;
        }

        private Reserva? Current()
        {
            return _bs.Current as Reserva;
        }

        private void OpenTermo()
        {
            if (Current() is { } r && !string.IsNullOrWhiteSpace(r.Termo) && File.Exists(r.Termo))
            {
                try { Process.Start(new ProcessStartInfo { FileName = r.Termo, UseShellExecute = true }); }
                catch { MessageBox.Show("Não foi possível abrir o PDF."); }
            }
        }

        private void MarcarDevolvido()
        {
            var r = Current();
            if (r == null)
            {
                MessageBox.Show("Selecione uma entrega.");
                return;
            }
            r.Devolvido = true;
            r.EmUso = false;
            r.DataDevolucao = DateTime.Now;
            var tempo = r.DataDevolucao.Value - r.DataHora;
            var tempoStr = tempo.ToString().Split('.')[0];
            // inserir no relatorio
            Database.InsertRelatorio(new Relatorio
            {
                Chave = r.Chave,
                Aluno = r.Aluno,
                Professor = r.Professor,
                DataHora = r.DataHora,
                DataDevolucao = r.DataDevolucao,
                TempoComChave = tempoStr,
                Termo = r.Termo
            });
            // remover da reservas
            Database.DeleteReserva(r.Chave, r.Aluno, r.Professor, r.DataHora);
            RefreshGrid();
        }

        private void EditarEntrega()
        {
            var r = Current();
            if (r == null)
            {
                MessageBox.Show("Selecione uma entrega.");
                return;
            }
            using var dlg = new EditarEntregaDialog(r);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                Database.UpdateReserva(dlg.Result);
                RefreshGrid();
            }
        }

        private void ExcluirEntrega()
        {
            var r = Current();
            if (r == null) { MessageBox.Show("Selecione uma entrega."); return; }
            if (MessageBox.Show("Excluir esta entrega?", "Excluir", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Database.DeleteReserva(r.Chave, r.Aluno, r.Professor, r.DataHora);
                RefreshGrid();
            }
        }

        private void MaisInfo()
        {
            var r = Current();
            if (r == null) { MessageBox.Show("Selecione uma entrega."); return; }
            MessageBox.Show($"Chave: {r.Chave}\nAluno: {r.Aluno}\nProfessor: {r.Professor}\nData e Hora: {r.DataHora:dd/MM/yyyy HH:mm:ss}", "Informações");
        }

        private class EditarEntregaDialog : Form
        {
            private ComboBox _status = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            private CheckBox _atualizarTermo = new CheckBox { Text = "Atualizar novo termo de compromisso" };
            private TextBox _termo = new TextBox { ReadOnly = true };
            private Button _btnSelecionar = new Button { Text = "Importar Termo" };
            public Reserva Result { get; private set; }

            public EditarEntregaDialog(Reserva r)
            {
                Text = "Editar Entrega";
                Width = 420; Height = 220;
                Result = new Reserva
                {
                    Chave = r.Chave,
                    Aluno = r.Aluno,
                    Professor = r.Professor,
                    DataHora = r.DataHora,
                    EmUso = r.EmUso,
                    Termo = r.Termo,
                    Devolvido = r.Devolvido,
                    DataDevolucao = r.DataDevolucao
                };

                var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10) };
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

                table.Controls.Add(new Label { Text = "Status:" }, 0, 0);
                _status.Items.AddRange(new object[] { "Em Uso", "Reservado" });
                _status.SelectedIndex = r.EmUso ? 0 : 1;
                table.Controls.Add(_status, 1, 0);

                table.Controls.Add(_atualizarTermo, 1, 1);
                var panelTermo = new FlowLayoutPanel { Dock = DockStyle.Fill };
                panelTermo.Controls.Add(_termo); panelTermo.Controls.Add(_btnSelecionar); _termo.Width = 240;
                _btnSelecionar.Enabled = false;
                table.Controls.Add(panelTermo, 1, 2);

                var ok = new Button { Text = "Salvar", DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel };
                var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
                buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
                table.Controls.Add(buttons, 0, 3); table.SetColumnSpan(buttons, 2);

                Controls.Add(table);

                _atualizarTermo.CheckedChanged += (_, __) => { _btnSelecionar.Enabled = _atualizarTermo.Checked; if (!_atualizarTermo.Checked) _termo.Text = string.Empty; };
                _btnSelecionar.Click += (_, __) =>
                {
                    using var ofd = new OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf" };
                    if (ofd.ShowDialog(this) == DialogResult.OK) _termo.Text = ofd.FileName;
                };
                ok.Click += (_, __) =>
                {
                    Result.EmUso = _status.SelectedIndex == 0;
                    if (_atualizarTermo.Checked) Result.Termo = _termo.Text;
                };
            }
        }
    }
}
