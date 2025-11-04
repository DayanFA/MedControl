using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class EntregasForm : Form
    {
        private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
        private BindingSource _bs = new BindingSource();

        public EntregasForm()
        {
            Text = "Relação";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Width = 1000; Height = 600;

            // Top actions styled
            var panelTop = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(12, 8, 12, 8),
                BackColor = Color.White
            };
            Button MkPrimary(string text)
            {
                var b = new Button
                {
                    Text = text,
                    AutoSize = true,
                    Padding = new Padding(10, 6, 10, 6),
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point),
                    BackColor = Color.FromArgb(0, 123, 255),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                b.FlatAppearance.BorderSize = 0;
                var baseCol = b.BackColor;
                b.MouseEnter += (_, __) => b.BackColor = Darken(baseCol, 12);
                b.MouseLeave += (_, __) => b.BackColor = baseCol;
                b.MouseDown += (_, __) => b.BackColor = Darken(baseCol, 22);
                b.MouseUp +=   (_, __) => b.BackColor = Darken(baseCol, 12);
                return b;
            }
            Button MkNeutral(string text)
            {
                var b = new Button
                {
                    Text = text,
                    AutoSize = true,
                    Padding = new Padding(10, 6, 10, 6),
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                    BackColor = Color.WhiteSmoke,
                    ForeColor = Color.Black,
                    FlatStyle = FlatStyle.Flat
                };
                b.FlatAppearance.BorderSize = 1;
                b.FlatAppearance.BorderColor = Color.Silver;
                return b;
            }
            Button MkDanger(string text)
            {
                var b = new Button
                {
                    Text = text,
                    AutoSize = true,
                    Padding = new Padding(10, 6, 10, 6),
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point),
                    BackColor = Color.FromArgb(220, 53, 69),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                b.FlatAppearance.BorderSize = 0;
                return b;
            }

            var btnDevolvido = MkPrimary("Marcar como Devolvido");
            var btnEditar = MkNeutral("Editar");
            var btnExcluir = MkDanger("Excluir");
            var btnMaisInfo = MkNeutral("Mais Info");
            panelTop.Controls.AddRange(new Control[] { btnDevolvido, btnEditar, btnExcluir, btnMaisInfo });

            // Grid styling
            _grid.ReadOnly = true;
            _grid.MultiSelect = false;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.RowHeadersVisible = false;
            _grid.BackgroundColor = Color.White;
            _grid.BorderStyle = BorderStyle.None;
            _grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            _grid.GridColor = Color.Gainsboro;
            _grid.DefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.Gainsboro;
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            _grid.EnableHeadersVisualStyles = false;
            try { typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(_grid, true, null); } catch { }

            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Reserva.Chave), Width = 140 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Reserva.Aluno), Width = 180 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Reserva.Professor), Width = 180 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Data e Hora", DataPropertyName = nameof(Reserva.DataHora), Width = 180, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" } });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 110 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Item Atribuído", Width = 130 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Termo", DataPropertyName = nameof(Reserva.Termo), Width = 260 });
            _grid.CellFormatting += (_, e) =>
            {
                if (_grid.Columns[e.ColumnIndex].HeaderText == "Status" && _grid.Rows[e.RowIndex].DataBoundItem is Reserva r)
                {
                    e.Value = r.EmUso ? "Em Uso" : (r.Devolvido ? "Devolvido" : "Reservado");
                    if (e.CellStyle != null)
                    {
                        if (r.EmUso) e.CellStyle.BackColor = Color.MistyRose;
                        else if (r.Devolvido) e.CellStyle.BackColor = Color.Honeydew;
                        else e.CellStyle.BackColor = Color.LemonChiffon;
                    }
                }
                if (_grid.Columns[e.ColumnIndex].HeaderText == "Item Atribuído" && _grid.Rows[e.RowIndex].DataBoundItem is Reserva r2)
                {
                    e.Value = string.IsNullOrWhiteSpace(r2.Termo) ? "Não" : "Sim";
                }
            };

            Controls.Add(_grid);
            Controls.Add(panelTop);

            Load += (_, __) =>
            {
                try { BeginInvoke(new Action(() => { try { MedControl.UI.FluentEffects.ApplyWin11Mica(this); } catch { } })); } catch { }
                RefreshGrid();
            };
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
            private CheckBox _atualizarTermo = new CheckBox { Text = "Anexar Termo" };
            private TextBox _termo = new TextBox { ReadOnly = true };
            private Button _btnSelecionar = new Button { Text = "📎 Anexar PDF" };
            public Reserva Result { get; private set; }

            public EditarEntregaDialog(Reserva r)
            {
                Text = "Editar Entrega";
                StartPosition = FormStartPosition.CenterParent;
                BackColor = Color.White;
                AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
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

                var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(18, 18, 18, 8), AutoSize = true };
                table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                table.Controls.Add(new Label { Text = "Status:", AutoSize = true, Font = new Font("Segoe UI", 10F) }, 0, 0);
                _status.Items.AddRange(new object[] { "Em Uso", "Reservado" });
                _status.SelectedIndex = r.EmUso ? 0 : 1;
                _status.Font = new Font("Segoe UI", 10F);
                _status.Margin = new Padding(3, 8, 3, 8);
                _status.MinimumSize = new Size(260, 30);
                _status.Dock = DockStyle.Fill;
                table.Controls.Add(_status, 1, 0);

                _atualizarTermo.Margin = new Padding(3, 8, 3, 8);
                _atualizarTermo.Font = new Font("Segoe UI", 10F);
                table.Controls.Add(new Label { Text = string.Empty, AutoSize = true }, 0, 1);
                table.Controls.Add(_atualizarTermo, 1, 1);

                var panelTermo = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
                _termo.Width = 300; _termo.ReadOnly = true; _termo.Enabled = false; _termo.Text = "Selecione o PDF do termo..."; _termo.ForeColor = Color.Gray;
                _btnSelecionar.AutoSize = true; _btnSelecionar.Padding = new Padding(10, 6, 10, 6); _btnSelecionar.Font = new Font("Segoe UI", 10F);
                _btnSelecionar.FlatStyle = FlatStyle.Flat; _btnSelecionar.FlatAppearance.BorderSize = 1; _btnSelecionar.FlatAppearance.BorderColor = Color.Silver; _btnSelecionar.BackColor = Color.WhiteSmoke; _btnSelecionar.Enabled = false;
                panelTermo.Controls.Add(_termo);
                panelTermo.Controls.Add(_btnSelecionar);
                table.Controls.Add(new Label { Text = "Termo (PDF)", AutoSize = true, Font = new Font("Segoe UI", 10F) }, 0, 2);
                table.Controls.Add(panelTermo, 1, 2);

                var ok = new Button { Text = "Salvar", DialogResult = DialogResult.OK, AutoSize = true, Padding = new Padding(10, 6, 10, 6), Font = new Font("Segoe UI", 10F, FontStyle.Bold), BackColor = Color.FromArgb(0,123,255), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                ok.FlatAppearance.BorderSize = 0;
                var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel, AutoSize = true, Padding = new Padding(10, 6, 10, 6), Font = new Font("Segoe UI", 10F), BackColor = Color.WhiteSmoke, FlatStyle = FlatStyle.Flat };
                cancel.FlatAppearance.BorderSize = 1; cancel.FlatAppearance.BorderColor = Color.Silver;
                var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(8) };
                buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
                table.Controls.Add(buttons, 0, 3); table.SetColumnSpan(buttons, 2);

                Controls.Add(table);

                _atualizarTermo.CheckedChanged += (_, __) => { var en = _atualizarTermo.Checked; _btnSelecionar.Enabled = en; _termo.Enabled = en; if (!en) { _termo.Text = "Selecione o PDF do termo..."; _termo.ForeColor = Color.Gray; } };
                _btnSelecionar.Click += (_, __) =>
                {
                    using var ofd = new OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf" };
                    if (ofd.ShowDialog(this) == DialogResult.OK) { _termo.Text = ofd.FileName; _termo.ForeColor = Color.Black; }
                };
                ok.Click += (_, __) =>
                {
                    Result.EmUso = _status.SelectedIndex == 0;
                    if (_atualizarTermo.Checked) Result.Termo = _termo.Text;
                };
            }
        }

        private static Color Darken(Color color, int amount)
        {
            int r = Math.Max(0, color.R - amount);
            int g = Math.Max(0, color.G - amount);
            int b = Math.Max(0, color.B - amount);
            return Color.FromArgb(color.A, r, g, b);
        }
    }
}
