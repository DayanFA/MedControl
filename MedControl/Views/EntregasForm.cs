using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Data;

namespace MedControl.Views
{
    public class EntregasForm : Form
    {
        private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
        private BindingSource _bs = new BindingSource();
        private TextBox _search = new TextBox();
        private System.Collections.Generic.List<Reserva> _allReservas = new System.Collections.Generic.List<Reserva>();
        private string _searchText = string.Empty;
        private string? _sortKey = null; // Chave, Aluno, Professor, DataHora, Status, Item, Termo
        private bool _sortAsc = true;

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
            // search box com placeholder e emoji
            _search.Width = 260; _search.Font = new Font("Segoe UI", 10F); _search.Margin = new Padding(18, 6, 3, 3);
            try { _search.PlaceholderText = "🔎 Pesquisar..."; } catch { }
            _search.TextChanged += (_, __) => { _searchText = _search.Text?.Trim() ?? string.Empty; ApplyFiltersAndBind(); };
            panelTop.Controls.AddRange(new Control[] { btnDevolvido, btnEditar, btnExcluir, btnMaisInfo, _search });

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

            var cChave = new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Reserva.Chave), Width = 140, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cAluno = new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Reserva.Aluno), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cProf = new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Reserva.Professor), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cData = new DataGridViewTextBoxColumn { HeaderText = "Data e Hora", DataPropertyName = nameof(Reserva.DataHora), Width = 180, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" }, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cStatus = new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 110, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cItem = new DataGridViewTextBoxColumn { HeaderText = "Item Atribuído", Width = 130, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cTermo = new DataGridViewTextBoxColumn { HeaderText = "Termo", DataPropertyName = nameof(Reserva.Termo), Width = 260, SortMode = DataGridViewColumnSortMode.Programmatic };
            _grid.Columns.AddRange(new DataGridViewColumn[] { cChave, cAluno, cProf, cData, cStatus, cItem, cTermo });
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
            _grid.ColumnHeaderMouseClick += (_, e) => OnGridHeaderClick(e.ColumnIndex);
            btnDevolvido.Click += (_, __) => MarcarDevolvido();
            btnEditar.Click += (_, __) => EditarEntrega();
            btnExcluir.Click += (_, __) => ExcluirEntrega();
            btnMaisInfo.Click += (_, __) => MaisInfo();
        }

        private void RefreshGrid()
        {
            _allReservas = Database.GetReservas();
            ApplyFiltersAndBind();
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

        private void OnGridHeaderClick(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= _grid.Columns.Count) return;
            var col = _grid.Columns[columnIndex];
            var key = GetSortKey(col);
            if (key == null) return;
            if (_sortKey == key) _sortAsc = !_sortAsc; else { _sortKey = key; _sortAsc = true; }
            ApplyFiltersAndBind();
            UpdateSortGlyphs(col);
        }

        private void UpdateSortGlyphs(DataGridViewColumn sortedCol)
        {
            foreach (DataGridViewColumn c in _grid.Columns) c.HeaderCell.SortGlyphDirection = SortOrder.None;
            sortedCol.HeaderCell.SortGlyphDirection = _sortAsc ? SortOrder.Ascending : SortOrder.Descending;
        }

        private string? GetSortKey(DataGridViewColumn col)
        {
            var header = col.HeaderText;
            switch (header)
            {
                case "Chave": return nameof(Reserva.Chave);
                case "Aluno": return nameof(Reserva.Aluno);
                case "Professor": return nameof(Reserva.Professor);
                case "Data e Hora": return nameof(Reserva.DataHora);
                case "Status": return "Status"; // calculado
                case "Item Atribuído": return "Item"; // calculado
                case "Termo": return nameof(Reserva.Termo);
            }
            return null;
        }

        private void ApplyFiltersAndBind()
        {
            var q = _allReservas.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(_searchText))
            {
                var s = _searchText.ToLowerInvariant();
                q = q.Where(r =>
                    (r.Chave ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.Aluno ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.Professor ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.Termo ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    r.DataHora.ToString("dd/MM/yyyy HH:mm:ss").Contains(s)
                );
            }

            if (!string.IsNullOrEmpty(_sortKey))
            {
                Func<Reserva, object?> selector = _sortKey switch
                {
                    nameof(Reserva.Chave) => r => r.Chave,
                    nameof(Reserva.Aluno) => r => r.Aluno,
                    nameof(Reserva.Professor) => r => r.Professor,
                    nameof(Reserva.DataHora) => r => r.DataHora,
                    nameof(Reserva.Termo) => r => r.Termo,
                    "Status" => r => r.EmUso ? "Em Uso" : (r.Devolvido ? "Devolvido" : "Reservado"),
                    "Item" => r => string.IsNullOrWhiteSpace(r.Termo) ? "Não" : "Sim",
                    _ => r => r.Chave
                };
                q = _sortAsc ? q.OrderBy(selector).ThenBy(r => r.DataHora) : q.OrderByDescending(selector).ThenByDescending(r => r.DataHora);
            }
            var list = q.ToList();
            _bs.DataSource = list;
            _grid.DataSource = _bs;
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
            using var dlg = new MaisInfoForm(r.Aluno, r.Professor);
            dlg.ShowDialog(this);
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

                // Se já existe termo, marcar e destacar automaticamente
                if (!string.IsNullOrWhiteSpace(r.Termo))
                {
                    _atualizarTermo.Checked = true; // dispara o CheckedChanged e habilita os controles
                    _termo.Text = r.Termo;
                    _termo.ForeColor = Color.Black;
                    _termo.BackColor = Color.LightYellow; // leve destaque visual
                }

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

        private class MaisInfoForm : Form
        {
            private DataGridView _gridAluno = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = true, ReadOnly = true, MultiSelect = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
            private DataGridView _gridAtivos = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, ReadOnly = true, MultiSelect = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
            private DataGridView _gridHistorico = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, ReadOnly = true, MultiSelect = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
            private TextBox _search = new TextBox();
            private string _searchText = string.Empty;
            private System.Collections.Generic.List<Reserva> _baseAtivos = new System.Collections.Generic.List<Reserva>();
            private System.Collections.Generic.List<Relatorio> _baseHistorico = new System.Collections.Generic.List<Relatorio>();
            private DataView? _viewAluno;
            private string? _sortAtivosKey; private bool _sortAtivosAsc = true;
            private string? _sortHistKey; private bool _sortHistAsc = true;

            public MaisInfoForm(string? aluno, string? professor)
            {
                Text = "Mais Informações";
                StartPosition = FormStartPosition.CenterParent;
                BackColor = Color.White;
                Width = 1000; Height = 640;

                var header = new Label
                {
                    Text = $"Aluno: {aluno ?? "-"}",
                    Dock = DockStyle.Top,
                    AutoSize = false,
                    Height = 36,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(12, 8, 12, 8),
                    Font = new Font("Segoe UI", 11F, FontStyle.Bold)
                };

                // tabs + search
                var tabs = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 10F) };
                var tabInfoAluno = new TabPage("Info do Aluno");
                var tabAtivos = new TabPage("Em andamento");
                var tabHistorico = new TabPage("Histórico");
                tabs.TabPages.Add(tabInfoAluno);
                tabs.TabPages.Add(tabAtivos);
                tabs.TabPages.Add(tabHistorico);

                var searchPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12, 8, 12, 0), BackColor = Color.White };
                _search.Width = 260; _search.Font = new Font("Segoe UI", 10F); _search.Margin = new Padding(0, 2, 3, 6);
                try { _search.PlaceholderText = "🔎 Pesquisar..."; } catch { }
                _search.TextChanged += (_, __) => { _searchText = _search.Text?.Trim() ?? string.Empty; ApplySearchAndSort(tabs.SelectedTab); };
                searchPanel.Controls.Add(_search);

                StyleGrid(_gridAluno);
                StyleGrid(_gridAtivos);
                StyleGrid(_gridHistorico);
                SetupColumnsAtivos(_gridAtivos);
                SetupColumnsHistorico(_gridHistorico);

                tabInfoAluno.Controls.Add(_gridAluno);
                tabAtivos.Controls.Add(_gridAtivos);
                tabHistorico.Controls.Add(_gridHistorico);

                Controls.Add(tabs);
                Controls.Add(searchPanel);
                Controls.Add(header);

                Load += (_, __) =>
                {
                    // Info do Aluno: mostrar apenas a linha do cadastro do aluno
                    try
                    {
                        var alunosDt = Database.GetAlunosAsDataTable();
                        DataTable filtered = alunosDt.Clone();
                        if (!string.IsNullOrWhiteSpace(aluno))
                        {
                            string? alunosNomeCol = null;
                            foreach (DataColumn col in alunosDt.Columns)
                            {
                                var name = col.ColumnName;
                                if (string.Equals(name, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                                if (alunosNomeCol == null || string.Equals(name, "Nome", StringComparison.OrdinalIgnoreCase)) alunosNomeCol = name;
                            }
                            if (alunosNomeCol != null)
                            {
                                foreach (DataRow row in alunosDt.Rows)
                                {
                                    var v = Convert.ToString(row[alunosNomeCol]) ?? string.Empty;
                                    if (!string.IsNullOrWhiteSpace(v) && string.Equals(v, aluno, StringComparison.OrdinalIgnoreCase))
                                    {
                                        filtered.ImportRow(row);
                                    }
                                }
                            }
                        }
                        _viewAluno = new DataView(filtered);
                        _gridAluno.DataSource = _viewAluno;
                    }
                    catch { _gridAluno.DataSource = null; }

                    // filtro por aluno/professor (OR) nas listas
                    _baseAtivos = Database.GetReservas()
                        .Where(r => MatchPessoa(r.Aluno, r.Professor, aluno, professor))
                        .OrderByDescending(r => r.DataHora)
                        .ToList();
                    _baseHistorico = Database.GetRelatorios()
                        .Where(r => MatchPessoa(r.Aluno, r.Professor, aluno, professor))
                        .OrderByDescending(r => r.DataHora)
                        .ToList();

                    ApplySearchAndSort(tabAtivos);
                };

                _gridAtivos.CellDoubleClick += (_, e) => OpenTermoFromGrid(_gridAtivos, e.RowIndex);
                _gridHistorico.CellDoubleClick += (_, e) => OpenTermoFromGrid(_gridHistorico, e.RowIndex);
                _gridAtivos.ColumnHeaderMouseClick += (_, e) => { OnHeaderClickAtivos(e.ColumnIndex); ApplySearchAndSort(tabAtivos); };
                _gridHistorico.ColumnHeaderMouseClick += (_, e) => { OnHeaderClickHistorico(e.ColumnIndex); ApplySearchAndSort(tabHistorico); };
                _gridAluno.ColumnHeaderMouseClick += (_, e) => { if (_viewAluno != null && e.ColumnIndex >= 0) { var col = _gridAluno.Columns[e.ColumnIndex]; ToggleSortAluno(col); } };
                tabs.SelectedIndexChanged += (_, __) => ApplySearchAndSort(tabs.SelectedTab);
            }

            private static bool MatchPessoa(string? alunoItem, string? profItem, string? alunoFiltro, string? profFiltro)
            {
                bool matchAluno = !string.IsNullOrWhiteSpace(alunoFiltro) && string.Equals(alunoItem ?? string.Empty, alunoFiltro, StringComparison.OrdinalIgnoreCase);
                bool matchProf = !string.IsNullOrWhiteSpace(profFiltro) && string.Equals(profItem ?? string.Empty, profFiltro, StringComparison.OrdinalIgnoreCase);
                // Se nenhum filtro veio, não filtra (mostra tudo)
                if (string.IsNullOrWhiteSpace(alunoFiltro) && string.IsNullOrWhiteSpace(profFiltro)) return true;
                return matchAluno || matchProf;
            }

            private void OpenTermoFromGrid(DataGridView grid, int rowIndex)
            {
                if (rowIndex < 0) return;
                var item = grid.Rows[rowIndex].DataBoundItem;
                string? termo = null;
                switch (item)
                {
                    case Reserva rr:
                        termo = rr.Termo; break;
                    case Relatorio rl:
                        termo = rl.Termo; break;
                }
                if (!string.IsNullOrWhiteSpace(termo) && File.Exists(termo))
                {
                    try { Process.Start(new ProcessStartInfo { FileName = termo, UseShellExecute = true }); }
                    catch { MessageBox.Show("Não foi possível abrir o PDF."); }
                }
            }

            private static void StyleGrid(DataGridView grid)
            {
                grid.BackgroundColor = Color.White;
                grid.BorderStyle = BorderStyle.None;
                grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
                grid.GridColor = Color.Gainsboro;
                grid.DefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
                grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
                grid.ColumnHeadersDefaultCellStyle.BackColor = Color.Gainsboro;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
                grid.EnableHeadersVisualStyles = false;
                try { typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(grid, true, null); } catch { }
            }

            private static void SetupColumnsAtivos(DataGridView grid)
            {
                grid.Columns.Clear();
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Reserva.Chave), Width = 140, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Reserva.Aluno), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Reserva.Professor), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Data/Hora", DataPropertyName = nameof(Reserva.DataHora), Width = 170, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" }, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Status", Width = 110, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Termo", DataPropertyName = nameof(Reserva.Termo), Width = 260, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.CellFormatting += (_, e) =>
                {
                    if (grid.Columns[e.ColumnIndex].HeaderText == "Status" && grid.Rows[e.RowIndex].DataBoundItem is Reserva r)
                    {
                        e.Value = r.EmUso ? "Em Uso" : (r.Devolvido ? "Devolvido" : "Reservado");
                    }
                };
            }

            private static void SetupColumnsHistorico(DataGridView grid)
            {
                grid.Columns.Clear();
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Relatorio.Chave), Width = 140, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Relatorio.Aluno), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Relatorio.Professor), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Retirada", DataPropertyName = nameof(Relatorio.DataHora), Width = 170, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" }, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Devolução", DataPropertyName = nameof(Relatorio.DataDevolucao), Width = 170, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" }, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tempo", DataPropertyName = nameof(Relatorio.TempoComChave), Width = 120, SortMode = DataGridViewColumnSortMode.Programmatic });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Termo", DataPropertyName = nameof(Relatorio.Termo), Width = 260, SortMode = DataGridViewColumnSortMode.Programmatic });
            }

            private void ApplySearchAndSort(TabPage? currentTab)
            {
                // Info do Aluno
                if (_viewAluno != null && _viewAluno.Table != null)
                {
                    _viewAluno.RowFilter = BuildRowFilter(_viewAluno.Table, _searchText);
                }

                // Em andamento
                var qa = _baseAtivos.AsEnumerable();
                if (!string.IsNullOrWhiteSpace(_searchText))
                {
                    var s = _searchText.ToLowerInvariant();
                    qa = qa.Where(r =>
                        (r.Chave ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.Aluno ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.Professor ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.Termo ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        r.DataHora.ToString("dd/MM/yyyy HH:mm:ss").Contains(s)
                    );
                }
                if (!string.IsNullOrEmpty(_sortAtivosKey))
                {
                    Func<Reserva, object?> selector = _sortAtivosKey switch
                    {
                        nameof(Reserva.Chave) => r => r.Chave,
                        nameof(Reserva.Aluno) => r => r.Aluno,
                        nameof(Reserva.Professor) => r => r.Professor,
                        nameof(Reserva.DataHora) => r => r.DataHora,
                        nameof(Reserva.Termo) => r => r.Termo,
                        "Status" => r => r.EmUso ? "Em Uso" : (r.Devolvido ? "Devolvido" : "Reservado"),
                        _ => r => r.Chave
                    };
                    qa = _sortAtivosAsc ? qa.OrderBy(selector) : qa.OrderByDescending(selector);
                }
                _gridAtivos.DataSource = qa.ToList();

                // Histórico
                var qh = _baseHistorico.AsEnumerable();
                if (!string.IsNullOrWhiteSpace(_searchText))
                {
                    var s = _searchText.ToLowerInvariant();
                    qh = qh.Where(r =>
                        (r.Chave ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.Aluno ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.Professor ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.Termo ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        (r.TempoComChave ?? string.Empty).ToLowerInvariant().Contains(s) ||
                        r.DataHora.ToString("dd/MM/yyyy HH:mm:ss").Contains(s) ||
                        (r.DataDevolucao?.ToString("dd/MM/yyyy HH:mm:ss") ?? string.Empty).Contains(s)
                    );
                }
                if (!string.IsNullOrEmpty(_sortHistKey))
                {
                    Func<Relatorio, object?> selector = _sortHistKey switch
                    {
                        nameof(Relatorio.Chave) => r => r.Chave,
                        nameof(Relatorio.Aluno) => r => r.Aluno,
                        nameof(Relatorio.Professor) => r => r.Professor,
                        nameof(Relatorio.DataHora) => r => r.DataHora,
                        nameof(Relatorio.DataDevolucao) => r => r.DataDevolucao,
                        nameof(Relatorio.TempoComChave) => r => r.TempoComChave,
                        nameof(Relatorio.Termo) => r => r.Termo,
                        _ => r => r.Chave
                    };
                    qh = _sortHistAsc ? qh.OrderBy(selector) : qh.OrderByDescending(selector);
                }
                _gridHistorico.DataSource = qh.ToList();
            }

            private static string BuildRowFilter(DataTable table, string search)
            {
                if (string.IsNullOrWhiteSpace(search)) return string.Empty;
                string s = search.Replace("'", "''");
                var parts = table.Columns.Cast<DataColumn>()
                    .Select(c => $"CONVERT([{c.ColumnName}], System.String) LIKE '%{s}%'");
                return string.Join(" OR ", parts);
            }

            private void OnHeaderClickAtivos(int colIndex)
            {
                if (colIndex < 0) return;
                var col = _gridAtivos.Columns[colIndex];
                string? key = col.HeaderText switch
                {
                    "Chave" => nameof(Reserva.Chave),
                    "Aluno" => nameof(Reserva.Aluno),
                    "Professor" => nameof(Reserva.Professor),
                    "Data/Hora" => nameof(Reserva.DataHora),
                    "Status" => "Status",
                    "Termo" => nameof(Reserva.Termo),
                    _ => null
                };
                if (key == null) return;
                if (_sortAtivosKey == key) _sortAtivosAsc = !_sortAtivosAsc; else { _sortAtivosKey = key; _sortAtivosAsc = true; }
                foreach (DataGridViewColumn c in _gridAtivos.Columns) c.HeaderCell.SortGlyphDirection = SortOrder.None;
                col.HeaderCell.SortGlyphDirection = _sortAtivosAsc ? SortOrder.Ascending : SortOrder.Descending;
            }

            private void OnHeaderClickHistorico(int colIndex)
            {
                if (colIndex < 0) return;
                var col = _gridHistorico.Columns[colIndex];
                string? key = col.HeaderText switch
                {
                    "Chave" => nameof(Relatorio.Chave),
                    "Aluno" => nameof(Relatorio.Aluno),
                    "Professor" => nameof(Relatorio.Professor),
                    "Retirada" => nameof(Relatorio.DataHora),
                    "Devolução" => nameof(Relatorio.DataDevolucao),
                    "Tempo" => nameof(Relatorio.TempoComChave),
                    "Termo" => nameof(Relatorio.Termo),
                    _ => null
                };
                if (key == null) return;
                if (_sortHistKey == key) _sortHistAsc = !_sortHistAsc; else { _sortHistKey = key; _sortHistAsc = true; }
                foreach (DataGridViewColumn c in _gridHistorico.Columns) c.HeaderCell.SortGlyphDirection = SortOrder.None;
                col.HeaderCell.SortGlyphDirection = _sortHistAsc ? SortOrder.Ascending : SortOrder.Descending;
            }

            private void ToggleSortAluno(DataGridViewColumn col)
            {
                if (_viewAluno == null) return;
                var name = col.DataPropertyName;
                if (string.IsNullOrWhiteSpace(name)) name = col.Name;
                // alterna asc/desc lendo o Sort atual
                bool toAsc = !(_viewAluno.Sort?.StartsWith("[" + name + "] ASC") ?? false);
                _viewAluno.Sort = $"[{name}] {(toAsc ? "ASC" : "DESC")}";
                foreach (DataGridViewColumn c in _gridAluno.Columns) c.HeaderCell.SortGlyphDirection = SortOrder.None;
                col.HeaderCell.SortGlyphDirection = toAsc ? SortOrder.Ascending : SortOrder.Descending;
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
