using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class RelatorioForm : Form
    {
    private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, ReadOnly = true, MultiSelect = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
    private BindingSource _bs = new BindingSource();
    private Button _export = new Button { Text = "Exportar Relatório" };
    private TextBox _search = new TextBox();
    private List<Relatorio> _all = new List<Relatorio>();
    private string _searchText = string.Empty;
    private string? _sortKey = null; // columns
    private bool _sortAsc = true;
    private readonly System.Windows.Forms.Timer _refreshTimer = new System.Windows.Forms.Timer();
    private string? _lastSeenChange;

        public RelatorioForm()
        {
            Text = "Relatórios";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Width = 1000;
            Height = 600;

            var panelTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12, 8, 12, 8), BackColor = Color.White };
            // export button style
            _export.AutoSize = true;
            _export.Padding = new Padding(10, 6, 10, 6);
            _export.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            _export.BackColor = Color.FromArgb(0, 123, 255);
            _export.ForeColor = Color.White;
            _export.FlatStyle = FlatStyle.Flat;
            _export.FlatAppearance.BorderSize = 0;
            _export.Tag = "accent";

            // search box with placeholder
            _search.Width = 260; _search.Font = new Font("Segoe UI", 10F); _search.Margin = new Padding(18, 6, 3, 3);
            try { _search.PlaceholderText = "🔎 Pesquisar..."; } catch { }
            _search.TextChanged += (_, __) => { _searchText = _search.Text?.Trim() ?? string.Empty; ApplyFiltersAndBind(); };
            panelTop.Controls.AddRange(new Control[] { _export, _search });

            // Grid styling
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

            var cChave = new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Relatorio.Chave), Width = 140, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cAluno = new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Relatorio.Aluno), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cProf = new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Relatorio.Professor), Width = 180, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cRet = new DataGridViewTextBoxColumn { HeaderText = "Retirada", DataPropertyName = nameof(Relatorio.DataHora), Width = 170, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" }, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cDev = new DataGridViewTextBoxColumn { HeaderText = "Devolução", DataPropertyName = nameof(Relatorio.DataDevolucao), Width = 170, DefaultCellStyle = new DataGridViewCellStyle { Format = "dd/MM/yyyy HH:mm:ss" }, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cTempo = new DataGridViewTextBoxColumn { HeaderText = "Tempo", DataPropertyName = nameof(Relatorio.TempoComChave), Width = 120, SortMode = DataGridViewColumnSortMode.Programmatic };
            var cTermo = new DataGridViewTextBoxColumn { HeaderText = "Termo", DataPropertyName = nameof(Relatorio.Termo), Width = 260, SortMode = DataGridViewColumnSortMode.Programmatic };
            _grid.Columns.AddRange(new DataGridViewColumn[] { cChave, cAluno, cProf, cRet, cDev, cTempo, cTermo });

            Controls.Add(_grid);
            Controls.Add(panelTop);

            Load += (_, __) => {
                try
                {
                    var t = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                    t = t switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => t };
                    if (t == "mica") { BeginInvoke(new Action(() => { try { MedControl.UI.FluentEffects.ApplyWin11Mica(this); } catch { } })); }
                }
                catch { }
                try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
                RefreshGrid();
                // setup auto-refresh polling for multiuser sync
                try
                {
                    _lastSeenChange = Database.GetConfig("last_change_at");
                    _refreshTimer.Interval = 2000;
                    _refreshTimer.Tick += (_, ____) =>
                    {
                        try
                        {
                            var cur = Database.GetConfig("last_change_at");
                            if (!string.Equals(cur, _lastSeenChange, StringComparison.Ordinal))
                            {
                                _lastSeenChange = cur;
                                RefreshGrid();
                            }
                        }
                        catch { }
                    };
                    _refreshTimer.Start();
                }
                catch { }
            };
            _grid.ColumnHeaderMouseClick += (_, e) => OnGridHeaderClick(e.ColumnIndex);
            _grid.CellDoubleClick += (_, e) => OpenTermo();
            _export.Click += (_, __) => Exportar();
        }

        private void RefreshGrid()
        {
            _all = Database.GetRelatorios();
            ApplyFiltersAndBind();
        }

        private void ApplyFiltersAndBind()
        {
            var q = _all.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(_searchText))
            {
                var s = _searchText.ToLowerInvariant();
                q = q.Where(r =>
                    (r.Chave ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.Aluno ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.Professor ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.Termo ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    (r.TempoComChave ?? string.Empty).ToLowerInvariant().Contains(s) ||
                    r.DataHora.ToString("dd/MM/yyyy HH:mm:ss").Contains(s) ||
                    (r.DataDevolucao?.ToString("dd/MM/yyyy HH:mm:ss") ?? string.Empty).Contains(s)
                );
            }
            if (!string.IsNullOrEmpty(_sortKey))
            {
                Func<Relatorio, object?> selector = _sortKey switch
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
                q = _sortAsc ? q.OrderBy(selector) : q.OrderByDescending(selector);
            }
            _bs.DataSource = q.ToList();
            _grid.DataSource = _bs;
        }

        private void OnGridHeaderClick(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= _grid.Columns.Count) return;
            var col = _grid.Columns[columnIndex];
            var key = GetSortKey(col);
            if (key == null) return;
            if (_sortKey == key) _sortAsc = !_sortAsc; else { _sortKey = key; _sortAsc = true; }
            ApplyFiltersAndBind();
            foreach (DataGridViewColumn c in _grid.Columns) c.HeaderCell.SortGlyphDirection = SortOrder.None;
            col.HeaderCell.SortGlyphDirection = _sortAsc ? SortOrder.Ascending : SortOrder.Descending;
        }

        private string? GetSortKey(DataGridViewColumn col)
        {
            switch (col.HeaderText)
            {
                case "Chave": return nameof(Relatorio.Chave);
                case "Aluno": return nameof(Relatorio.Aluno);
                case "Professor": return nameof(Relatorio.Professor);
                case "Retirada": return nameof(Relatorio.DataHora);
                case "Devolução": return nameof(Relatorio.DataDevolucao);
                case "Tempo": return nameof(Relatorio.TempoComChave);
                case "Termo": return nameof(Relatorio.Termo);
            }
            return null;
        }

        private void Exportar()
        {
            var dados = Database.GetRelatorios();
            if (dados.Count == 0)
            {
                MessageBox.Show("Não há dados no relatório para exportar.");
                return;
            }

            using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "Relatorio.xlsx" };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Relatório");
            // Cabeçalhos
            string[] headers = { "Chave", "Aluno", "Professor", "Data e Hora", "Data de Devolução", "Tempo com a Chave", "Termo de Compromisso", "Item Atribuído" };
            for (int i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];

            // Linhas
            int row = 2;
            foreach (var r in dados)
            {
                ws.Cell(row, 1).Value = r.Chave;
                ws.Cell(row, 2).Value = r.Aluno ?? "";
                ws.Cell(row, 3).Value = r.Professor ?? "";
                ws.Cell(row, 4).Value = r.DataHora;
                ws.Cell(row, 5).Value = r.DataDevolucao;
                ws.Cell(row, 6).Value = r.TempoComChave ?? "";

                var termo = r.Termo ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(termo) && File.Exists(termo))
                {
                    var cell = ws.Cell(row, 7);
                    cell.Value = "Abrir PDF";
                    cell.GetHyperlink().ExternalAddress = new Uri(termo);
                    ws.Cell(row, 8).Value = "Sim";
                }
                else
                {
                    ws.Cell(row, 7).Value = termo;
                    ws.Cell(row, 8).Value = "Não";
                }
                row++;
            }

            wb.SaveAs(sfd.FileName);
            MessageBox.Show($"Relatório exportado com sucesso para '{sfd.FileName}'.");
        }

        private void OpenTermo()
        {
            if (_bs.Current is Relatorio r && !string.IsNullOrWhiteSpace(r.Termo) && File.Exists(r.Termo))
            {
                try { Process.Start(new ProcessStartInfo { FileName = r.Termo, UseShellExecute = true }); }
                catch { MessageBox.Show("Não foi possível abrir o PDF."); }
            }
        }
    }
}
