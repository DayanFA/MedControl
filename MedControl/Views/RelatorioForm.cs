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
    // Excel-centric buttons (unify style with CadastroAlunosForm)
    private Button _importExcel = new Button { Text = "⬇️ Importar" };
    private Button _exportExcel = new Button { Text = "📤 Exportar" };
    private TextBox _search = new TextBox();
    private List<Relatorio> _all = new List<Relatorio>();
    private string _searchText = string.Empty;
    private string? _sortKey = null; // columns
    private bool _sortAsc = true;
    private readonly System.Windows.Forms.Timer _refreshTimer = new System.Windows.Forms.Timer();
    private string? _lastSeenChange;

        public RelatorioForm()
        {
            Text = "Relatórios de Chaves";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Width = 1000;
            Height = 600;

            var panelTop = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12, 8, 12, 8), BackColor = Color.White };

            // Unified button styling (copy approach from CadastroAlunosForm)
            void ApplyButtonStyle(Button btn, string text, Color baseColor)
            {
                btn.Text = text;
                btn.AutoSize = false;
                btn.Size = new Size(150, 44);
                btn.MinimumSize = new Size(150, 44);
                btn.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.BackColor = baseColor;
                btn.ForeColor = Color.White;
                btn.Cursor = Cursors.Hand;
                btn.Margin = new Padding(4, 6, 4, 6);
                Color Darken(Color c, int amt) => Color.FromArgb(c.A, Math.Max(0, c.R - amt), Math.Max(0, c.G - amt), Math.Max(0, c.B - amt));
                var hover = Darken(baseColor, 12);
                var pressed = Darken(baseColor, 22);
                btn.MouseEnter += (_, __) => btn.BackColor = hover;
                btn.MouseLeave += (_, __) => btn.BackColor = baseColor;
                btn.MouseDown += (_, __) => btn.BackColor = pressed;
                btn.MouseUp += (_, __) => btn.BackColor = hover;
            }

            ApplyButtonStyle(_importExcel, "⬇️ Importar", Color.FromArgb(33, 150, 243));
            ApplyButtonStyle(_exportExcel, "📤 Exportar", Color.FromArgb(76, 175, 80));

            // search box with placeholder and aligned height
            _search.Width = 260; _search.Font = new Font("Segoe UI", 10F); _search.Margin = new Padding(18, 6, 3, 6); _search.Height = 44;
            try { _search.PlaceholderText = "🔎 Pesquisar..."; } catch { }
            _search.TextChanged += (_, __) => { _searchText = _search.Text?.Trim() ?? string.Empty; ApplyFiltersAndBind(); };

            panelTop.Controls.AddRange(new Control[] { _importExcel, _exportExcel, _search });

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
            _exportExcel.Click += (_, __) => ExportarExcel();
            _importExcel.Click += (_, __) => ImportarExcel();
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

        private void ExportarExcel()
        {
            var dados = Database.GetRelatorios();
            if (dados.Count == 0)
            {
                MessageBox.Show("Não há dados no relatório para exportar.");
                return;
            }

            using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "Relacao_Chaves.xlsx" };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Relação de Chaves");
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

        // New Excel import replacing CSV import
        private void ImportarExcel()
        {
            using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", Title = "Importar Excel - Relação de Chaves" };
            if (ofd.ShowDialog(this) != DialogResult.OK) return;
            int imported = 0; int skipped = 0; int errors = 0;
            var existing = Database.GetRelatorios();
            try
            {
                using var wb = new XLWorkbook(ofd.FileName);
                var ws = wb.Worksheets.First();
                // Build header map (case-insensitive)
                var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var cell in ws.Row(1).CellsUsed())
                {
                    var h = cell.GetString().Trim();
                    if (!string.IsNullOrEmpty(h) && !headers.ContainsKey(h)) headers[h] = cell.Address.ColumnNumber;
                }

                // Accept multiple header variants
                int Col(string name, params string[] aliases)
                {
                    if (headers.TryGetValue(name, out var idx)) return idx;
                    foreach (var a in aliases) if (headers.TryGetValue(a, out idx)) return idx;
                    return -1;
                }

                int cChave = Col("Chave");
                int cAluno = Col("Aluno");
                int cProf = Col("Professor");
                int cDataHora = Col("DataHora", "Data e Hora", "Retirada");
                int cDataDev = Col("DataDevolucao", "Data de Devolução", "Devolução");
                int cTempo = Col("TempoComChave", "Tempo com a Chave", "Tempo");
                int cTermo = Col("Termo", "Termo de Compromisso");

                if (cChave == -1 || cDataHora == -1)
                {
                    MessageBox.Show("Planilha inválida: colunas obrigatórias ausentes (Chave, DataHora).", "Importar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var lastUsed = ws.LastRowUsed();
                int lastRow = lastUsed?.RowNumber() ?? 1;
                for (int r = 2; r <= lastRow; r++)
                {
                    try
                    {
                        var row = ws.Row(r);
                        string chave = cChave > 0 ? row.Cell(cChave).GetString().Trim() : string.Empty;
                        if (string.IsNullOrWhiteSpace(chave)) continue; // skip empty
                        string aluno = cAluno > 0 ? row.Cell(cAluno).GetString().Trim() : string.Empty;
                        string prof = cProf > 0 ? row.Cell(cProf).GetString().Trim() : string.Empty;
                        string tempo = cTempo > 0 ? row.Cell(cTempo).GetString().Trim() : string.Empty;
                        string termo = cTermo > 0 ? row.Cell(cTermo).GetString().Trim() : string.Empty;

                        // DataHora parsing
                        DateTime dataHora;
                        var rawDataHora = cDataHora > 0 ? row.Cell(cDataHora).GetString().Trim() : string.Empty;
                        if (string.IsNullOrWhiteSpace(rawDataHora) && cDataHora > 0 && row.Cell(cDataHora).DataType == XLDataType.DateTime)
                            dataHora = row.Cell(cDataHora).GetDateTime();
                        else if (!TryParseDate(rawDataHora, out dataHora)) { errors++; continue; }

                        // DataDevolucao optional
                        DateTime? dataDev = null;
                        if (cDataDev > 0)
                        {
                            var rawDataDev = row.Cell(cDataDev).GetString().Trim();
                            if (string.IsNullOrWhiteSpace(rawDataDev) && row.Cell(cDataDev).DataType == XLDataType.DateTime)
                                dataDev = row.Cell(cDataDev).GetDateTime();
                            else if (!string.IsNullOrWhiteSpace(rawDataDev) && TryParseDate(rawDataDev, out var tmp)) dataDev = tmp;
                            else if (!string.IsNullOrWhiteSpace(rawDataDev)) { errors++; continue; }
                        }

                        var rel = new Relatorio
                        {
                            Chave = chave,
                            Aluno = aluno,
                            Professor = prof,
                            DataHora = dataHora,
                            DataDevolucao = dataDev,
                            TempoComChave = tempo,
                            Termo = termo
                        };

                        if (existing.Any(e => e.Chave == rel.Chave && e.Aluno == rel.Aluno && e.Professor == rel.Professor && e.DataHora == rel.DataHora))
                        { skipped++; continue; }
                        try { Database.InsertRelatorio(rel); imported++; }
                        catch { errors++; }
                    }
                    catch { errors++; }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha ao importar: " + ex.Message, "Importar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            RefreshGrid();
            MessageBox.Show($"Importação concluída. Inseridos: {imported}. Ignorados (duplicados): {skipped}. Erros: {errors}.", "Importar", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool TryParseDate(string input, out DateTime dt)
        {
            var formats = new[]{ "dd/MM/yyyy HH:mm:ss", "dd/MM/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm" };
            foreach (var f in formats)
            {
                if (DateTime.TryParseExact(input, f, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt)) return true;
            }
            // fallback generic parse
            if (DateTime.TryParse(input, out dt)) return true;
            dt = DateTime.MinValue; return false;
        }

        // (Old secondary button styling removed after Excel-centric redesign)

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
