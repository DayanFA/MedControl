using System;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class CadastroProfessoresForm : Form
    {
    // Grid & Data
    private readonly DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = true };
    private DataTable _table = new DataTable();
    private DataTable _filtered = new DataTable();
    private bool _handlingCellAction = false; // evita duplo disparo (CellClick + CellContentClick)

    // Top bar
    private readonly FlowLayoutPanel _top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 64, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(8) };
    private readonly TextBox _search = new TextBox { Width = 300, PlaceholderText = "Pesquisar... 🔎" };
    private readonly Button _add = new Button();
    private readonly Button _import = new Button();
    private readonly Button _export = new Button();

    // Pagination
    private readonly Panel _bottomBar = new Panel { Dock = DockStyle.Bottom, Height = 48 };
    private readonly Button _prevBtn = new Button();
    private readonly Button _nextBtn = new Button();
    private readonly FlowLayoutPanel _pagesPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
    private readonly ComboBox _pageSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70, Visible = false };
    private int _pageSize = 10;
    private int _currentPage = 1;
    private string? _lastSortColumn = null;
    private bool _sortAsc = true;

        public CadastroProfessoresForm()
        {
            Text = "Cadastro de Professores";
            Width = 900; Height = 650;

            // Grid visuals (modern look similar to Alunos)
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // permitir edição inline nas células (exceto colunas de ação e _id)
            _grid.ReadOnly = false;
            _grid.RowHeadersVisible = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.BorderStyle = BorderStyle.None;
            _grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            _grid.EnableHeadersVisualStyles = false;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 242, 245);
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
            _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(220, 235, 255);
            _grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            // Keep rows uniform (no alternating row color)
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Color.White;
            _grid.RowTemplate.Height = 34;
            _grid.ColumnHeaderMouseClick += (_, e) => OnHeaderClick(e.ColumnIndex);
            _grid.CellEndEdit += Grid_CellEndEdit;

            Controls.Add(_grid);
            Controls.Add(_top);
            Controls.Add(_bottomBar);

            Load += (_, __) =>
            {
                try { Database.Setup(); } catch { }
                try
                {
                    _table = Database.GetProfessoresAsDataTable();
                }
                catch { _table = new DataTable(); }

                // Se base estiver vazia, carrega pré-visualização da planilha (não persiste até o upload)
                if (_table.Rows.Count == 0)
                {
                    var path = Database.GetConfig("caminho_professores");
                    if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                        _table = ExcelHelper.LoadToDataTable(path);
                }
                EnsureIds();
                if (_table.Columns.Count == 1)
                {
                    if (!_table.Columns.Contains("Nome")) _table.Columns.Add("Nome");
                }
                _filtered = _table.Copy();
                ApplyFilterAndRefresh();
                try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
            };

            _import.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        UseWaitCursor = true; Enabled = false;
                        _table = ExcelHelper.LoadToDataTable(ofd.FileName);
                        EnsureIds();
                        _filtered = _table.Copy();
                        _currentPage = 1;
                        ApplyFilterAndRefresh();
                        Database.SetConfig("caminho_professores", ofd.FileName);

                        // Persiste toda a planilha na base, como em Alunos
                        _ = System.Threading.Tasks.Task.Run(() => Database.ReplaceAllProfessores(_table))
                            .ContinueWith(_ =>
                            {
                                try
                                {
                                    var fresh = Database.GetProfessoresAsDataTable();
                                    BeginInvoke(new Action(() => { _table = fresh; ApplyFilterAndRefresh(); }));
                                }
                                catch { }
                            })
                            .ContinueWith(_ => BeginInvoke(new Action(() => { Enabled = true; UseWaitCursor = false; })));
                    }
                    catch (Exception ex)
                    {
                        Enabled = true; UseWaitCursor = false;
                        try { MessageBox.Show(this, "Falha ao importar: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                    }
                }
            };
            _export.Click += (_, __) =>
            {
                using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "cadastro_professores.xlsx" };
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ExcelHelper.SaveDataTable(sfd.FileName, _table);
                    MessageBox.Show("Exportado com sucesso.");
                }
            };
            // Removido botão 'Caminho' por solicitação

            // Build top bar with emoji buttons and aligned search (square buttons)
            var actions = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(4), Margin = new Padding(0) };
            // Mesmo padrão do Alunos: Adicionar (azul primário), Upload (azul 33,150,243), Exportar (verde)
            ConfigureActionButton(_add, "➕ Adicionar", Color.FromArgb(0, 123, 255));
            ConfigureActionButton(_import, "⬇️ Upload", Color.FromArgb(33, 150, 243));
            ConfigureActionButton(_export, "📤 Exportar", Color.FromArgb(76, 175, 80));
            _add.Click += (_, __) => AddProfessor();
            actions.Controls.AddRange(new Control[] { _add, _import, _export });
            actions.Margin = new Padding(0, 6, 0, 6);
            _top.Controls.Add(actions);
            _search.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            _search.Height = 44;
            _search.Margin = new Padding(12, 6, 8, 6);
            _search.BorderStyle = BorderStyle.FixedSingle;
            _search.TextChanged += (_, __) => { _currentPage = 1; ApplyFilterAndRefresh(); };
            // Keep search box square (no rounded corners)
            _search.Region = null;
            _top.Controls.Add(_search);

            // Bottom bar pagination
            StylePagerButton(_prevBtn, "‹");
            StylePagerButton(_nextBtn, "›");
            _prevBtn.Click += (_, __) => { if (_currentPage > 1) { _currentPage--; RefreshGrid(); } };
            _nextBtn.Click += (_, __) => { if (_currentPage < TotalPages) { _currentPage++; RefreshGrid(); } };
            _pageSelector.SelectedIndexChanged += (_, __) => { if (_pageSelector.SelectedItem is int p) { _currentPage = p; _pageSelector.Visible = false; RefreshGrid(); } };
            _pagesPanel.Controls.Add(_prevBtn);
            _pagesPanel.Controls.Add(_pageSelector);
            _pagesPanel.Controls.Add(_nextBtn);
            _bottomBar.Controls.Add(_pagesPanel);
            // Use apenas CellContentClick para botões; evitar duplicidade com CellClick
            _grid.CellContentClick += Grid_CellClick;
            _grid.DataBindingComplete += (_, __) => { EnsureActionColumnsAtEnd(); ApplyReadOnlyRules(); };
        }

        private void Grid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var col = _grid.Columns[e.ColumnIndex];
            var colName = col.DataPropertyName ?? col.Name ?? col.HeaderText;
            if (string.IsNullOrEmpty(colName)) return;
            if (colName == "EDITAR" || colName == "EXCLUIR" || string.Equals(colName, "_id", StringComparison.OrdinalIgnoreCase)) return;

            var row = GetDataRowFromGridRow(e.RowIndex);
            if (row == null) return;

            var newVal = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty;
            row[colName] = newVal;
            try
            {
                var id = Convert.ToString(row["_id"]) ?? Guid.NewGuid().ToString();
                var data = BuildDictFromRow(row);
                Database.UpsertProfessor(id, data);
                // Evita reentrância do DataGridView (mudar DataSource dentro do evento de edição)
                BeginInvoke(new Action(() => ReloadFromDatabase()));
            }
            catch
            {
                // Defer também em caso de erro para evitar conflitos com o ciclo de edição
                BeginInvoke(new Action(() => ApplyFilterAndRefresh()));
            }
        }

        private DataRow? GetDataRowFromGridRow(int gridRowIndex)
        {
            if (gridRowIndex < 0 || gridRowIndex >= _grid.Rows.Count) return null;
            var gridRow = _grid.Rows[gridRowIndex];
            // Localiza coluna do _id (mesmo oculta)
            DataGridViewColumn? idCol = null;
            foreach (DataGridViewColumn c in _grid.Columns)
            {
                if (string.Equals(c.DataPropertyName, "_id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(c.Name, "_id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(c.HeaderText, "_id", StringComparison.OrdinalIgnoreCase))
                { idCol = c; break; }
            }
            if (idCol == null) return null;
            var val = gridRow.Cells[idCol.Index].Value;
            var id = val?.ToString();
            if (string.IsNullOrEmpty(id)) return null;
            return _table.AsEnumerable().FirstOrDefault(r => Convert.ToString(r["_id"]) == id);
        }

        private void ApplyReadOnlyRules()
        {
            try
            {
                foreach (DataGridViewColumn c in _grid.Columns)
                {
                    if (c.Name == "EDITAR" || c.Name == "EXCLUIR" || string.Equals(c.Name, "_id", StringComparison.OrdinalIgnoreCase))
                        c.ReadOnly = true;
                    else
                        c.ReadOnly = false;
                }
            }
            catch { }
        }

        private void ConfigureActionButton(Button btn, string text, Color baseCol)
        {
            var size = new Size(150, 44);
            btn.Text = text;
            btn.AutoSize = false;
            btn.Size = size;
            btn.MinimumSize = size;
            btn.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            btn.ImageAlign = ContentAlignment.MiddleLeft;
            btn.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Cursor = Cursors.Hand;
            btn.BackColor = baseCol;
            btn.ForeColor = Color.White;
            btn.Tag = "accent";
            var hover = Darken(baseCol, 12);
            var pressed = Darken(baseCol, 22);
            btn.MouseEnter += (_, __) => btn.BackColor = hover;
            btn.MouseLeave += (_, __) => btn.BackColor = baseCol;
            btn.MouseDown += (_, __) => btn.BackColor = pressed;
            btn.MouseUp += (_, __) => btn.BackColor = hover;
        }

        private void StylePagerButton(Button b, string text)
        {
            var accent = Color.FromArgb(0, 123, 255);
            b.Text = text; b.AutoSize = true; b.FlatStyle = FlatStyle.Flat; b.Padding = new Padding(8);
            b.BackColor = Color.White; b.ForeColor = accent;
            b.FlatAppearance.BorderSize = 1; b.FlatAppearance.BorderColor = accent;
            b.MouseEnter += (_, __) => { b.BackColor = accent; b.ForeColor = Color.White; };
            b.MouseLeave += (_, __) => { b.BackColor = Color.White; b.ForeColor = accent; };
        }

        private static void Roundify(Control control, int radius)
        {
            if (control.Width <= 0 || control.Height <= 0) return;
            int diameter = radius * 2;
            Rectangle bounds = new Rectangle(0, 0, control.Width, control.Height);
            using var path = new GraphicsPath();
            path.StartFigure();
            path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            control.Region = new Region(path);
        }

        private static Color Darken(Color color, int amount)
        {
            int r = Math.Max(0, color.R - amount);
            int g = Math.Max(0, color.G - amount);
            int b = Math.Max(0, color.B - amount);
            return Color.FromArgb(color.A, r, g, b);
        }

    private int TotalPages => Math.Max(1, (int)Math.Ceiling((_filtered?.Rows.Count ?? 0) / (double)_pageSize));

        private void OnHeaderClick(int columnIndex)
        {
            if (columnIndex < 0 || _filtered == null) return;
            var col = _grid.Columns[columnIndex];
            var colName = col.DataPropertyName ?? col.Name ?? col.HeaderText;
            if (string.IsNullOrEmpty(colName)) return;
            _sortAsc = _lastSortColumn == colName ? !_sortAsc : true;
            _lastSortColumn = colName;
            try
            {
                var dv = _filtered.DefaultView;
                dv.Sort = colName + (_sortAsc ? " ASC" : " DESC");
                _filtered = dv.ToTable();
                _currentPage = 1;
                RefreshGrid();
            }
            catch { }
        }

        private void ApplyFilterAndRefresh()
        {
            var term = _search.Text?.Trim();
            if (string.IsNullOrEmpty(term)) _filtered = _table.Copy();
            else
            {
                var result = _table.Clone();
                foreach (DataRow r in _table.Rows)
                {
                    bool match = false;
                    foreach (DataColumn c in _table.Columns)
                    {
                        var val = r[c]?.ToString() ?? string.Empty;
                        if (val.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) { match = true; break; }
                    }
                    if (match) result.ImportRow(r);
                }
                _filtered = result;
            }
            _currentPage = Math.Min(Math.Max(1, _currentPage), TotalPages);
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            if (_filtered == null) _filtered = _table.Copy();
            var dt = _filtered.Clone();
            int start = (_currentPage - 1) * _pageSize;
            for (int i = start; i < Math.Min(start + _pageSize, _filtered.Rows.Count); i++) dt.ImportRow(_filtered.Rows[i]);
            _grid.DataSource = dt;
            // hide internal row id if exists
            if (_grid.Columns.Contains("_id")) _grid.Columns["_id"].Visible = false;
            EnsureButtonsColumn();
            EnsureActionColumnsAtEnd();
            UpdatePaginationControls();
        }

        private void UpdatePaginationControls()
        {
            _pagesPanel.SuspendLayout();
            _pagesPanel.Controls.Clear();
            _pagesPanel.Controls.Add(_prevBtn);

            int total = TotalPages;
            int show = 5;
            int start = Math.Max(1, _currentPage - 2);
            int end = Math.Min(total, start + show - 1);
            if (end - start + 1 < show) start = Math.Max(1, end - show + 1);

            if (start > 1)
            {
                var firstBtn = new Button { Text = "1", AutoSize = true, Padding = new Padding(6), Margin = new Padding(4) };
                firstBtn.BackColor = 1 == _currentPage ? Color.FromArgb(0, 123, 255) : Color.White;
                firstBtn.ForeColor = 1 == _currentPage ? Color.White : Color.Black;
                firstBtn.Click += (_, __) => { _currentPage = 1; _pageSelector.Visible = false; RefreshGrid(); };
                _pagesPanel.Controls.Add(firstBtn);

                if (start > 2)
                {
                    var leftEll = new LinkLabel { Text = "...", AutoSize = true, LinkColor = Color.Black, ActiveLinkColor = Color.DimGray, Padding = new Padding(10) };
                    leftEll.LinkClicked += (_, __) =>
                    {
                        _pageSelector.Items.Clear();
                        for (int p = 1; p < start; p++) _pageSelector.Items.Add(p);
                        if (_pageSelector.Items.Count > 0)
                        {
                            _pageSelector.SelectedIndex = 0;
                            _pageSelector.Visible = true; _pageSelector.DroppedDown = true; _pageSelector.Focus();
                        }
                    };
                    _pagesPanel.Controls.Add(leftEll);
                }
            }

            for (int p = start; p <= end; p++)
            {
                var b = new Button { Text = p.ToString(), AutoSize = true, Padding = new Padding(6), Margin = new Padding(4) };
                b.BackColor = p == _currentPage ? Color.FromArgb(0, 123, 255) : Color.White;
                b.ForeColor = p == _currentPage ? Color.White : Color.Black;
                int page = p;
                b.Click += (_, __) => { _currentPage = page; _pageSelector.Visible = false; RefreshGrid(); };
                _pagesPanel.Controls.Add(b);
            }

            if (end < total)
            {
                if (end < total - 1)
                {
                    var rightEll = new LinkLabel { Text = "...", AutoSize = true, LinkColor = Color.Black, ActiveLinkColor = Color.DimGray, Padding = new Padding(10) };
                    rightEll.LinkClicked += (_, __) =>
                    {
                        _pageSelector.Items.Clear();
                        for (int p = end + 1; p <= total; p++) _pageSelector.Items.Add(p);
                        if (_pageSelector.Items.Count > 0)
                        {
                            _pageSelector.SelectedIndex = 0;
                            _pageSelector.Visible = true; _pageSelector.DroppedDown = true; _pageSelector.Focus();
                        }
                    };
                    _pagesPanel.Controls.Add(rightEll);
                }

                var lastBtn = new Button { Text = total.ToString(), AutoSize = true, Padding = new Padding(6), Margin = new Padding(4) };
                lastBtn.BackColor = total == _currentPage ? Color.FromArgb(0, 123, 255) : Color.White;
                lastBtn.ForeColor = total == _currentPage ? Color.White : Color.Black;
                int lastPage = total;
                lastBtn.Click += (_, __) => { _currentPage = lastPage; _pageSelector.Visible = false; RefreshGrid(); };
                _pagesPanel.Controls.Add(lastBtn);
            }
            else _pageSelector.Visible = false;

            _pagesPanel.Controls.Add(_nextBtn);
            _pagesPanel.ResumeLayout();
        }

        private void EnsureIds()
        {
            if (!_table.Columns.Contains("_id")) _table.Columns.Add("_id", typeof(string));
            foreach (DataRow r in _table.Rows)
            {
                if (r["_id"] == DBNull.Value || string.IsNullOrEmpty(Convert.ToString(r["_id"])) )
                {
                    r["_id"] = Guid.NewGuid().ToString();
                }
            }
        }

        private void EnsureButtonsColumn()
        {
            if (_grid.Columns.Contains("EDITAR") && _grid.Columns.Contains("EXCLUIR")) return;
            var editCol = new DataGridViewButtonColumn
            {
                Name = "EDITAR",
                Text = "✏️ Editar",
                UseColumnTextForButtonValue = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FlatStyle = FlatStyle.Flat
            };
            editCol.DefaultCellStyle.BackColor = Color.FromArgb(0, 123, 255);
            editCol.DefaultCellStyle.ForeColor = Color.White;
            editCol.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 123, 255);
            editCol.DefaultCellStyle.SelectionForeColor = Color.White;

            var delCol = new DataGridViewButtonColumn
            {
                Name = "EXCLUIR",
                Text = "🗑️ Excluir",
                UseColumnTextForButtonValue = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FlatStyle = FlatStyle.Flat
            };
            delCol.DefaultCellStyle.BackColor = Color.FromArgb(220, 53, 69); // red
            delCol.DefaultCellStyle.ForeColor = Color.White;
            delCol.DefaultCellStyle.SelectionBackColor = Color.FromArgb(220, 53, 69);
            delCol.DefaultCellStyle.SelectionForeColor = Color.White;

            _grid.Columns.Add(editCol);
            _grid.Columns.Add(delCol);
            // enforce colors via CellFormatting to avoid gray overrides
            _grid.CellFormatting -= Grid_CellFormatting;
            _grid.CellFormatting += Grid_CellFormatting;
        }

        private void EnsureActionColumnsAtEnd()
        {
            try
            {
                if (_grid.Columns.Contains("EDITAR") && _grid.Columns.Contains("EXCLUIR"))
                {
                    var lastIndex = _grid.Columns.Count - 1;
                    _grid.Columns["EXCLUIR"].DisplayIndex = lastIndex;
                    _grid.Columns["EDITAR"].DisplayIndex = Math.Max(0, lastIndex - 1);
                }
            }
            catch { }
        }

        private void Grid_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (_handlingCellAction) return; // evita reentrância
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var colName = _grid.Columns[e.ColumnIndex].Name;
            if (colName != "EDITAR" && colName != "EXCLUIR") return;

            try
            {
                _handlingCellAction = true;
                // get row id from bound page table
                string? rid = null;
                if (_grid.Columns.Contains("_id"))
                {
                    var val = _grid.Rows[e.RowIndex].Cells["_id"].Value;
                    rid = val?.ToString();
                }
                if (string.IsNullOrEmpty(rid)) return;

                var orig = _table.AsEnumerable().FirstOrDefault(r => Convert.ToString(r["_id"]) == rid);
                if (orig == null) return;

                if (colName == "EXCLUIR")
                {
                    var res = MessageBox.Show(this, "Excluir este registro?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (res == DialogResult.Yes)
                    {
                        var id = Convert.ToString(orig["_id"])!;
                        _table.Rows.Remove(orig);
                        try { Database.DeleteProfessor(id); ReloadFromDatabase(); }
                        catch { _filtered = _table.Copy(); ApplyFilterAndRefresh(); }
                    }
                }
                else if (colName == "EDITAR")
                {
                    var values = MedControl.Views.AddEditForm.ShowDialog("Editar professor", _table.Columns, orig);
                    if (values != null)
                    {
                        foreach (DataColumn c in _table.Columns)
                        {
                            if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                            if (values.TryGetValue(c.ColumnName, out var v)) orig[c.ColumnName] = v;
                        }
                        try
                        {
                            var id = Convert.ToString(orig["_id"]) ?? Guid.NewGuid().ToString();
                            var data = BuildDictFromRow(orig);
                            Database.UpsertProfessor(id, data);
                            ReloadFromDatabase();
                        }
                        catch { _filtered = _table.Copy(); ApplyFilterAndRefresh(); }
                    }
                }
            }
            finally
            {
                _handlingCellAction = false;
            }
        }

        private void Grid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0) return;
            if (e.CellStyle == null) return;
            var col = _grid.Columns[e.ColumnIndex];
            if (col.Name == "EDITAR")
            {
                var azul = Color.FromArgb(0, 123, 255);
                e.CellStyle.BackColor = azul;
                e.CellStyle.ForeColor = Color.White;
                e.CellStyle.SelectionBackColor = azul;
                e.CellStyle.SelectionForeColor = Color.White;
            }
            else if (col.Name == "EXCLUIR")
            {
                var vermelho = Color.FromArgb(220, 53, 69);
                e.CellStyle.BackColor = vermelho;
                e.CellStyle.ForeColor = Color.White;
                e.CellStyle.SelectionBackColor = vermelho;
                e.CellStyle.SelectionForeColor = Color.White;
            }
        }

        private void AddProfessor()
        {
            // Adiciona baseado somente nas colunas atuais da planilha carregada
            if (_table.Columns.Count == 0)
            {
                MessageBox.Show(this, "Antes de adicionar, importe uma planilha para definir as colunas.", "Sem colunas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var values = MedControl.Views.AddEditForm.ShowDialog("Adicionar professor", _table.Columns, null);
            if (values != null)
            {
                var nr = _table.NewRow();
                // garante coluna de id interno e valor (usa _id, como em Alunos)
                if (!_table.Columns.Contains("_id")) _table.Columns.Add("_id", typeof(string));
                foreach (DataColumn c in _table.Columns)
                {
                    if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase))
                    {
                        nr[c.ColumnName] = Guid.NewGuid().ToString();
                    }
                    else if (values.TryGetValue(c.ColumnName, out var v)) nr[c.ColumnName] = v;
                    else nr[c.ColumnName] = string.Empty;
                }
                _table.Rows.Add(nr);
                try
                {
                    var id = Convert.ToString(nr["_id"]) ?? Guid.NewGuid().ToString();
                    var data = BuildDictFromRow(nr);
                    Database.UpsertProfessor(id, data);
                    ReloadFromDatabase();
                    _currentPage = 1;
                    ApplyFilterAndRefresh();
                    try { MessageBox.Show(this, "Registro adicionado.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                }
                catch
                {
                    _filtered = _table.Copy();
                    _currentPage = 1;
                    ApplyFilterAndRefresh();
                }
            }
        }

        // Helpers para persistência (igual aos Alunos)
        private System.Collections.Generic.Dictionary<string, string> BuildDictFromRow(DataRow row)
        {
            var dict = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn c in row.Table.Columns)
            {
                if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                dict[c.ColumnName] = Convert.ToString(row[c]) ?? string.Empty;
            }
            return dict;
        }

        private void ReloadFromDatabase()
        {
            try
            {
                var fresh = Database.GetProfessoresAsDataTable();
                _table = fresh;
                ApplyFilterAndRefresh();
            }
            catch { }
        }
    }
}
