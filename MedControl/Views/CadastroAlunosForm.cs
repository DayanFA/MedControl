using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class CadastroAlunosForm : Form
    {
    // Core UI
    private readonly DataGridView _grid = new DataGridView { Dock = DockStyle.None, AutoGenerateColumns = true };
    private readonly Panel _container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(30, 12, 30, 12) }; // container with margin so grid is well-distributed
    private readonly TextBox _searchBox = new TextBox { Width = 300, PlaceholderText = "Pesquisar... 🔎" };
    private readonly FlowLayoutPanel _topPanel = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 120, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(8) };
    private readonly Button _importBtn = CreateButton("Importar", Color.FromArgb(33, 150, 243));
    private readonly Button _exportBtn = CreateButton("Exportar", Color.FromArgb(76, 175, 80));
    private readonly Button _addBtn = CreateButton("Adicionar", Color.FromArgb(0, 123, 255));

    // Persistência agora é baseada em banco (Database.cs); removido _sharedTable
    // Tooltips for icon-only buttons
    private readonly ToolTip _buttonToolTip = new ToolTip();
    private readonly Panel _bottomBar = new Panel { Dock = DockStyle.Bottom, Height = 48 };

    // Pagination controls
        private readonly Button _prevBtn = CreateButton("‹", Color.Transparent);
        private readonly Button _nextBtn = CreateButton("›", Color.Transparent);
        private readonly FlowLayoutPanel _pagesPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
        private readonly ComboBox _pageSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70, Visible = false };

    // Sorting state (header click toggles)
    private string? _lastSortColumn = null;
    private bool _sortAscending = true;

        // Data
        private DataTable _fullTable = new DataTable();
        private DataTable _filteredTable = new DataTable();
        private int _pageSize = 10;
        private int _currentPage = 1;
    private bool _actionColorsHooked = false;

        public CadastroAlunosForm()
        {
            Text = "Cadastro de Alunos";
            Width = 1000; Height = 700;
            BackColor = Color.White;

            // Layout: top search row + grid filling remaining area inside the container with horizontal padding
            // topPanel is a FlowLayoutPanel so search and action buttons sit next to each other
            _searchBox.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            _topPanel.Controls.Add(_searchBox);
            _container.Controls.Add(_grid);
            _container.Controls.Add(_topPanel);
            _grid.Dock = DockStyle.Fill;
            Controls.Add(_container);
            Controls.Add(_bottomBar);

            // Bottom bar: pagination
            _prevBtn.AutoSize = true; _nextBtn.AutoSize = true; _prevBtn.FlatStyle = FlatStyle.Flat; _nextBtn.FlatStyle = FlatStyle.Flat;
            _prevBtn.Padding = new Padding(8); _nextBtn.Padding = new Padding(8);
            // Improve arrow visibility with accent color and hover effect
            var accent = Color.FromArgb(0, 123, 255);
            foreach (var b in new[] { _prevBtn, _nextBtn })
            {
                b.BackColor = Color.White;
                b.ForeColor = accent;
                b.FlatAppearance.BorderSize = 1;
                b.FlatAppearance.BorderColor = accent;
                var btn = b; // avoid closure capture issues
                btn.MouseEnter += (_, __) => { btn.BackColor = accent; btn.ForeColor = Color.White; };
                btn.MouseLeave += (_, __) => { btn.BackColor = Color.White; btn.ForeColor = accent; };
            }
            _prevBtn.Click += (_, __) => { if (_currentPage > 1) { _currentPage--; RefreshGrid(); } };
            _nextBtn.Click += (_, __) => { if (_currentPage < TotalPages) { _currentPage++; RefreshGrid(); } };

            _pageSelector.SelectedIndexChanged += (_, __) =>
            {
                if (_pageSelector.SelectedItem is int p)
                {
                    _currentPage = p; _pageSelector.Visible = false; RefreshGrid();
                }
            };

            _searchBox.TextChanged += (_, __) => { _currentPage = 1; ApplyFilterAndRefresh(); };

            _pagesPanel.Controls.Add(_prevBtn);
            // keep page selector in the panel (hidden by default) so it can drop down near the pagination controls
            _pageSelector.Visible = false;
            _pageSelector.DropDownWidth = 120;
            _pagesPanel.Controls.Add(_pageSelector);
            _pagesPanel.Controls.Add(_nextBtn);
            _pagesPanel.Dock = DockStyle.Fill;
            _bottomBar.Controls.Add(_pagesPanel);

            // Grid appearance
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.ReadOnly = true;
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
            _grid.CellDoubleClick += Grid_CellDoubleClick;
            _grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;
            _grid.DataBindingComplete += (_, __) => EnsureActionColumnsAtEnd();

            // Events
            Load += CadastroAlunosForm_Load;
            // Apply Windows 11-like effects (Mica/rounded) after handle is created
            this.HandleCreated += (_, __) =>
            {
                try { MedControl.UI.FluentEffects.ApplyWin11Mica(this); } catch { }
            };
        }

        private void CenterGrid()
        {
            var area = new Rectangle(0, 0, ClientSize.Width, Math.Max(0, ClientSize.Height - _bottomBar.Height));
            int x = Math.Max(0, (area.Width - _grid.Width) / 2);
            int y = Math.Max(0, (area.Height - _grid.Height) / 2);
            _grid.Location = new Point(x, y);
        }

        private static Button CreateButton(string text, Color back)
        {
            return new Button
            {
                Text = text,
                AutoSize = true,
                Padding = new Padding(8, 6, 8, 6),
                Margin = new Padding(6, 6, 6, 6),
                FlatStyle = FlatStyle.Flat,
                BackColor = back,
                ForeColor = Color.White
                ,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                TextAlign = ContentAlignment.MiddleCenter,
                MinimumSize = new Size(96, 34)
            };
        }

        // Create a small icon bitmap for common actions so we don't rely on external image files.
        private static Bitmap CreateIconBitmap(string kind, int width = 20, int height = 20)
        {
            var bmp = new Bitmap(width, height);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using var pen = new Pen(Color.White, Math.Max(2, width / 8));
            using var brush = new SolidBrush(Color.White);
            if (kind == "add")
            {
                // draw plus
                int cx = width / 2, cy = height / 2; int len = Math.Max(6, width / 3);
                g.DrawLine(pen, cx - len / 2, cy, cx + len / 2, cy);
                g.DrawLine(pen, cx, cy - len / 2, cx, cy + len / 2);
            }
            else if (kind == "import")
            {
                // down arrow
                var pts = new Point[] { new Point(width / 2, height - 2), new Point(2, height / 3), new Point(width - 2, height / 3) };
                g.FillPolygon(brush, pts);
                g.DrawLine(pen, width / 2, 2, width / 2, height / 3);
            }
            else if (kind == "export")
            {
                // up arrow
                var pts = new Point[] { new Point(width / 2, 2), new Point(2, height * 2 / 3), new Point(width - 2, height * 2 / 3) };
                g.FillPolygon(brush, pts);
                g.DrawLine(pen, width / 2, height - 2, width / 2, height * 2 / 3);
            }
            return bmp;
        }

        private void CadastroAlunosForm_Load(object? sender, EventArgs e)
        {
            // Inicializa DB e carrega alunos do banco
            try { MedControl.Database.Setup(); } catch { }

            try
            {
                _fullTable = MedControl.Database.GetAlunosAsDataTable();
            }
            catch
            {
                _fullTable = new DataTable();
            }

            // Se vazio, tenta pré-preencher a visualização com Excel configurado (não persiste até importar)
            if (_fullTable.Rows.Count == 0)
            {
                var path = Database.GetConfig("caminho_alunos");
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    try { _fullTable = ExcelHelper.LoadToDataTable(path); } catch { _fullTable = new DataTable(); }
                }
            }

            // Garante coluna _id como string (GUID)
            if (!_fullTable.Columns.Contains("_id"))
            {
                _fullTable.Columns.Add("_id", typeof(string));
                foreach (DataRow r in _fullTable.Rows) r["_id"] = Guid.NewGuid().ToString();
            }
            else
            {
                var colObj = _fullTable.Columns.Contains("_id") ? _fullTable.Columns["_id"] : null;
                if (colObj is DataColumn idCol && idCol.DataType != typeof(string))
                {
                    idCol.ColumnName = "_id_old";
                    _fullTable.Columns.Add("_id", typeof(string));
                    foreach (DataRow r in _fullTable.Rows)
                    {
                        var v = r.Table.Columns.Contains("_id_old") ? r["_id_old"] : null;
                        r["_id"] = v == null || v == DBNull.Value ? Guid.NewGuid().ToString() : v.ToString();
                    }
                    _fullTable.Columns.Remove("_id_old");
                }
            }

            if (_fullTable.Columns.Count == 1)
            {
                _fullTable.Columns.Add("Nome");
            }

            _filteredTable = _fullTable.Copy();
            _currentPage = 1;
            ApplyFilterAndRefresh();
            CenterGrid();

            // Add import/export/add controls to top panel
            _topPanel.Controls.Clear();
            _searchBox.Margin = new Padding(8, 10, 8, 8);
            // action buttons will appear to the left of the search box per user preference
            _importBtn.Click += ImportBtn_Click;
            _exportBtn.Click += ExportBtn_Click;
            _addBtn.Click += (_, __) => AddNewRow();
            // smaller margins so buttons sit closer
            _importBtn.Margin = new Padding(4, 6, 4, 6);
            _exportBtn.Margin = new Padding(4, 6, 4, 6);
            _addBtn.Margin = new Padding(4, 6, 4, 6);
            // Make all three action buttons identical size and with emoji + text labels
            var uniformSize = new Size(150, 44);
            foreach (var b in new[] { _addBtn, _importBtn, _exportBtn })
            {
                b.AutoSize = false;
                b.Size = uniformSize;
                b.MinimumSize = uniformSize;
                b.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
                b.ImageAlign = ContentAlignment.MiddleLeft;
                b.TextImageRelation = TextImageRelation.ImageBeforeText;
                b.FlatStyle = FlatStyle.Flat;
                b.FlatAppearance.BorderSize = 0;
            }
            // Set emoji + text labels and apply hover style (buttons stay square)
            ApplyButtonStyle(_addBtn, "➕ Adicionar", Color.FromArgb(0, 123, 255));
            ApplyButtonStyle(_importBtn, "⬇️ Upload", Color.FromArgb(33, 150, 243));
            ApplyButtonStyle(_exportBtn, "📤 Exportar", Color.FromArgb(76, 175, 80));
            _buttonToolTip.SetToolTip(_addBtn, "Adicionar novo registro");
            _buttonToolTip.SetToolTip(_importBtn, "Importar planilha (pré-preenchimento)");
            _buttonToolTip.SetToolTip(_exportBtn, "Exportar para Excel");

            // place action buttons directly to the left of the search box with small spacing
            var actions = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(4), Margin = new Padding(0) };
            actions.Controls.Add(_addBtn);
            actions.Controls.Add(_importBtn);
            actions.Controls.Add(_exportBtn);
            // vertically align actions with search box (same top/bottom margin as buttons)
            actions.Margin = new Padding(0, 6, 0, 6);
            // add actions first so they appear to the left of search
            _topPanel.Controls.Add(actions);
            // Style search box: same height and vertical margins as buttons for alignment
            _searchBox.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            _searchBox.Height = 44;
            _searchBox.Margin = new Padding(12, 6, 8, 6);
            _searchBox.BorderStyle = BorderStyle.FixedSingle;
            this.Load += (_, __) => Roundify(_searchBox, 12);
            _searchBox.SizeChanged += (_, __) => Roundify(_searchBox, 12);
            _topPanel.Controls.Add(_searchBox);
        }

        private static void Roundify(Control control, int radius)
        {
            if (control.Width <= 0 || control.Height <= 0) return;
            int diameter = radius * 2;
            Rectangle bounds = new Rectangle(0, 0, control.Width, control.Height);
            using (GraphicsPath path = new GraphicsPath())
            {
                path.StartFigure();
                path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
                path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
                path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                control.Region = new Region(path);
            }
        }

        private void ApplyButtonStyle(Button btn, string text, Color baseColor)
        {
            btn.Text = text;
            btn.BackColor = baseColor;
            btn.ForeColor = Color.White;
            btn.Cursor = Cursors.Hand;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;

            // Keep buttons square as requested (no rounded region applied)

            // Hover/press visual feedback
            Color hover = Darken(baseColor, 12);
            Color pressed = Darken(baseColor, 22);
            btn.MouseEnter += (_, __) => btn.BackColor = hover;
            btn.MouseLeave += (_, __) => btn.BackColor = baseColor;
            btn.MouseDown += (_, __) => btn.BackColor = pressed;
            btn.MouseUp += (_, __) => btn.BackColor = hover;
        }

        private static Color Darken(Color color, int amount)
        {
            int r = Math.Max(0, color.R - amount);
            int g = Math.Max(0, color.G - amount);
            int b = Math.Max(0, color.B - amount);
            return Color.FromArgb(color.A, r, g, b);
        }

        private void ApplyFilterAndRefresh()
        {
            var term = _searchBox.Text?.Trim();
            if (string.IsNullOrEmpty(term))
            {
                _filteredTable = _fullTable.Copy();
            }
            else
            {
                // Simple global search across string-convertible columns
                var result = _fullTable.Clone();
                foreach (DataRow r in _fullTable.Rows)
                {
                    bool match = false;
                    foreach (DataColumn c in _fullTable.Columns)
                    {
                        var val = r[c]?.ToString() ?? string.Empty;
                        if (val.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) { match = true; break; }
                    }
                    if (match) result.ImportRow(r);
                }
                _filteredTable = result;
            }

            ApplySort();
            _currentPage = Math.Min(Math.Max(1, _currentPage), TotalPages);
            RefreshGrid();
        }

        private void ApplySort()
        {
            if (_filteredTable == null) return;
            if (string.IsNullOrEmpty(_lastSortColumn)) return;
            try
            {
                var dv = _filteredTable.DefaultView;
                dv.Sort = _lastSortColumn + (_sortAscending ? " ASC" : " DESC");
                _filteredTable = dv.ToTable();
            }
            catch
            {
                // ignore sort errors
            }
        }

        private void Grid_ColumnHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0) return;
            var col = _grid.Columns[e.ColumnIndex];
            var colName = col.DataPropertyName ?? col.Name ?? col.HeaderText;
            if (string.IsNullOrEmpty(colName)) return;
            // ignore action columns
            if (colName.StartsWith("__") || colName == "_id") return;
            if (_lastSortColumn == colName) _sortAscending = !_sortAscending; else { _lastSortColumn = colName; _sortAscending = true; }
            _currentPage = 1;
            ApplyFilterAndRefresh();
        }

        private int TotalPages => Math.Max(1, (int)Math.Ceiling((_filteredTable?.Rows.Count ?? 0) / (double)_pageSize));

        private void RefreshGrid()
        {
            if (_filteredTable == null) _filteredTable = _fullTable.Copy();
            var dt = _filteredTable.Clone();
            int start = (_currentPage - 1) * _pageSize;
            for (int i = start; i < Math.Min(start + _pageSize, _filteredTable.Rows.Count); i++) dt.ImportRow(_filteredTable.Rows[i]);

            // add Edit/Delete buttons as last columns if not present
            var bindTable = dt.Copy();
            // REMOVE __actions column if present
            if (bindTable.Columns.Contains("__actions")) bindTable.Columns.Remove("__actions");

            _grid.DataSource = bindTable;

            // hide internal id column for cleaner view
            if (_grid.Columns.Contains("_id")) _grid.Columns["_id"].Visible = false;

            UpdatePaginationControls();
            EnsureButtonsColumn();
            EnsureActionColumnsAtEnd();
        }

        private void EnsureButtonsColumn()
        {
            // Remove existing action buttons column if present
            if (_grid.Columns.Contains("EDITAR")) return; // already wired

            // Add Edit/Delete as button columns with new names
            var editCol = new DataGridViewButtonColumn { Name = "EDITAR", Text = "✏️ Editar", UseColumnTextForButtonValue = true, Width = 80, SortMode = DataGridViewColumnSortMode.NotSortable, FlatStyle = FlatStyle.Flat };
            var azul = Color.FromArgb(0, 123, 255);
            var vermelho = Color.FromArgb(220, 53, 69);
            editCol.DefaultCellStyle.BackColor = azul;
            editCol.DefaultCellStyle.ForeColor = Color.White;
            editCol.DefaultCellStyle.SelectionBackColor = azul;
            editCol.DefaultCellStyle.SelectionForeColor = Color.White;
            var delCol = new DataGridViewButtonColumn { Name = "EXCLUIR", Text = "🗑️ Excluir", UseColumnTextForButtonValue = true, Width = 80, SortMode = DataGridViewColumnSortMode.NotSortable, FlatStyle = FlatStyle.Flat };
            delCol.DefaultCellStyle.BackColor = vermelho;
            delCol.DefaultCellStyle.ForeColor = Color.White;
            delCol.DefaultCellStyle.SelectionBackColor = vermelho;
            delCol.DefaultCellStyle.SelectionForeColor = Color.White;
            _grid.Columns.Add(editCol);
            _grid.Columns.Add(delCol);
            _grid.CellClick += Grid_CellClick;
            HookActionColumnColors();
            EnsureActionColumnsAtEnd();
        }

        private void HookActionColumnColors()
        {
            if (_actionColorsHooked) return;
            _actionColorsHooked = true;
            _grid.CellFormatting += (s, e) =>
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
            };
        }

        private void EnsureActionColumnsAtEnd()
        {
            try
            {
                if (_grid.Columns.Contains("EDITAR") && _grid.Columns.Contains("EXCLUIR"))
                {
                    // Push action columns to the last positions
                    var lastIndex = _grid.Columns.Count - 1;
                    _grid.Columns["EXCLUIR"].DisplayIndex = lastIndex;
                    _grid.Columns["EDITAR"].DisplayIndex = Math.Max(0, lastIndex - 1);
                }
            }
            catch { }
        }

        private void Grid_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            // If Ctrl is held during double-click, copy the cell value to clipboard
            var clickedCell = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
            try
            {
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    var txt = clickedCell.Value?.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(txt))
                    {
                        Clipboard.SetText(txt);
                        MessageBox.Show(this, "Valor copiado para a área de transferência.", "Copiado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    return;
                }
            }
            catch
            {
                // clipboard can fail silently in some environments; ignore
            }

            // Otherwise, open a single-field editor for the clicked column (do not open full multi-field editor)
            var col = _grid.Columns[e.ColumnIndex];
            var colName = col.DataPropertyName ?? col.Name ?? col.HeaderText;
            if (string.IsNullOrEmpty(colName)) return;
            // ignore action/internal columns
            if (colName.StartsWith("__") || string.Equals(colName, "_id", StringComparison.OrdinalIgnoreCase)) return;

            var row = GetDataRowFromGridRow(e.RowIndex);
            if (row == null) return;

            var current = row[colName]?.ToString() ?? string.Empty;
            var input = Prompt.ShowDialog("Editar " + colName, current);
            if (input != null)
            {
                var id = row["_id"].ToString();
                var orig = _fullTable.AsEnumerable().FirstOrDefault(r => r["_id"].ToString() == id);
                if (orig != null)
                {
                    orig[colName] = input;
                    _ = SaveRowToDatabaseAsync(orig);
                    ApplyFilterAndRefresh();
                }
            }
        }

        private void Grid_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_grid.Columns[e.ColumnIndex].Name == "EDITAR")
            {
                EditRowAtGridIndex(e.RowIndex);
            }
            else if (_grid.Columns[e.ColumnIndex].Name == "EXCLUIR")
            {
                var res = MessageBox.Show(this, "Excluir este registro?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes) DeleteRowAtGridIndex(e.RowIndex);
            }
        }

        private void EditRowAtGridIndex(int gridRowIndex)
        {
            var row = GetDataRowFromGridRow(gridRowIndex);
            if (row == null) return;
            // Ensure there's at least one editable column (not counting internal _id)
            if (!_fullTable.Columns.Cast<DataColumn>().Any(c => !string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)))
            {
                _fullTable.Columns.Add("Nome");
            }

            // Show multi-field editor and update all fields
            var values = AddEditForm.ShowDialog("Editar registro", _fullTable.Columns, row);
            if (values != null)
            {
                var id = row["_id"].ToString();
                var orig = _fullTable.AsEnumerable().FirstOrDefault(r => r["_id"].ToString() == id);
                if (orig != null)
                {
                    foreach (var kv in values)
                    {
                        if (string.Equals(kv.Key, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                        if (_fullTable.Columns.Contains(kv.Key)) orig[kv.Key] = kv.Value;
                    }
                }
                _ = SaveRowToDatabaseAsync(orig!);
                ApplyFilterAndRefresh();
            }
        }

        private void DeleteRowAtGridIndex(int gridRowIndex)
        {
            var row = GetDataRowFromGridRow(gridRowIndex);
            if (row == null) return;
            var id = row["_id"].ToString();
            var orig = _fullTable.AsEnumerable().FirstOrDefault(r => r["_id"].ToString() == id);
            if (orig != null)
            {
                _fullTable.Rows.Remove(orig);
                _ = DeleteRowFromDatabaseAsync(id!);
                ApplyFilterAndRefresh();
            }
        }

        private DataRow? GetDataRowFromGridRow(int gridRowIndex)
        {
            if (gridRowIndex < 0 || gridRowIndex >= _grid.Rows.Count) return null;
            var gridRow = _grid.Rows[gridRowIndex];
            // try to find the column index that represents the _id data property (works for bound and generated columns)
            int? idColIndex = null;
            foreach (DataGridViewColumn col in _grid.Columns)
            {
                if (string.Equals(col.DataPropertyName, "_id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.Name, "_id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col.HeaderText, "_id", StringComparison.OrdinalIgnoreCase))
                {
                    idColIndex = col.Index;
                    break;
                }
            }

            if (idColIndex == null)
            {
                // fallback: try to find a cell whose value looks like a GUID
                foreach (DataGridViewCell c in gridRow.Cells)
                {
                    if (c.Value is Guid) { idColIndex = c.ColumnIndex; break; }
                    if (c.Value is string s)
                    {
                        if (Guid.TryParse(s, out _)) { idColIndex = c.ColumnIndex; break; }
                    }
                }
            }

            if (idColIndex != null)
            {
                var val = gridRow.Cells[idColIndex.Value].Value;
                if (val != null)
                {
                    var id = val.ToString() ?? string.Empty;
                    return _fullTable.AsEnumerable().FirstOrDefault(r => r["_id"].ToString() == id);
                }
            }

            return null;
        }

        private void AddNewRow()
        {
            // Ensure there's at least one editable column (not counting internal _id)
            if (!_fullTable.Columns.Cast<DataColumn>().Any(c => !string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)))
            {
                _fullTable.Columns.Add("Nome");
            }

            // Show multi-field editor to gather values for all fields
            var values = AddEditForm.ShowDialog("Adicionar registro", _fullTable.Columns, null);
            if (values != null)
            {
                // Basic validation: if there's a 'Nome' column, require it non-empty
                if (_fullTable.Columns.Contains("Nome"))
                {
                    while (string.IsNullOrWhiteSpace(values.GetValueOrDefault("Nome")))
                    {
                        MessageBox.Show(this, "O campo 'Nome' é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        values = MultiFieldPrompt.ShowDialog("Adicionar registro", _fullTable.Columns, null);
                        if (values == null) return; // user cancelled
                    }
                }

                var nr = _fullTable.NewRow();
                foreach (DataColumn c in _fullTable.Columns)
                {
                    if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) nr["_id"] = Guid.NewGuid().ToString();
                    else
                    {
                        if (values.TryGetValue(c.ColumnName, out var v)) nr[c.ColumnName] = v;
                        else nr[c.ColumnName] = string.Empty;
                    }
                }
                _fullTable.Rows.Add(nr);
                _ = SaveRowToDatabaseAsync(nr);
                ApplyFilterAndRefresh();
                MessageBox.Show(this, "Registro adicionado.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // ===== Helpers: persist row changes asynchronously to Database =====
        private Dictionary<string, string> BuildDictFromRow(DataRow row)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn c in row.Table.Columns)
            {
                if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                dict[c.ColumnName] = Convert.ToString(row[c]) ?? string.Empty;
            }
            return dict;
        }

        private Task SaveRowToDatabaseAsync(DataRow row)
        {
            try
            {
                var uid = Convert.ToString(row["_id"]) ?? Guid.NewGuid().ToString();
                var data = BuildDictFromRow(row);
                // Grava de forma síncrona para garantir persistência antes de fechar a tela/app
                Database.Setup();
                Database.UpsertAluno(uid, data);
                // Recarrega da base para refletir o estado canônico
                ReloadFromDatabase();
            }
            catch (Exception ex)
            {
                try { MessageBox.Show(this, "Falha ao salvar no banco: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
            return Task.CompletedTask;
        }

        private Task DeleteRowFromDatabaseAsync(string uid)
        {
            try
            {
                Database.Setup();
                Database.DeleteAluno(uid);
                ReloadFromDatabase();
            }
            catch (Exception ex)
            {
                try { MessageBox.Show(this, "Falha ao excluir no banco: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
            return Task.CompletedTask;
        }

        private void ReloadFromDatabase()
        {
            try
            {
                var fresh = Database.GetAlunosAsDataTable();
                _fullTable = fresh;
                ApplyFilterAndRefresh();
            }
            catch { }
        }

        private void UpdatePaginationControls()
        {
            _pagesPanel.SuspendLayout();
            _pagesPanel.Controls.Clear();
            _pagesPanel.Controls.Add(_prevBtn);

            int total = TotalPages;
            // show up to 5 page buttons around current
            int show = 5;
            int start = Math.Max(1, _currentPage - 2);
            int end = Math.Min(total, start + show - 1);
            if (end - start + 1 < show) start = Math.Max(1, end - show + 1);

            // If the window doesn't start at 1, show a first-page button and optionally an ellipsis
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
                        // populate selector with earlier pages
                        _pageSelector.Items.Clear();
                        for (int p = 1; p < start; p++) _pageSelector.Items.Add(p);
                        if (_pageSelector.Items.Count > 0)
                        {
                            _pageSelector.SelectedIndex = 0;
                            _pageSelector.Visible = true;
                            _pageSelector.DroppedDown = true;
                            _pageSelector.Focus();
                        }
                    };
                    _pagesPanel.Controls.Add(leftEll);
                }
            }

            // Main page buttons
            for (int p = start; p <= end; p++)
            {
                var b = new Button { Text = p.ToString(), AutoSize = true, Padding = new Padding(6), Margin = new Padding(4) };
                b.BackColor = p == _currentPage ? Color.FromArgb(0, 123, 255) : Color.White;
                b.ForeColor = p == _currentPage ? Color.White : Color.Black;
                int page = p;
                b.Click += (_, __) => { _currentPage = page; _pageSelector.Visible = false; RefreshGrid(); };
                _pagesPanel.Controls.Add(b);
            }

            // If window doesn't end at total, optionally show ellipsis and the last page button
            if (end < total)
            {
                if (end < total - 1)
                {
                    var rightEll = new LinkLabel { Text = "...", AutoSize = true, LinkColor = Color.Black, ActiveLinkColor = Color.DimGray, Padding = new Padding(10) };
                    rightEll.LinkClicked += (_, __) =>
                    {
                        // populate selector with later pages
                        _pageSelector.Items.Clear();
                        for (int p = end + 1; p <= total; p++) _pageSelector.Items.Add(p);
                        if (_pageSelector.Items.Count > 0)
                        {
                            _pageSelector.SelectedIndex = 0;
                            _pageSelector.Visible = true;
                            _pageSelector.DroppedDown = true;
                            _pageSelector.Focus();
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
            else
            {
                _pageSelector.Visible = false;
            }

            _pagesPanel.Controls.Add(_nextBtn);
            _pagesPanel.ResumeLayout();
        }

        private void ImportBtn_Click(object? sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                try
                {
                    UseWaitCursor = true;
                    Enabled = false;
                    _fullTable = ExcelHelper.LoadToDataTable(ofd.FileName);
                    if (!_fullTable.Columns.Contains("_id"))
                    {
                        _fullTable.Columns.Add("_id", typeof(string));
                        foreach (DataRow r in _fullTable.Rows) r["_id"] = Guid.NewGuid().ToString();
                    }
                    else
                    {
                        var idCol = _fullTable.Columns["_id"];
                        if (idCol != null && idCol.DataType != typeof(string))
                        {
                            idCol.ColumnName = "_id_old";
                            _fullTable.Columns.Add("_id", typeof(string));
                            foreach (DataRow r in _fullTable.Rows)
                            {
                                var v = r.Table.Columns.Contains("_id_old") ? r["_id_old"] : null;
                                r["_id"] = v == null || v == DBNull.Value ? Guid.NewGuid().ToString() : v.ToString();
                            }
                            _fullTable.Columns.Remove("_id_old");
                        }
                    }

                    ApplyFilterAndRefresh();

                    _ = Task.Run(() => Database.ReplaceAllAlunos(_fullTable))
                        .ContinueWith(_ =>
                        {
                            try
                            {
                                var fresh = Database.GetAlunosAsDataTable();
                                BeginInvoke(new Action(() =>
                                {
                                    _fullTable = fresh;
                                    ApplyFilterAndRefresh();
                                }));
                            }
                            catch { }
                        })
                        .ContinueWith(_ => BeginInvoke(new Action(() => { Enabled = true; UseWaitCursor = false; })));
                }
                catch (Exception ex)
                {
                    Enabled = true; UseWaitCursor = false;
                    MessageBox.Show(this, "Falha ao importar: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ExportBtn_Click(object? sender, EventArgs e)
        {
            using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "cadastro_alunos.xlsx" };
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                ExcelHelper.SaveDataTable(sfd.FileName, _fullTable);
                MessageBox.Show("Exportado com sucesso.");
            }
        }
    }

    // Minimal prompt helper dialog for simple text entry
    internal static class Prompt
    {
        public static string? ShowDialog(string title, string defaultText)
        {
            using var form = new Form();
            form.Text = title; form.StartPosition = FormStartPosition.CenterParent; form.Width = 400; form.Height = 150; form.FormBorderStyle = FormBorderStyle.FixedDialog; form.MinimizeBox = false; form.MaximizeBox = false;
            var tb = new TextBox { Dock = DockStyle.Top, Text = defaultText, Margin = new Padding(10) };
            var panel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, FlowDirection = FlowDirection.RightToLeft };
            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Padding = new Padding(8) };
            var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel, AutoSize = true, Padding = new Padding(8) };
            panel.Controls.Add(ok); panel.Controls.Add(cancel);
            form.Controls.Add(tb); form.Controls.Add(panel);
            form.AcceptButton = ok; form.CancelButton = cancel;
            return form.ShowDialog() == DialogResult.OK ? tb.Text : null;
        }
    }

    // Dialog to edit/add multiple fields at once (all columns except internal _id)
    internal static class MultiFieldPrompt
    {
        // Returns null if canceled, or a dictionary mapping column name -> value
        public static Dictionary<string, string>? ShowDialog(string title, DataColumnCollection columns, DataRow? existing = null)
        {
            using var form = new Form();
            form.Text = title;
            form.StartPosition = FormStartPosition.CenterParent;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MinimizeBox = false; form.MaximizeBox = false;
            form.AutoSize = true; form.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Padding = new Padding(10), Dock = DockStyle.Fill, GrowStyle = TableLayoutPanelGrowStyle.AddRows };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var inputs = new Dictionary<string, TextBox>();
            int row = 0;
            foreach (DataColumn col in columns)
            {
                if (string.Equals(col.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                table.RowCount = row + 1;
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var lbl = new Label { Text = col.ColumnName, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 8, 3, 3) };
                var tb = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.None, Multiline = false, Height = 26, Margin = new Padding(3, 6, 3, 6) };
                tb.MinimumSize = new Size(240, 26);
                if (existing != null && existing.Table.Columns.Contains(col.ColumnName))
                {
                    var val = existing[col.ColumnName];
                    tb.Text = val?.ToString() ?? string.Empty;
                }

                table.Controls.Add(lbl, 0, row);
                table.Controls.Add(tb, 1, row);
                inputs[col.ColumnName] = tb;
                row++;
            }

            var panel = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(8) };
            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, Padding = new Padding(8) };
            var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel, AutoSize = true, Padding = new Padding(8) };
            panel.Controls.Add(ok); panel.Controls.Add(cancel);

            // Add table first (fills the form), then the buttons panel at bottom
            var container = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            container.Controls.Add(table);
            form.Controls.Add(container);
            form.Controls.Add(panel);

            // Compute a sensible client size and set it before showing so textboxes get layouted
            var pref = table.GetPreferredSize(new Size(800, 0));
            var width = Math.Min(900, Math.Max(480, pref.Width + 80));
            var height = Math.Min(800, Math.Max(220, pref.Height + 140));
            form.ClientSize = new Size(width, height);
            form.AcceptButton = ok; form.CancelButton = cancel;

            var dr = form.ShowDialog();
            if (dr != DialogResult.OK) return null;

            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in inputs)
            {
                result[kv.Key] = kv.Value.Text ?? string.Empty;
            }
            return result;
        }
    }

    // New dedicated Add/Edit form with explicit Save and Cancel buttons
    internal class AddEditForm : Form
    {
        private readonly Dictionary<string, TextBox> _inputs = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> Results { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private AddEditForm(string title)
        {
            Text = title;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false; MaximizeBox = false;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
        }

        // Show dialog and return values or null if canceled
        public static Dictionary<string, string>? ShowDialog(string title, DataColumnCollection columns, DataRow? existing = null)
        {
            using var form = new AddEditForm(title);

            var layout = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 0,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Padding = new Padding(18, 18, 18, 8),
                BackColor = Color.White
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            int row = 0;
            foreach (DataColumn col in columns)
            {
                if (string.Equals(col.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                layout.RowCount = row + 1;
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var lbl = new Label
                {
                    Text = col.ColumnName,
                    AutoSize = true,
                    Anchor = AnchorStyles.Right,
                    Margin = new Padding(3, 8, 12, 8),
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
                };
                var tb = new TextBox
                {
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    Dock = DockStyle.Fill,
                    Height = 28,
                    Margin = new Padding(3, 8, 3, 8),
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
                };
                tb.MinimumSize = new Size(220, 28);
                if (existing != null && existing.Table.Columns.Contains(col.ColumnName))
                {
                    var val = existing[col.ColumnName];
                    tb.Text = val?.ToString() ?? string.Empty;
                }
                layout.Controls.Add(lbl, 0, row);
                layout.Controls.Add(tb, 1, row);
                form._inputs[col.ColumnName] = tb;
                row++;
            }

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                AutoSize = true,
                Padding = new Padding(8),
                BackColor = Color.White
            };
            var save = new Button { Text = "Salvar", AutoSize = true, Padding = new Padding(10, 6, 10, 6), Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            var cancel = new Button { Text = "Cancelar", AutoSize = true, Padding = new Padding(10, 6, 10, 6), Font = new Font("Segoe UI", 10F, FontStyle.Regular) };
            buttons.Controls.Add(save); buttons.Controls.Add(cancel);

            form.Controls.Add(layout);
            form.Controls.Add(buttons);

            // Ajusta tamanho inicial
            form.ClientSize = new Size(Math.Max(420, layout.PreferredSize.Width + 60), Math.Max(220, layout.PreferredSize.Height + 80));

            save.Click += (_, __) =>
            {
                foreach (var kv in form._inputs)
                    form.Results[kv.Key] = kv.Value.Text ?? string.Empty;
                form.DialogResult = DialogResult.OK;
                form.Close();
            };
            cancel.Click += (_, __) => { form.DialogResult = DialogResult.Cancel; form.Close(); };

            var dr = form.ShowDialog();
            return dr == DialogResult.OK ? form.Results : null;
        }
    }
}
