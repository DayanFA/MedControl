using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class CadastroChavesForm : Form
    {
    private readonly DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false };
    private readonly BindingSource _bs = new BindingSource();
    private bool _handlingCellAction = false; // evita duplo disparo
    private readonly System.Collections.Generic.Dictionary<int, string> _oldNameByRow = new System.Collections.Generic.Dictionary<int, string>();

    // Top bar & actions
    private readonly FlowLayoutPanel _top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 64, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(8) };
    private readonly TextBox _search = new TextBox { Width = 300, PlaceholderText = "Pesquisar... 🔎" };
    private readonly Button _btnAdd = new Button();
    private readonly Button _btnEdit = new Button();
    private readonly Button _btnDel = new Button();

    // Pagination
    private readonly Panel _bottomBar = new Panel { Dock = DockStyle.Bottom, Height = 48 };
    private readonly Button _prevBtn = new Button();
    private readonly Button _nextBtn = new Button();
    private readonly FlowLayoutPanel _pagesPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
    private readonly ComboBox _pageSelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70, Visible = false };
    private int _pageSize = 10;
    private int _currentPage = 1;
    private string? _lastSort = null;
    private bool _sortAsc = true;
    private List<Chave> _all = new List<Chave>();
    private List<Chave> _filtered = new List<Chave>();

        public CadastroChavesForm()
        {
            Text = "Cadastro de Chaves";
            Width = 800;
            Height = 600;

            // Top actions with emoji and square buttons
            ConfigureActionButton(_btnAdd, "➕ Adicionar", Color.FromArgb(0, 123, 255));
            // Remover botões de Editar/Excluir do topo (já existem no grid)
            _btnEdit.Visible = false;
            _btnDel.Visible = false;
            _btnAdd.Click += (_, __) => AddChave();
            var actions = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(4), Margin = new Padding(0, 6, 0, 6) };
            actions.Controls.AddRange(new Control[] { _btnAdd });
            _top.Controls.Add(actions);
            // Search aligned with buttons
            _search.Font = new Font("Segoe UI", 10F);
            _search.Height = 44;
            _search.Margin = new Padding(12, 6, 8, 6);
            _search.BorderStyle = BorderStyle.FixedSingle;
            _search.TextChanged += (_, __) => { _currentPage = 1; ApplyFilterAndRefresh(); };
            // Keep search box square (no rounded corners)
            _search.Region = null;
            _top.Controls.Add(_search);

            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.ReadOnly = false; // permitir edição inline
            _grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
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

            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nome", DataPropertyName = nameof(Chave.Nome), Name = nameof(Chave.Nome), Width = 200 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Número de Cópias", DataPropertyName = nameof(Chave.NumCopias), Name = nameof(Chave.NumCopias), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Descrição", DataPropertyName = nameof(Chave.Descricao), Name = nameof(Chave.Descricao), Width = 300 });
            EnsureButtonsColumn();
            _grid.CellContentClick += Grid_CellContentClick; // usar apenas este para botões
            _grid.CellBeginEdit += Grid_CellBeginEdit;       // capturar nome antigo para renomear
            _grid.CellEndEdit += Grid_CellEndEdit;           // persistir ao terminar edição
            _grid.CellValidating += Grid_CellValidating;     // validar NumCopias e Nome
            _grid.DataError += (s, e) => { e.Cancel = false; }; // evita crash em conversões

            Controls.Add(_grid);
            Controls.Add(_top);
            Controls.Add(_bottomBar);

            Load += (_, __) => { ReloadFromDb(); ApplyFilterAndRefresh(); try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { } };

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
        }

        private void RefreshGrid()
        {
            var pageItems = Paginate(_filtered, _currentPage, _pageSize);
            _bs.DataSource = pageItems;
            _grid.DataSource = _bs;
            EnsureActionColumnsAtEnd();
            UpdatePaginationControls();
        }

        private void AddChave()
        {
            var dlg = new ChaveDialog();
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                Database.UpsertChave(dlg.Result);
                ReloadFromDb();
                ApplyFilterAndRefresh();
            }
        }

        private void EditChave()
        {
            if (_bs.Current is Chave c)
            {
                var dlg = new ChaveDialog(c);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        Database.UpdateChave(c.Nome, dlg.Result);
                    }
                    catch (Exception ex)
                    {
                        try { MessageBox.Show(this, "Não foi possível salvar a edição: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                    }
                    ReloadFromDb();
                    ApplyFilterAndRefresh();
                }
            }
        }

        private void DeleteChave()
        {
            if (_bs.Current is Chave c)
            {
                if (MessageBox.Show($"Deletar a chave '{c.Nome}'?", "Deletar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Database.DeleteChave(c.Nome);
                    ReloadFromDb();
                    ApplyFilterAndRefresh();
                }
            }
        }

        private void ReloadFromDb()
        {
            _all = Database.GetChaves() ?? new List<Chave>();
        }

        private void ApplyFilterAndRefresh()
        {
            var term = _search.Text?.Trim();
            IEnumerable<Chave> q = _all;
            if (!string.IsNullOrEmpty(term))
            {
                q = q.Where(c =>
                    (c.Nome ?? string.Empty).IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (c.Descricao ?? string.Empty).IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    c.NumCopias.ToString().IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0
                );
            }

            // sorting
            if (!string.IsNullOrEmpty(_lastSort))
            {
                q = _lastSort switch
                {
                    nameof(Chave.Nome) => (_sortAsc ? q.OrderBy(c => c.Nome) : q.OrderByDescending(c => c.Nome)),
                    nameof(Chave.NumCopias) => (_sortAsc ? q.OrderBy(c => c.NumCopias) : q.OrderByDescending(c => c.NumCopias)),
                    nameof(Chave.Descricao) => (_sortAsc ? q.OrderBy(c => c.Descricao) : q.OrderByDescending(c => c.Descricao)),
                    _ => q
                };
            }

            _filtered = q.ToList();
            _currentPage = Math.Min(Math.Max(1, _currentPage), TotalPages);
            RefreshGrid();
        }

        private List<Chave> Paginate(List<Chave> list, int page, int size)
        {
            int start = (page - 1) * size;
            if (start < 0) start = 0;
            return list.Skip(start).Take(size).ToList();
        }

        private int TotalPages => Math.Max(1, (int)Math.Ceiling((_filtered?.Count ?? 0) / (double)_pageSize));

        private void OnHeaderClick(int columnIndex)
        {
            if (columnIndex < 0) return;
            var col = _grid.Columns[columnIndex];
            var colName = col.DataPropertyName ?? col.Name ?? col.HeaderText;
            if (string.IsNullOrEmpty(colName)) return;
            if (colName == nameof(Chave.Nome) || colName == nameof(Chave.NumCopias) || colName == nameof(Chave.Descricao))
            {
                _sortAsc = _lastSort == colName ? !_sortAsc : true;
                _lastSort = colName;
                ApplyFilterAndRefresh();
            }
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

        private void ConfigureActionButton(Button btn, string text, Color baseColor)
        {
            var size = new Size(150, 44);
            btn.Text = text;
            btn.AutoSize = false;
            btn.Size = size;
            btn.MinimumSize = size;
            btn.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            // Marca como botão de ação/acento para o ThemeHelper não sobrescrever cores/estilo
            btn.Tag = "accent";
            btn.ImageAlign = ContentAlignment.MiddleLeft;
            btn.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Cursor = Cursors.Hand;
            btn.BackColor = baseColor;
            btn.ForeColor = Color.White;
            var baseCol = btn.BackColor;
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

        private void EnsureButtonsColumn()
        {
            if (_grid.Columns.Contains("EDITAR") && _grid.Columns.Contains("EXCLUIR")) return;
            var editCol = new DataGridViewButtonColumn
            {
                Name = "EDITAR",
                Text = "✏️ Editar",
                UseColumnTextForButtonValue = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FlatStyle = FlatStyle.Flat,
                ReadOnly = true
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
                FlatStyle = FlatStyle.Flat,
                ReadOnly = true
            };
            delCol.DefaultCellStyle.BackColor = Color.FromArgb(220, 53, 69);
            delCol.DefaultCellStyle.ForeColor = Color.White;
            delCol.DefaultCellStyle.SelectionBackColor = Color.FromArgb(220, 53, 69);
            delCol.DefaultCellStyle.SelectionForeColor = Color.White;

            _grid.Columns.Add(editCol);
            _grid.Columns.Add(delCol);
            _grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
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
                    var lastIndex = _grid.Columns.Count - 1;
                    _grid.Columns["EXCLUIR"].DisplayIndex = lastIndex;
                    _grid.Columns["EDITAR"].DisplayIndex = Math.Max(0, lastIndex - 1);
                }
            }
            catch { }
        }

        private void Grid_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (_handlingCellAction) return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var colName = _grid.Columns[e.ColumnIndex].Name;
            if (colName != "EDITAR" && colName != "EXCLUIR") return;
            try
            {
                _handlingCellAction = true;
                var item = _grid.Rows[e.RowIndex].DataBoundItem as Chave;
                if (item == null) return;
                if (colName == "EXCLUIR")
                {
                    if (MessageBox.Show($"Deletar a chave '{item.Nome}'?", "Deletar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Database.DeleteChave(item.Nome);
                        ReloadFromDb();
                        ApplyFilterAndRefresh();
                    }
                }
                else if (colName == "EDITAR")
                {
                    var oldName = item.Nome;
                    var dlg = new ChaveDialog(item);
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        try
                        {
                            Database.UpdateChave(oldName, dlg.Result);
                        }
                        catch (Exception ex)
                        {
                            try { MessageBox.Show(this, "Não foi possível salvar a edição: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                        }
                        ReloadFromDb();
                        ApplyFilterAndRefresh();
                    }
                }
            }
            finally
            {
                _handlingCellAction = false;
            }
        }

        

        private void Grid_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var item = _grid.Rows[e.RowIndex].DataBoundItem as Chave;
            if (item == null) return;
            _oldNameByRow[e.RowIndex] = item.Nome;
        }

        private void Grid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var item = _grid.Rows[e.RowIndex].DataBoundItem as Chave;
            if (item == null) return;
            var oldName = _oldNameByRow.ContainsKey(e.RowIndex) ? _oldNameByRow[e.RowIndex] : item.Nome;

            try
            {
                // Se Nome mudou, renomeia; senão só atualiza os demais campos
                Database.UpdateChave(oldName, item);
                _oldNameByRow.Remove(e.RowIndex);
                ReloadFromDb();
                // mantém página atual e filtro
                ApplyFilterAndRefresh();
            }
            catch (Exception ex)
            {
                try { MessageBox.Show(this, "Falha ao salvar a edição: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
        }

        private void Grid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var col = _grid.Columns[e.ColumnIndex];
            if (col.Name == nameof(Chave.Nome))
            {
                var newVal = (e.FormattedValue?.ToString() ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(newVal))
                {
                    e.Cancel = true;
                    try { MessageBox.Show(this, "O campo 'Nome' é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                }
            }
            else if (col.Name == nameof(Chave.NumCopias))
            {
                var txt = e.FormattedValue?.ToString() ?? string.Empty;
                if (!int.TryParse(txt, out _))
                {
                    e.Cancel = true;
                    try { MessageBox.Show(this, "'Número de Cópias' deve ser um número inteiro.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                }
            }
        }

        private class ChaveDialog : Form
        {
            private TextBox _nome = new TextBox();
            private NumericUpDown _num = new NumericUpDown { Minimum = 0, Maximum = 10000 };
            private TextBox _desc = new TextBox();
            public Chave Result { get; private set; } = new Chave();

            public ChaveDialog(Chave? c = null)
            {
                Text = c == null ? "Adicionar Chave" : "Editar Chave";
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                AutoSize = true;
                AutoSizeMode = AutoSizeMode.GrowAndShrink;
                BackColor = Color.White;

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

                // Nome
                layout.RowCount++;
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var lblNome = new Label
                {
                    Text = "Nome",
                    AutoSize = true,
                    Anchor = AnchorStyles.Right,
                    Margin = new Padding(3, 8, 12, 8),
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
                };
                _nome.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
                _nome.Margin = new Padding(3, 8, 3, 8);
                _nome.MinimumSize = new Size(240, 28);
                _nome.Dock = DockStyle.Fill;
                layout.Controls.Add(lblNome, 0, layout.RowCount - 1);
                layout.Controls.Add(_nome, 1, layout.RowCount - 1);

                // Número de Cópias
                layout.RowCount++;
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var lblNum = new Label
                {
                    Text = "Número de Cópias",
                    AutoSize = true,
                    Anchor = AnchorStyles.Right,
                    Margin = new Padding(3, 8, 12, 8),
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
                };
                _num.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
                _num.Margin = new Padding(3, 8, 3, 8);
                _num.MinimumSize = new Size(120, 28);
                _num.Maximum = 100000;
                _num.Dock = DockStyle.Left;
                layout.Controls.Add(lblNum, 0, layout.RowCount - 1);
                layout.Controls.Add(_num, 1, layout.RowCount - 1);

                // Descrição
                layout.RowCount++;
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var lblDesc = new Label
                {
                    Text = "Descrição",
                    AutoSize = true,
                    Anchor = AnchorStyles.Right,
                    Margin = new Padding(3, 8, 12, 8),
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
                };
                _desc.Multiline = true;
                _desc.Dock = DockStyle.Fill;
                _desc.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
                _desc.Margin = new Padding(3, 8, 3, 8);
                _desc.MinimumSize = new Size(240, 80);
                layout.Controls.Add(lblDesc, 0, layout.RowCount - 1);
                layout.Controls.Add(_desc, 1, layout.RowCount - 1);

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
                buttons.Controls.Add(save);
                buttons.Controls.Add(cancel);

                Controls.Add(layout);
                Controls.Add(buttons);

                if (c != null)
                {
                    _nome.Text = c.Nome;
                    _num.Value = c.NumCopias;
                    _desc.Text = c.Descricao;
                }

                // Tamanho inicial baseado no conteúdo
                var width = Math.Max(420, layout.PreferredSize.Width + 60);
                var height = Math.Max(260, layout.PreferredSize.Height + buttons.PreferredSize.Height + 60);
                ClientSize = new Size(Math.Min(900, width), Math.Min(800, height));

                AcceptButton = save;
                CancelButton = cancel;

                save.Click += (_, __) =>
                {
                    var nome = _nome.Text.Trim();
                    if (string.IsNullOrWhiteSpace(nome))
                    {
                        MessageBox.Show(this, "O campo 'Nome' é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    Result = new Chave { Nome = nome, NumCopias = (int)_num.Value, Descricao = _desc.Text };
                    DialogResult = DialogResult.OK;
                    Close();
                };
                cancel.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            }
        }
    }
}
