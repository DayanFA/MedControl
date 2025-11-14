using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class EntregarChaveForm : Form
    {
    private ComboBox _chave = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
    private ComboBox _aluno = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
    private ComboBox _prof = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
    private CheckBox _outros = new CheckBox { Text = "Outros" };
    private TextBox _quem = new TextBox();
    private TextBox _contato = new TextBox();
    private CheckBox _itemEmprestado = new CheckBox { Text = "Anexar Termo" };
        private TextBox _termo = new TextBox { ReadOnly = true };
        private Button _btnTermo = new Button { Text = "Importar Termo" };
        private Button _btn = new Button { Text = "Entregar" };
        private Button _verEntregas = new Button { Text = "Relação" };

        public EntregarChaveForm()
        {
            Text = "Entregar Chave";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.DoubleBuffered = true;

            SuspendLayout();
            var layout = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Padding = new Padding(18, 18, 18, 8),
                BackColor = Color.White
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label MkLabel(string text) => new Label
            {
                Text = text,
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(3, 8, 12, 8),
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
            };

            var inputFont = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            foreach (var c in new Control[] { _chave, _aluno, _prof, _termo, _quem, _contato })
            {
                c.Font = inputFont;
                c.Margin = new Padding(3, 8, 3, 8);
                c.MinimumSize = new Size(260, 30);
                c.Dock = DockStyle.Fill;
            }
            _outros.Font = inputFont; _outros.Margin = new Padding(3,8,3,8);
            _chave.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _chave.AutoCompleteSource = AutoCompleteSource.ListItems;
            _aluno.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _aluno.AutoCompleteSource = AutoCompleteSource.ListItems;
            _prof.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _prof.AutoCompleteSource = AutoCompleteSource.ListItems;

            int row = 0;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Chave"), 0, row);
            layout.Controls.Add(_chave, 1, row++);

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Aluno"), 0, row);
            layout.Controls.Add(_aluno, 1, row++);

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Professor"), 0, row);
            layout.Controls.Add(_prof, 1, row++);

            // Checkbox Outros
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = string.Empty, AutoSize = true }, 0, row);
            layout.Controls.Add(_outros, 1, row++);

            // Campos Quem e Contato (inicialmente ocultos)
            var _lblQuem = MkLabel("Quem");
            var _lblContato = MkLabel("Contato");
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_lblQuem, 0, row);
            layout.Controls.Add(_quem, 1, row++);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_lblContato, 0, row);
            layout.Controls.Add(_contato, 1, row++);
            Action toggleOutros = () =>
            {
                var on = _outros.Checked;
                _lblQuem.Visible = on; _quem.Visible = on;
                _lblContato.Visible = on; _contato.Visible = on;
                _aluno.Enabled = !on; _prof.Enabled = !on;
                if (on) { _aluno.Text = string.Empty; _prof.Text = string.Empty; }
            };
            _lblQuem.Visible = _quem.Visible = _lblContato.Visible = _contato.Visible = false;
            _outros.CheckedChanged += (_, __) => toggleOutros();

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var spacer = new Panel { Height = 6, Dock = DockStyle.Top };
            layout.Controls.Add(spacer, 0, row);
            layout.SetColumnSpan(spacer, 2);
            row++;

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _itemEmprestado.Margin = new Padding(3, 8, 3, 8);
            _itemEmprestado.Font = inputFont;
            layout.Controls.Add(new Label { Text = string.Empty, AutoSize = true }, 0, row);
            layout.Controls.Add(_itemEmprestado, 1, row++);

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Termo (PDF)"), 0, row);
            var termoPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, AutoSize = true };
            _termo.Width = 320;
            _termo.Enabled = false;
            _termo.BackColor = Color.White;
            _termo.ForeColor = Color.Gray;
            _termo.Text = "Selecione o PDF do termo...";
            _btnTermo.Text = "📎 Anexar PDF";
            _btnTermo.AutoSize = true;
            _btnTermo.Padding = new Padding(10, 6, 10, 6);
            _btnTermo.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            _btnTermo.MinimumSize = new Size(140, 32);
            _btnTermo.FlatStyle = FlatStyle.Flat;
            _btnTermo.FlatAppearance.BorderSize = 1;
            _btnTermo.FlatAppearance.BorderColor = Color.Silver;
            _btnTermo.BackColor = Color.WhiteSmoke;
            termoPanel.Controls.Add(_termo);
            termoPanel.Controls.Add(_btnTermo);
            layout.Controls.Add(termoPanel, 1, row++);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                AutoSize = true,
                Padding = new Padding(8),
                BackColor = Color.White
            };
            _btn.AutoSize = true;
            _btn.Padding = new Padding(10, 6, 10, 6);
            _btn.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            _btn.BackColor = Color.FromArgb(0, 123, 255);
            _btn.ForeColor = Color.White;
            _btn.FlatStyle = FlatStyle.Flat;
            _btn.FlatAppearance.BorderSize = 0;
            _btn.Tag = "accent";
            var saveBase = _btn.BackColor;
            _btn.MouseEnter += (_, __) => _btn.BackColor = Darken(saveBase, 12);
            _btn.MouseLeave += (_, __) => _btn.BackColor = saveBase;
            _btn.MouseDown += (_, __) => _btn.BackColor = Darken(saveBase, 22);
            _btn.MouseUp += (_, __) => _btn.BackColor = Darken(saveBase, 12);

            _verEntregas.AutoSize = true;
            _verEntregas.Padding = new Padding(10, 6, 10, 6);
            _verEntregas.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);

            buttons.Controls.Add(_btn);
            buttons.Controls.Add(_verEntregas);

            Controls.Add(layout);
            Controls.Add(buttons);
            ResumeLayout();

            // Aplicar tema atual no carregamento
            Load += (_, __) => { try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { } };

            // Carregamento assíncrono para não travar
            Shown += async (_, __) =>
            {
                try
                {
                    var t = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                    t = t switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => t };
                    if (t == "mica") { BeginInvoke(new Action(() => { try { MedControl.UI.FluentEffects.ApplyWin11Mica(this); } catch { } })); }
                }
                catch { }
                _btnTermo.Enabled = false;
                UseWaitCursor = true;
                var loaded = await Task.Run(() =>
                {
                    var chaves = new System.Collections.Generic.List<string>();
                    try { foreach (var c in Database.GetChaves()) chaves.Add(c.Nome); } catch { }

                    var alunos = new System.Collections.Generic.List<string>();
                    var alunosSeen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        var alunosDt = Database.GetAlunosAsDataTable();
                        string? alunosNomeCol = null;
                        foreach (System.Data.DataColumn col in alunosDt.Columns)
                        {
                            var name = col.ColumnName;
                            if (string.Equals(name, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                            if (alunosNomeCol == null || string.Equals(name, "Nome", StringComparison.OrdinalIgnoreCase)) alunosNomeCol = name;
                        }
                        if (alunosNomeCol != null)
                        {
                            foreach (System.Data.DataRow r in alunosDt.Rows)
                            {
                                var v = Convert.ToString(r[alunosNomeCol]) ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(v) && alunosSeen.Add(v)) alunos.Add(v);
                            }
                        }
                        var alunosPath = Database.GetConfig("caminho_alunos");
                        if (!string.IsNullOrWhiteSpace(alunosPath))
                        {
                            foreach (var a in ExcelHelper.LoadFirstColumn(alunosPath))
                                if (alunosSeen.Add(a)) alunos.Add(a);
                        }
                    }
                    catch { }

                    var profs = new System.Collections.Generic.List<string>();
                    var profsSeen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        var profsDt = Database.GetProfessoresAsDataTable();
                        string? profsNomeCol = null;
                        foreach (System.Data.DataColumn col in profsDt.Columns)
                        {
                            var name = col.ColumnName;
                            if (string.Equals(name, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                            if (profsNomeCol == null || string.Equals(name, "Nome", StringComparison.OrdinalIgnoreCase)) profsNomeCol = name;
                        }
                        if (profsNomeCol != null)
                        {
                            foreach (System.Data.DataRow r in profsDt.Rows)
                            {
                                var v = Convert.ToString(r[profsNomeCol]) ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(v) && profsSeen.Add(v)) profs.Add(v);
                            }
                        }
                        var profsPath = Database.GetConfig("caminho_professores");
                        if (!string.IsNullOrWhiteSpace(profsPath))
                        {
                            foreach (var p in ExcelHelper.LoadFirstColumn(profsPath))
                                if (profsSeen.Add(p)) profs.Add(p);
                        }
                    }
                    catch { }

                    return (chaves, alunos, profs);
                });
                _chave.BeginUpdate();
                _aluno.BeginUpdate();
                _prof.BeginUpdate();
                _chave.DataSource = null; _aluno.DataSource = null; _prof.DataSource = null;
                _chave.DataSource = loaded.chaves;
                _aluno.DataSource = loaded.alunos;
                _prof.DataSource = loaded.profs;
                _chave.SelectedIndex = loaded.chaves.Count > 0 ? 0 : -1;
                _aluno.SelectedIndex = -1; _aluno.Text = string.Empty;
                _prof.SelectedIndex = -1; _prof.Text = string.Empty;
                _prof.EndUpdate();
                _aluno.EndUpdate();
                _chave.EndUpdate();
                UseWaitCursor = false;
            };

            _itemEmprestado.CheckedChanged += (_, __) =>
            {
                var enable = _itemEmprestado.Checked;
                _btnTermo.Enabled = enable;
                _termo.Enabled = enable;
                if (!enable)
                {
                    _termo.Text = "Selecione o PDF do termo...";
                    _termo.ForeColor = Color.Gray;
                }
            };

            _btnTermo.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf" };
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    _termo.Text = ofd.FileName;
                    _termo.ForeColor = Color.Black;
                }
            };

            _btn.Click += (_, __) =>
            {
                var usandoOutros = _outros.Checked;
                if (string.IsNullOrWhiteSpace(_chave.Text) ||
                    (
                        !usandoOutros && string.IsNullOrWhiteSpace(_aluno.Text) && string.IsNullOrWhiteSpace(_prof.Text)
                    ) ||
                    (
                        usandoOutros && (string.IsNullOrWhiteSpace(_quem.Text) || string.IsNullOrWhiteSpace(_contato.Text))
                    ))
                {
                    MessageBox.Show(this, "Informe: Chave e Aluno/Professor ou marque Outros com Quem e Contato.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var now = DateTime.Now;
                if (_itemEmprestado.Checked && (string.IsNullOrWhiteSpace(_termo.Text) || _termo.Text == "Selecione o PDF do termo..."))
                {
                    MessageBox.Show(this, "Selecione o arquivo PDF do termo.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string alunoVal = _aluno.Text.Trim();
                string profVal = _prof.Text.Trim();
                if (usandoOutros)
                {
                    alunoVal = $"{_quem.Text.Trim()} (Contato: {_contato.Text.Trim()})";
                    profVal = string.Empty;
                }
                var r = new Reserva
                {
                    Chave = _chave.Text.Trim(),
                    Aluno = alunoVal,
                    Professor = profVal,
                    DataHora = now,
                    EmUso = true,
                    Termo = _itemEmprestado.Checked ? _termo.Text : string.Empty,
                    Devolvido = false,
                    DataDevolucao = null
                };
                Database.InsertReserva(r);
                MessageBox.Show(this, $"Chave {r.Chave} entregue por {(string.IsNullOrWhiteSpace(r.Aluno) ? r.Professor : r.Aluno)} em {r.DataHora:dd/MM/yyyy HH:mm:ss}", "Entrega registrada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            };

            _verEntregas.Click += (_, __) => new EntregasForm().ShowDialog(this);
        }

        private static Color Darken(Color color, int amount)
        {
            int r = Math.Max(0, color.R - amount);
            int g = Math.Max(0, color.G - amount);
            int b = Math.Max(0, color.B - amount);
            return Color.FromArgb(color.A, r, g, b);
        }

        // Composição para reduzir flicker geral
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x02000000; // WS_EX_COMPOSITED
                return cp;
            }
        }
    }
}
