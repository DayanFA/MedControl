using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class ReservaForm : Form
    {
    private ComboBox _chave = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown }; // permitir digitar como Aluno/Professor
        private ComboBox _aluno = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown }; // permite digitar
        private ComboBox _prof = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
        private DateTimePicker _dataHora = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy HH:mm:ss", ShowUpDown = true };
        private CheckBox _outros = new CheckBox { Text = "Outros" };
        private TextBox _quem = new TextBox();
        private TextBox _contato = new TextBox();
        private Button _btnSalvar = new Button();
        private Button _btnCancelar = new Button();

        public ReservaForm()
        {
            Text = "Fazer Reserva";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // Reduzir flicker geral
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
            layout.SuspendLayout();
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

            // Fonte e dimensões dos inputs
            var inputFont = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            foreach (var c in new Control[] { _chave, _aluno, _prof, _quem, _contato })
            {
                c.Font = inputFont;
                c.Margin = new Padding(3, 8, 3, 8);
                c.MinimumSize = new Size(260, 30);
                c.Dock = DockStyle.Fill;
            }
            _outros.Font = inputFont; _outros.Margin = new Padding(3,8,3,8);
            _dataHora.Font = inputFont;
            _dataHora.Margin = new Padding(3, 8, 3, 8);
            _dataHora.MinimumSize = new Size(200, 30);
            _dataHora.Dock = DockStyle.Left;

            int row = 0;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Chave"), 0, row);
            _chave.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _chave.AutoCompleteSource = AutoCompleteSource.ListItems;
            layout.Controls.Add(_chave, 1, row++);

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Aluno"), 0, row);
            _aluno.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _aluno.AutoCompleteSource = AutoCompleteSource.ListItems;
            layout.Controls.Add(_aluno, 1, row++);

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(MkLabel("Professor"), 0, row);
            _prof.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _prof.AutoCompleteSource = AutoCompleteSource.ListItems;
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
            layout.Controls.Add(MkLabel("Data e Hora"), 0, row);
            layout.Controls.Add(_dataHora, 1, row++);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                AutoSize = true,
                Padding = new Padding(8),
                BackColor = Color.White
            };
            _btnSalvar.Text = "Reservar";
            _btnSalvar.AutoSize = true;
            _btnSalvar.Padding = new Padding(10, 6, 10, 6);
            _btnSalvar.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            _btnSalvar.BackColor = Color.FromArgb(0, 123, 255);
            _btnSalvar.ForeColor = Color.White;
            _btnSalvar.FlatStyle = FlatStyle.Flat;
            _btnSalvar.FlatAppearance.BorderSize = 0;
            _btnSalvar.Tag = "accent";
            var saveBase = _btnSalvar.BackColor;
            _btnSalvar.MouseEnter += (_, __) => _btnSalvar.BackColor = Darken(saveBase, 12);
            _btnSalvar.MouseLeave += (_, __) => _btnSalvar.BackColor = saveBase;
            _btnSalvar.MouseDown += (_, __) => _btnSalvar.BackColor = Darken(saveBase, 22);
            _btnSalvar.MouseUp += (_, __) => _btnSalvar.BackColor = Darken(saveBase, 12);

            _btnCancelar.Text = "Cancelar";
            _btnCancelar.AutoSize = true;
            _btnCancelar.Padding = new Padding(10, 6, 10, 6);
            _btnCancelar.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);

            buttons.Controls.Add(_btnSalvar);
            buttons.Controls.Add(_btnCancelar);

            Controls.Add(layout);
            Controls.Add(buttons);
            layout.ResumeLayout();
            ResumeLayout();

            // Aplicar tema atual, conforme configuração
            Load += (_, __) => { try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { } };

            // Carregamento assíncrono para evitar travar a abertura do diálogo (sem placeholders para não piscar)
            Shown += async (_, __) =>
            {
                // Aplicar Mica somente se o tema atual for 'mica'
                try
                {
                    var t = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                    // normaliza chaves legadas
                    t = t switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => t };
                    if (t == "mica") { BeginInvoke(new Action(() => { try { MedControl.UI.FluentEffects.ApplyWin11Mica(this); } catch { } })); }
                }
                catch { }

                // Indica carregamento sem desabilitar inputs (evita aparência acinzentada)
                this.UseWaitCursor = true;

                var result = await Task.Run(() =>
                {
                    var chaves = new System.Collections.Generic.List<string>();
                    try
                    {
                        foreach (var c in Database.GetChaves()) chaves.Add(c.Nome);
                    }
                    catch { }

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
                            foreach (System.Data.DataRow row in alunosDt.Rows)
                            {
                                var val = Convert.ToString(row[alunosNomeCol]) ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(val) && alunosSeen.Add(val)) alunos.Add(val);
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
                            foreach (System.Data.DataRow row in profsDt.Rows)
                            {
                                var val = Convert.ToString(row[profsNomeCol]) ?? string.Empty;
                                if (!string.IsNullOrWhiteSpace(val) && profsSeen.Add(val)) profs.Add(val);
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

                // Atualização única dos combos para evitar repaints múltiplos
                _chave.BeginUpdate();
                _aluno.BeginUpdate();
                _prof.BeginUpdate();

                _chave.DataSource = null;
                _aluno.DataSource = null;
                _prof.DataSource = null;

                _chave.DataSource = result.chaves;
                _aluno.DataSource = result.alunos;
                _prof.DataSource = result.profs;

                _chave.SelectedIndex = result.chaves.Count > 0 ? 0 : -1;
                _aluno.SelectedIndex = -1; _aluno.Text = string.Empty;
                _prof.SelectedIndex = -1; _prof.Text = string.Empty;

                _chave.EndUpdate();
                _aluno.EndUpdate();
                _prof.EndUpdate();

                this.UseWaitCursor = false;
            };

            AcceptButton = _btnSalvar;
            CancelButton = _btnCancelar;

            _btnCancelar.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            _btnSalvar.Click += (_, __) =>
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
                    DataHora = _dataHora.Value,
                    EmUso = false,
                    Termo = string.Empty,
                    Devolvido = false,
                    DataDevolucao = null
                };
                Database.InsertReserva(r);
                MessageBox.Show(this, $"Chave {r.Chave} reservada por {(string.IsNullOrWhiteSpace(r.Aluno) ? r.Professor : r.Aluno)} em {r.DataHora:dd/MM/yyyy HH:mm:ss}", "Reserva criada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            };

            // efeito Win11 deslocado para Shown (acima) para evitar custo no momento de criação do handle
        }

        // Composição de janela para reduzir flicker em redesenho de múltiplos controles
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x02000000; // WS_EX_COMPOSITED
                return cp;
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
