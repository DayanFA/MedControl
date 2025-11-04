using System;
using System.Drawing;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class ConfiguracoesForm : Form
    {
    private readonly TextBox _alunos = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _profs = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly Button _browseAlunos = new Button { Text = "📂 Selecionar Alunos.xlsx" };
    private readonly Button _browseProfs = new Button { Text = "📂 Selecionar Professores.xlsx" };
    private readonly Button _salvar = new Button { Text = "Salvar" };
        // Banco de dados
        private readonly ComboBox _dbProvider = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _mysqlConn = new TextBox { Dock = DockStyle.Fill };
        // Visual
    private readonly ComboBox _tema = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };

        public ConfiguracoesForm()
        {
            Text = "Configurações";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Width = 640;
            Height = 520;
            this.DoubleBuffered = true;

            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(18, 18, 18, 8), AutoSize = true };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            // Rows
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 0 alunos
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1 profs
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2 provider
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3 mysql conn
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 4 tema
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 5 salvar

            Label MkLabel(string text) => new Label
            {
                Text = text,
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(3, 8, 12, 8),
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point)
            };

            void StyleButton(Button b, bool primary = false)
            {
                b.AutoSize = true;
                b.Padding = new Padding(10, 6, 10, 6);
                b.Font = new Font("Segoe UI", primary ? 10F : 10F, primary ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Point);
                b.FlatStyle = FlatStyle.Flat;
                if (primary)
                {
                    b.BackColor = Color.FromArgb(0, 123, 255);
                    b.ForeColor = Color.White;
                    b.FlatAppearance.BorderSize = 0;
                }
                else
                {
                    b.BackColor = Color.WhiteSmoke;
                    b.FlatAppearance.BorderSize = 1;
                    b.FlatAppearance.BorderColor = Color.Silver;
                }
            }

            // Alunos
            _alunos.Font = new Font("Segoe UI", 10F);
            try { _alunos.PlaceholderText = "Selecione o arquivo de alunos (xlsx)…"; } catch { }
            table.Controls.Add(MkLabel("Planilha de Alunos:"), 0, 0);
            table.Controls.Add(_alunos, 1, 0);
            StyleButton(_browseAlunos);
            table.Controls.Add(_browseAlunos, 2, 0);
            // Professores
            _profs.Font = new Font("Segoe UI", 10F);
            try { _profs.PlaceholderText = "Selecione o arquivo de professores (xlsx)…"; } catch { }
            table.Controls.Add(MkLabel("Planilha de Professores:"), 0, 1);
            table.Controls.Add(_profs, 1, 1);
            StyleButton(_browseProfs);
            table.Controls.Add(_browseProfs, 2, 1);
            // DB provider
            table.Controls.Add(MkLabel("Banco de Dados:"), 0, 2);
            _dbProvider.Items.AddRange(new object[] { "SQLite (local)", "MySQL (servidor)" });
            _dbProvider.Font = new Font("Segoe UI", 10F);
            table.Controls.Add(_dbProvider, 1, 2);
            table.Controls.Add(new Label { Text = " " }, 2, 2);
            // MySQL connection string
            table.Controls.Add(MkLabel("MySQL Connection String:"), 0, 3);
            _mysqlConn.Font = new Font("Segoe UI", 10F);
            try { _mysqlConn.PlaceholderText = "Ex.: Server=localhost;Database=medcontrol;Uid=root;Pwd=***;"; } catch { }
            table.Controls.Add(_mysqlConn, 1, 3);
            table.Controls.Add(new Label { Text = " " }, 2, 3);
            // Visual theme
            table.Controls.Add(MkLabel("Tema Visual:"), 0, 4);
            _tema.Items.AddRange(new object[] { "Padrão", "Claro", "Escuro", "Clássico (95)", "Translúcido (Mica)", "Alto Contraste", "Terminal (CRT)" });
            _tema.Font = new Font("Segoe UI", 10F);
            table.Controls.Add(_tema, 1, 4);
            table.Controls.Add(new Label { Text = " " }, 2, 4);
            // Save
            StyleButton(_salvar, primary: true);
            _salvar.Tag = "accent";
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(8) };
            buttons.Controls.Add(_salvar);
            table.Controls.Add(buttons, 0, 5);
            table.SetColumnSpan(buttons, 3);
            Controls.Add(table);

            Load += (_, __) =>
            {
                try
                {
                    var t = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                    t = t switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => t };
                    if (t == "mica") { BeginInvoke(new Action(() => { try { MedControl.UI.FluentEffects.ApplyWin11Mica(this); } catch { } })); }
                }
                catch { }
                _alunos.Text = Database.GetConfig("caminho_alunos") ?? string.Empty;
                _profs.Text = Database.GetConfig("caminho_professores") ?? string.Empty;

                // DB provider/config from AppConfig
                var prov = (AppConfig.Instance.DbProvider ?? "sqlite").ToLowerInvariant();
                _dbProvider.SelectedIndex = prov == "mysql" ? 1 : 0;
                _mysqlConn.Text = AppConfig.Instance.MySqlConnectionString ?? string.Empty;

                var themeRaw = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                var theme = themeRaw switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => themeRaw };
                _tema.SelectedIndex = theme switch
                {
                    "padrao" => 0,
                    "claro" => 1,
                    "escuro" => 2,
                    "classico" => 3,
                    "mica" => 4,
                    "alto_contraste" => 5,
                    "terminal" => 6,
                    _ => 0
                };
                // Aplicar tema atual (inclui clássico, mica, alto contraste, terminal)
                try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
            };

            _browseAlunos.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog(this) == DialogResult.OK) _alunos.Text = ofd.FileName;
            };

            _browseProfs.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog(this) == DialogResult.OK) _profs.Text = ofd.FileName;
            };

            _dbProvider.SelectedIndexChanged += (_, __) =>
            {
                bool isMy = _dbProvider.SelectedIndex == 1;
                _mysqlConn.Enabled = isMy;
                _mysqlConn.BackColor = isMy ? Color.White : Color.Gainsboro;
            };

            _tema.SelectedIndexChanged += (_, __) =>
            {
                // Nada a fazer: usamos apenas temas pré-definidos agora
            };

            _salvar.Click += (_, __) =>
            {
                if (!string.IsNullOrWhiteSpace(_alunos.Text)) Database.SetConfig("caminho_alunos", _alunos.Text);
                if (!string.IsNullOrWhiteSpace(_profs.Text)) Database.SetConfig("caminho_professores", _profs.Text);

                // Save DB provider config (external JSON)
                AppConfig.Instance.DbProvider = _dbProvider.SelectedIndex == 1 ? "mysql" : "sqlite";
                AppConfig.Instance.MySqlConnectionString = _mysqlConn.Text?.Trim();
                if (_dbProvider.SelectedIndex == 1 && string.IsNullOrWhiteSpace(AppConfig.Instance.MySqlConnectionString))
                {
                    MessageBox.Show(this, "Informe a Connection String para MySQL.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                AppConfig.Save();

                // Save theme
                var themeKey = _tema.SelectedIndex switch
                {
                    0 => "padrao",
                    1 => "claro",
                    2 => "escuro",
                    3 => "classico",
                    4 => "mica",
                    5 => "alto_contraste",
                    6 => "terminal",
                    _ => "padrao"
                };
                Database.SetConfig("theme", themeKey);
                // Aplicar imediatamente a todos os formulários abertos e ajustar VisualStyleState
                try { MedControl.UI.ThemeHelper.ApplyVisualStyleStateForCurrentTheme(); MedControl.UI.ThemeHelper.ApplyToAllOpenForms(); } catch { }

                // Re-run setup to ensure schema exists on the newly selected provider
                try { Database.Setup(); } catch { }
                MessageBox.Show("Configurações salvas. O provedor de banco será usado imediatamente nas próximas operações.");
                Close();
            };
        }

    }
}
