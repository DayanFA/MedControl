using System;
using System.Drawing;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class ConfiguracoesForm : Form
    {
        private readonly TextBox _alunos = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        private readonly TextBox _profs = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        private readonly Button _browseAlunos = new Button { Text = "Selecionar Alunos.xlsx" };
        private readonly Button _browseProfs = new Button { Text = "Selecionar Professores.xlsx" };
        private readonly Button _salvar = new Button { Text = "Salvar" };
        // Banco de dados
        private readonly ComboBox _dbProvider = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _mysqlConn = new TextBox { Dock = DockStyle.Fill };
        // Visual
        private readonly ComboBox _tema = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _bgHex = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        private readonly TextBox _cardHex = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        private readonly Button _pickBg = new Button { Text = "Escolher cor…" };
        private readonly Button _pickCard = new Button { Text = "Escolher cor…" };

        public ConfiguracoesForm()
        {
            Text = "Configurações";
            Width = 600;
            Height = 460;

            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(10) };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            // Rows
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 0 alunos
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 1 profs
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 2 provider
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 3 mysql conn
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 4 tema
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 5 bg
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 6 card
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 7 salvar

            // Alunos
            table.Controls.Add(new Label { Text = "Planilha de Alunos:" }, 0, 0);
            table.Controls.Add(_alunos, 1, 0);
            table.Controls.Add(_browseAlunos, 2, 0);
            // Professores
            table.Controls.Add(new Label { Text = "Planilha de Professores:" }, 0, 1);
            table.Controls.Add(_profs, 1, 1);
            table.Controls.Add(_browseProfs, 2, 1);
            // DB provider
            table.Controls.Add(new Label { Text = "Banco de Dados:" }, 0, 2);
            _dbProvider.Items.AddRange(new object[] { "SQLite (local)", "MySQL (servidor)" });
            table.Controls.Add(_dbProvider, 1, 2);
            table.Controls.Add(new Label { Text = " " }, 2, 2);
            // MySQL connection string
            table.Controls.Add(new Label { Text = "MySQL Connection String:" }, 0, 3);
            table.Controls.Add(_mysqlConn, 1, 3);
            table.Controls.Add(new Label { Text = " " }, 2, 3);
            // Visual theme
            table.Controls.Add(new Label { Text = "Tema Visual:" }, 0, 4);
            _tema.Items.AddRange(new object[] { "Marrom (padrão)", "Preto", "Branco", "Azul" });
            table.Controls.Add(_tema, 1, 4);
            table.Controls.Add(new Label { Text = " " }, 2, 4);
            // Colors
            table.Controls.Add(new Label { Text = "Cor de Fundo (opcional):" }, 0, 5);
            table.Controls.Add(_bgHex, 1, 5);
            table.Controls.Add(_pickBg, 2, 5);
            table.Controls.Add(new Label { Text = "Cor dos Cartões (opcional):" }, 0, 6);
            table.Controls.Add(_cardHex, 1, 6);
            table.Controls.Add(_pickCard, 2, 6);
            // Save
            table.Controls.Add(_salvar, 2, 7);
            Controls.Add(table);

            Load += (_, __) =>
            {
                _alunos.Text = Database.GetConfig("caminho_alunos") ?? string.Empty;
                _profs.Text = Database.GetConfig("caminho_professores") ?? string.Empty;

                // DB provider/config from AppConfig
                var prov = (AppConfig.Instance.DbProvider ?? "sqlite").ToLowerInvariant();
                _dbProvider.SelectedIndex = prov == "mysql" ? 1 : 0;
                _mysqlConn.Text = AppConfig.Instance.MySqlConnectionString ?? string.Empty;

                var theme = (Database.GetConfig("theme") ?? "marrom").ToLowerInvariant();
                _tema.SelectedIndex = theme switch
                {
                    "preto" => 1,
                    "branco" => 2,
                    "azul" => 3,
                    _ => 0
                };
                _bgHex.Text = Database.GetConfig("color_background") ?? string.Empty;
                _cardHex.Text = Database.GetConfig("color_card") ?? string.Empty;
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

            _pickBg.Click += (_, __) =>
            {
                using var cd = new ColorDialog();
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    _bgHex.Text = ColorToHex(cd.Color);
                }
            };

            _pickCard.Click += (_, __) =>
            {
                using var cd = new ColorDialog();
                if (cd.ShowDialog(this) == DialogResult.OK)
                {
                    _cardHex.Text = ColorToHex(cd.Color);
                }
            };

            _tema.SelectedIndexChanged += (_, __) =>
            {
                var (bg, card) = GetDefaultsForThemeIndex(_tema.SelectedIndex);
                _bgHex.Text = bg;
                _cardHex.Text = card;
            };

            _salvar.Click += (_, __) =>
            {
                if (!string.IsNullOrWhiteSpace(_alunos.Text)) Database.SetConfig("caminho_alunos", _alunos.Text);
                if (!string.IsNullOrWhiteSpace(_profs.Text)) Database.SetConfig("caminho_professores", _profs.Text);

                // Save DB provider config (external JSON)
                AppConfig.Instance.DbProvider = _dbProvider.SelectedIndex == 1 ? "mysql" : "sqlite";
                AppConfig.Instance.MySqlConnectionString = _mysqlConn.Text?.Trim();
                AppConfig.Save();

                // Save theme
                var themeKey = _tema.SelectedIndex switch
                {
                    1 => "preto",
                    2 => "branco",
                    3 => "azul",
                    _ => "marrom"
                };
                Database.SetConfig("theme", themeKey);
                if (!string.IsNullOrWhiteSpace(_bgHex.Text)) Database.SetConfig("color_background", _bgHex.Text);
                if (!string.IsNullOrWhiteSpace(_cardHex.Text)) Database.SetConfig("color_card", _cardHex.Text);

                // Re-run setup to ensure schema exists on the newly selected provider
                try { Database.Setup(); } catch { }
                MessageBox.Show("Configurações salvas. O provedor de banco será usado imediatamente nas próximas operações.");
                Close();
            };
        }

        private static string ColorToHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

        private static (string bg, string card) GetDefaultsForThemeIndex(int idx)
        {
            // Return hex defaults for presets
            return idx switch
            {
                1 => ("#242424", "#3C3C3C"), // Preto
                2 => ("#F5F5F5", "#FFFFFF"), // Branco
                3 => ("#1E3A5F", "#2E5C8A"), // Azul
                _ => ("#8B4513", "#DEB887")  // Marrom padrão: SaddleBrown / BurlyWood
            };
        }
    }
}
