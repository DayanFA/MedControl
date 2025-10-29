using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MedControl
{
    public partial class Form1 : Form
    {
    private MenuStrip _menu;
    private Panel _mainPanel;
    private Label _headerLabel;
    private FlowLayoutPanel _keysPanel;
    private ToolTip _toolTip;
    private System.Windows.Forms.Timer _refreshTimer;
    private FlowLayoutPanel _statsPanel;
    private Label _badgeDisponiveis;
    private Label _badgeReservadas;
    private Label _badgeEmUso;

        public Form1()
        {
            InitializeComponent();
            Text = "Sistema de Empréstimo de Chaves";
            Width = 1000;
            Height = 700;

            Load += (_, __) => Database.Setup();

            _menu = new MenuStrip();
            var cadastro = new ToolStripMenuItem("Cadastro");
            cadastro.DropDownItems.Add("Cadastro de Alunos", null, (_, __) => { new Views.CadastroAlunosForm().ShowDialog(this); RefreshKeysUi(); });
            cadastro.DropDownItems.Add("Cadastro de Professores", null, (_, __) => { new Views.CadastroProfessoresForm().ShowDialog(this); RefreshKeysUi(); });
            cadastro.DropDownItems.Add("Cadastro de Chaves", null, (_, __) => { new Views.CadastroChavesForm().ShowDialog(this); RefreshKeysUi(); });

            var reserva = new ToolStripMenuItem("Reserva");
            reserva.DropDownItems.Add("Fazer Reserva", null, (_, __) => { new Views.ReservaForm().ShowDialog(this); RefreshKeysUi(); });

            var chave = new ToolStripMenuItem("Chave");
            chave.DropDownItems.Add("Entregar Chave", null, (_, __) => { new Views.EntregarChaveForm().ShowDialog(this); RefreshKeysUi(); });
            chave.DropDownItems.Add("Ver Entregas", null, (_, __) => { new Views.EntregasForm().ShowDialog(this); RefreshKeysUi(); });

            var relatorio = new ToolStripMenuItem("Relatório");
            relatorio.DropDownItems.Add("Ver Relatórios", null, (_, __) => new Views.RelatorioForm().ShowDialog(this));

            var configuracoes = new ToolStripMenuItem("Configurações");
            configuracoes.DropDownItems.Add("Configurações Gerais", null, (_, __) =>
            {
                new Views.ConfiguracoesForm().ShowDialog(this);
                ApplyTheme();
                RefreshKeysUi();
            });

            var ajuda = new ToolStripMenuItem("Ajuda");
            ajuda.DropDownItems.Add("Sobre", null, (_, __) => new Views.AjudaForm().ShowDialog(this));

            _menu.Items.AddRange(new ToolStripItem[] { cadastro, reserva, chave, relatorio, configuracoes, ajuda });
            MainMenuStrip = _menu;
            Controls.Add(_menu);

            // Styled main panel similar to Python UI
            _mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.SaddleBrown,
                Padding = new Padding(20, 20, 20, 20)
            };

            _headerLabel = new Label
            {
                Text = "Chaves Disponíveis:",
                ForeColor = Color.White,
                BackColor = Color.SaddleBrown,
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(0, 10, 0, 10)
            };

            _keysPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                WrapContents = true,
                BackColor = Color.SaddleBrown
            };

            _toolTip = new ToolTip
            {
                AutoPopDelay = 10000,
                InitialDelay = 300,
                ReshowDelay = 100,
                ShowAlways = true
            };

            // status badges panel
            _statsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 40,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = Color.SaddleBrown,
                Padding = new Padding(0, 5, 0, 10)
            };

            _badgeDisponiveis = CreateBadgeLabel("Disponíveis: 0", Color.LightGreen);
            _badgeReservadas = CreateBadgeLabel("Reservadas: 0", Color.Khaki);
            _badgeEmUso = CreateBadgeLabel("Em uso: 0", Color.LightCoral);
            _statsPanel.Controls.AddRange(new Control[] { _badgeDisponiveis, _badgeReservadas, _badgeEmUso });

            _mainPanel.Controls.Add(_keysPanel);
            _mainPanel.Controls.Add(_statsPanel);
            _mainPanel.Controls.Add(_headerLabel);
            Controls.Add(_mainPanel);

            // periodic refresh (optional)
            _refreshTimer = new System.Windows.Forms.Timer { Interval = 30000 };
            _refreshTimer.Tick += (_, __) => RefreshKeysUi();
            _refreshTimer.Start();

            Shown += (_, __) =>
            {
                ApplyTheme();
                RefreshKeysUi();
            };
        }

        private void RefreshKeysUi()
        {
            _keysPanel.SuspendLayout();
            _keysPanel.Controls.Clear();

            var reservas = Database.GetReservas();
            var chaves = Database.GetChaves();

            int countDisp = 0, countRes = 0, countUso = 0;

            foreach (var c in chaves)
            {
                // Determine status
                var statusColor = Color.LightGreen; // Disponível
                var statusText = "Disponível";
                var nowDate = DateTime.Now.Date;
                var reservasDaChave = reservas.Where(x => x.Chave == c.Nome).ToList();

                var ativa = reservasDaChave
                    .Where(r => !r.Devolvido && r.EmUso)
                    .OrderByDescending(r => r.DataHora)
                    .FirstOrDefault();
                var reservadaHoje = reservasDaChave
                    .Where(r => !r.Devolvido && !r.EmUso && r.DataHora.Date == nowDate)
                    .OrderBy(r => r.DataHora)
                    .FirstOrDefault();

                foreach (var r in reservasDaChave)
                {
                    if (!r.Devolvido)
                    {
                        if (r.EmUso)
                        {
                            statusColor = Color.LightCoral; // Em uso
                            statusText = "Em Uso"; break;
                        }
                        else if (r.DataHora.Date == nowDate)
                        {
                            statusColor = Color.Khaki; // Reservado hoje
                            statusText = "Reservado";
                        }
                    }
                }

                // counts
                if (statusText == "Disponível") countDisp++;
                else if (statusText == "Reservado") countRes++;
                else if (statusText == "Em Uso") countUso++;

                // Card panel styled like Python (burlywood) with extra details
                var panel = new Panel
                {
                    Width = 180,
                    Height = 130,
                    BackColor = GetCardColor(),
                    BorderStyle = BorderStyle.FixedSingle,
                    Padding = new Padding(6),
                    Margin = new Padding(10)
                };

                var lbl = new Label
                {
                    Text = c.Nome,
                    Dock = DockStyle.Top,
                    Height = 22,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = GetCardColor(),
                    Font = new Font("Segoe UI", 10, FontStyle.Bold)
                };

                var status = new Label
                {
                    Text = statusText,
                    BackColor = statusColor,
                    Dock = DockStyle.Top,
                    Height = 24,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI", 10, FontStyle.Regular)
                };

                // details line: who and since when
                string detailsText = string.Empty;
                if (ativa != null)
                {
                    var nome = string.IsNullOrWhiteSpace(ativa.Aluno) ? ativa.Professor : ativa.Aluno;
                    var dur = DateTime.Now - ativa.DataHora;
                    string durStr = FormatDurationShort(dur);
                    detailsText = $"Com: {nome}\nDesde: {ativa.DataHora:dd/MM HH:mm} • {durStr}";
                }
                else if (reservadaHoje != null)
                {
                    var nome = string.IsNullOrWhiteSpace(reservadaHoje.Aluno) ? reservadaHoje.Professor : reservadaHoje.Aluno;
                    detailsText = $"Reserva: {nome}\nHorário: {reservadaHoje.DataHora:HH:mm}";
                }

                var details = new Label
                {
                    Text = detailsText,
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = GetCardColor(),
                    Font = new Font("Segoe UI", 9, FontStyle.Regular)
                };

                // availability based on num_copias
                int emUsoCount = reservasDaChave.Count(r => !r.Devolvido && r.EmUso);
                int disponiveis = Math.Max(0, (c.NumCopias <= 0 ? 0 : c.NumCopias) - emUsoCount);
                var disponibilidade = new Label
                {
                    Text = $"Disponíveis: {disponiveis} de {c.NumCopias}",
                    Dock = DockStyle.Bottom,
                    Height = 18,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = GetCardColor(),
                    Font = new Font("Segoe UI", 8, FontStyle.Italic)
                };

                // Tooltip with details similar to Python
                var reservasText = string.Join("\n",
                    reservasDaChave.Select(r =>
                    {
                        var nome = string.IsNullOrWhiteSpace(r.Aluno) ? r.Professor : r.Aluno;
                        return $"{nome} - {r.DataHora:dd/MM/yyyy HH:mm:ss}";
                    }));
                var statusInfo = $"Status: {statusText}" + (string.IsNullOrEmpty(reservasText) ? string.Empty : $"\nReservas:\n{reservasText}");
                _toolTip.SetToolTip(status, statusInfo);
                _toolTip.SetToolTip(panel, statusInfo);

                // Click opens Entregas, then refresh
                panel.Cursor = Cursors.Hand;
                panel.Click += (_, __) => { new Views.EntregasForm().ShowDialog(this); RefreshKeysUi(); };

                // Add controls in docking order (bottom -> fill -> top -> top)
                panel.Controls.Add(disponibilidade);
                panel.Controls.Add(details);
                panel.Controls.Add(status);
                panel.Controls.Add(lbl);
                _keysPanel.Controls.Add(panel);
            }

            // update badges
            _badgeDisponiveis.Text = $"Disponíveis: {countDisp}";
            _badgeReservadas.Text = $"Reservadas: {countRes}";
            _badgeEmUso.Text = $"Em uso: {countUso}";

            _keysPanel.ResumeLayout();
        }

        private void ApplyTheme()
        {
            // theme selection
            var theme = (Database.GetConfig("theme") ?? "marrom").ToLowerInvariant();
            Color bg = theme switch
            {
                "preto" => Color.FromArgb(0x24, 0x24, 0x24),
                "branco" => Color.WhiteSmoke,
                "azul" => Color.FromArgb(0x1E, 0x3A, 0x5F),
                _ => Color.SaddleBrown
            };
            // override via custom hex
            var customBg = Database.GetConfig("color_background");
            if (!string.IsNullOrWhiteSpace(customBg) && TryParseHexColor(customBg, out var parsed))
                bg = parsed;

            _mainPanel.BackColor = bg;
            _headerLabel.BackColor = bg;
            _statsPanel.BackColor = bg;
            _keysPanel.BackColor = bg;
            _headerLabel.ForeColor = theme == "claro" ? Color.Black : Color.White;
        }

        private Color GetCardColor()
        {
            var cardHex = Database.GetConfig("color_card");
            if (!string.IsNullOrWhiteSpace(cardHex) && TryParseHexColor(cardHex, out var parsed))
                return parsed;
            // default per theme
            var theme = (Database.GetConfig("theme") ?? "marrom").ToLowerInvariant();
            return theme switch
            {
                "preto" => Color.FromArgb(60, 60, 60),
                "branco" => Color.White,
                "azul" => Color.FromArgb(0x2E, 0x5C, 0x8A),
                _ => Color.BurlyWood
            };
        }

        private static bool TryParseHexColor(string hex, out Color color)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(hex)) { color = Color.Empty; return false; }
                hex = hex.Trim();
                if (hex.StartsWith("#")) hex = hex.Substring(1);
                if (hex.Length == 6)
                {
                    var r = Convert.ToInt32(hex.Substring(0, 2), 16);
                    var g = Convert.ToInt32(hex.Substring(2, 2), 16);
                    var b = Convert.ToInt32(hex.Substring(4, 2), 16);
                    color = Color.FromArgb(r, g, b);
                    return true;
                }
            }
            catch { }
            color = Color.Empty;
            return false;
        }


        private static Label CreateBadgeLabel(string text, Color back)
        {
            return new Label
            {
                Text = text,
                AutoSize = false,
                Width = 150,
                Height = 28,
                Margin = new Padding(0, 0, 10, 0),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = back,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private static string FormatDurationShort(TimeSpan ts)
        {
            if (ts.TotalDays >= 1)
            {
                int d = (int)ts.TotalDays;
                int h = ts.Hours;
                return h > 0 ? $"{d}d {h}h" : $"{d}d";
            }
            if (ts.TotalHours >= 1)
            {
                int h = (int)ts.TotalHours;
                int m = ts.Minutes;
                return m > 0 ? $"{h}h {m}m" : $"{h}h";
            }
            if (ts.TotalMinutes >= 1)
            {
                int m = (int)ts.TotalMinutes;
                int s = ts.Seconds;
                return s > 0 ? $"{m}m {s}s" : $"{m}m";
            }
            return $"{ts.Seconds}s";
        }

    }
}
