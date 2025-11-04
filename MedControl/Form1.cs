using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MedControl.UI;

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

            var chave = new ToolStripMenuItem("Chave");
            chave.DropDownItems.Add("Fazer Reserva", null, (_, __) => { new Views.ReservaForm().ShowDialog(this); RefreshKeysUi(); });
            chave.DropDownItems.Add("Entregar Chave", null, (_, __) => { new Views.EntregarChaveForm().ShowDialog(this); RefreshKeysUi(); });
            chave.DropDownItems.Add("Relação", null, (_, __) => { new Views.EntregasForm().ShowDialog(this); RefreshKeysUi(); });

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

            _menu.Items.AddRange(new ToolStripItem[] { cadastro, chave, relatorio, configuracoes, ajuda });
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
            _headerLabel.Tag = "keep-font";

            _keysPanel = new DoubleBufferedFlowLayoutPanel
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
            // Atualiza a cada segundo para refletir mudanças em tempo real (segundos)
            _refreshTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _refreshTimer.Tick += (_, __) => RefreshKeysUi();
            _refreshTimer.Start();

            Shown += (_, __) =>
            {
                ApplyTheme();
                RefreshKeysUi();
            };
        }

        // Composição reduz flicker de múltiplos controles ao atualizar o painel de chaves
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x02000000; // WS_EX_COMPOSITED
                return cp;
            }
        }

        private void RefreshKeysUi()
        {
            _keysPanel.SuspendLayout();
            // Evita flicker: suspende repaints do painel durante a atualização
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

                // Card panel: quadrado, adapta tipografia ao conteúdo
                int tileSize = 180;
                var panel = new MedControl.UI.SquareCardPanel
                {
                    SizeHint = tileSize,
                    Width = tileSize,
                    Height = tileSize,
                    BackColor = GetCardColor(),
                    BorderStyle = BorderStyle.FixedSingle,
                    Padding = new Padding(6),
                    Margin = new Padding(10)
                };

                var lbl = new Label
                {
                    Text = c.Nome,
                    Dock = DockStyle.Top,
                    Height = 24,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = GetCardColor(),
                    Font = new Font("Segoe UI", 10, FontStyle.Bold)
                };
                lbl.AutoEllipsis = true;
                lbl.AutoSize = false;

                var status = new Label
                {
                    Text = statusText,
                    BackColor = statusColor,
                    Dock = DockStyle.Top,
                    Height = 24,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI", 10, FontStyle.Bold)
                };
                status.Tag = "keep-backcolor keep-font";

                // details line: who and since when
                string detailsText = string.Empty;
                if (ativa != null)
                {
                    var nome = string.IsNullOrWhiteSpace(ativa.Aluno) ? ativa.Professor : ativa.Aluno;
                    var dur = DateTime.Now - ativa.DataHora;
                    string durStr = FormatDurationShort(dur);
                    // Mostrar segundos tanto no horário de início quanto na duração
                    detailsText = $"Com: {nome}\nDesde: {ativa.DataHora:dd/MM HH:mm:ss} • {durStr}";
                }
                else if (reservadaHoje != null)
                {
                    var nome = string.IsNullOrWhiteSpace(reservadaHoje.Aluno) ? reservadaHoje.Professor : reservadaHoje.Aluno;
                    // Mostrar segundos no horário da reserva
                    detailsText = $"Reserva: {nome}\nHorário: {reservadaHoje.DataHora:HH:mm:ss}";
                }

                var details = new Label
                {
                    Text = detailsText,
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = GetCardColor(),
                    Font = new Font("Segoe UI", 9, FontStyle.Regular)
                };
                details.AutoSize = false;
                details.Padding = new Padding(2);
                details.UseMnemonic = false;
                details.AutoEllipsis = false; // permitir múltiplas linhas
                details.MaximumSize = new Size(int.MaxValue, int.MaxValue);

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
                // Clique em qualquer parte do card (ou seus filhos) abre relação filtrada pela chave
                var chaveNome = c.Nome; // capturar para o handler
                EventHandler open = (_, __) => { new Views.EntregasForm(chaveNome).ShowDialog(this); RefreshKeysUi(); };
                panel.Cursor = Cursors.Hand;
                panel.Click += open;
                // Garantir que cliques nos filhos também abram
                void WireClick(Control ctrl)
                {
                    ctrl.Cursor = Cursors.Hand;
                    ctrl.Click += open;
                }
                WireClick(lbl);
                WireClick(status);
                WireClick(details);
                WireClick(disponibilidade);

                // Add controls in docking order (bottom -> fill -> top -> top)
                panel.Controls.Add(disponibilidade);
                panel.Controls.Add(details);
                panel.Controls.Add(status);
                panel.Controls.Add(lbl);
                _keysPanel.Controls.Add(panel);

                // Ajusta fontes para caber no cartão sem perder o formato quadrado
                FitCardTypography(panel, lbl, details);
            }

            // update badges
            _badgeDisponiveis.Text = $"Disponíveis: {countDisp}";
            _badgeReservadas.Text = $"Reservadas: {countRes}";
            _badgeEmUso.Text = $"Em uso: {countUso}";

            _keysPanel.ResumeLayout();
        }

        private void ApplyTheme()
        {
            // Sempre aplicar estilos/ícone padrão primeiro (inclui tema 'padrao')
            try { ThemeHelper.ApplyCurrentTheme(this); } catch { }
            // theme selection
            var raw = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
            var theme = raw switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => raw };
            Color bg = theme switch
            {
                "mica" => Color.FromArgb(245, 246, 250),
                "alto_contraste" => Color.Black,
                "terminal" => Color.Black,
                "claro" => Color.FromArgb(246, 247, 250),
                "escuro" => Color.FromArgb(30, 30, 30),
                "classico" => SystemColors.Control,
                "padrao" => Color.SaddleBrown,
                _ => Color.SaddleBrown
            };
            
            // Se tema for clássico, desabilitar visual styles e aplicar estilo Win95-like
            if (theme == "classico")
            {
                try { Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.NoneEnabled; } catch { }
            }
            else
            {
                try { Application.VisualStyleState = System.Windows.Forms.VisualStyles.VisualStyleState.ClientAndNonClientAreasEnabled; } catch { }
            }

            _mainPanel.BackColor = bg;
            _headerLabel.BackColor = bg;
            _statsPanel.BackColor = bg;
            _keysPanel.BackColor = bg;
            _headerLabel.ForeColor = theme == "branco" || theme == "classico" ? Color.Black : Color.White;
            
            // Reaplica o tema detalhado para estilos avançados quando não for o 'padrao'
            if (theme == "classico" || theme == "mica" || theme == "alto_contraste" || theme == "terminal" || theme == "claro" || theme == "escuro")
            {
                try { ThemeHelper.ApplyCurrentTheme(this); } catch { }
            }
        }

        private Color GetCardColor()
        {
            // default per theme
            var theme = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
            theme = theme switch { "marrom" => "padrao", "branco" => "claro", "preto" => "escuro", "azul" => "padrao", _ => theme };
            return theme switch
            {
                "mica" => Color.White,
                "alto_contraste" => Color.Black,
                "terminal" => Color.Black,
                "claro" => Color.White,
                "escuro" => Color.FromArgb(45, 45, 48),
                "classico" => SystemColors.ControlLight,
                "padrao" => Color.BurlyWood,
                _ => Color.BurlyWood
            };
        }

        private void ApplyClassicStyle(Control root)
        {
            try
            {
                var classicFont = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular);
                void Walk(Control c)
                {
                    // Fonte padrão clássica
                    try { c.Font = classicFont; } catch { }

                    // Cores clássicas
                    if (c is TextBox || c is ComboBox || c is ListBox)
                    {
                        c.BackColor = SystemColors.Window;
                        c.ForeColor = SystemColors.WindowText;
                    }
                    else if (c is DataGridView dgv)
                    {
                        c.BackColor = SystemColors.Control;
                        c.ForeColor = SystemColors.ControlText;
                        dgv.EnableHeadersVisualStyles = false;
                        dgv.BackgroundColor = SystemColors.Control;
                        dgv.GridColor = SystemColors.ControlDark;
                        dgv.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
                        dgv.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
                        dgv.ColumnHeadersDefaultCellStyle.Font = classicFont;
                        dgv.DefaultCellStyle.BackColor = SystemColors.Window;
                        dgv.DefaultCellStyle.ForeColor = SystemColors.WindowText;
                        dgv.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
                        dgv.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
                    }
                    else if (c is Button btn)
                    {
                        btn.FlatStyle = FlatStyle.Standard;
                        c.BackColor = SystemColors.Control;
                        c.ForeColor = SystemColors.ControlText;
                    }
                    else
                    {
                        c.BackColor = SystemColors.Control;
                        c.ForeColor = SystemColors.ControlText;
                    }

                    foreach (Control child in c.Controls) Walk(child);
                }
                Walk(root);
                root.PerformLayout();
                root.Refresh();
            }
            catch { }
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

        // Ajusta a tipografia do cartão para caber no espaço quadrado
        private void FitCardTypography(Panel panel, Label title, Label details)
        {
            try
            {
                // Garantir layout antes de medir
                panel.PerformLayout();

                // Ajuste do título (uma linha com reticências)
                int maxTitleWidth = Math.Max(20, panel.ClientSize.Width - (title.Margin.Horizontal + panel.Padding.Horizontal));
                float titleSize = title.Font.Size;
                while (titleSize > 8f && TextRenderer.MeasureText(title.Text, new Font(title.Font, FontStyle.Bold)).Width > maxTitleWidth)
                {
                    titleSize -= 0.5f;
                    title.Font = new Font(title.Font.FontFamily, titleSize, FontStyle.Bold);
                }

                // Ajuste dos detalhes (múltiplas linhas). Reduz a fonte se o texto for grande
                if (!string.IsNullOrWhiteSpace(details.Text))
                {
                    int usedTop = title.Height + (panel.Padding.Top);
                    int usedBottom = 18 + (panel.Padding.Bottom); // disponibilidade tem 18px
                    int availHeight = Math.Max(20, panel.ClientSize.Height - usedTop - usedBottom);
                    int availWidth = Math.Max(50, panel.ClientSize.Width - (details.Margin.Horizontal + panel.Padding.Horizontal));

                    details.MaximumSize = new Size(availWidth, availHeight);
                    float detSize = details.Font.Size;
                    int tries = 0;
                    while (tries < 12 && detSize > 7f)
                    {
                        var testFont = new Font(details.Font.FontFamily, detSize, details.Font.Style);
                        var size = TextRenderer.MeasureText(details.Text, testFont, new Size(availWidth, int.MaxValue), TextFormatFlags.WordBreak);
                        if (size.Height <= availHeight)
                        {
                            details.Font = testFont;
                            break;
                        }
                        detSize -= 0.5f; tries++;
                    }
                }
            }
            catch { }
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
                BorderStyle = BorderStyle.FixedSingle,
                Tag = "keep-backcolor keep-font"
            };
        }

        private static string FormatDurationShort(TimeSpan ts)
        {
            // Sempre que possível, incluir segundos para ver a atualização em tempo real
            if (ts.TotalDays >= 1)
            {
                int d = (int)ts.TotalDays;
                int h = ts.Hours;
                int m = ts.Minutes;
                int s = ts.Seconds;
                // Para dias, manter mais compacto: dias e horas
                return h > 0 ? $"{d}d {h}h" : $"{d}d";
            }
            if (ts.TotalHours >= 1)
            {
                int h = (int)ts.TotalHours;
                int m = ts.Minutes;
                int s = ts.Seconds;
                return $"{h}h {m}m {s}s";
            }
            if (ts.TotalMinutes >= 1)
            {
                int m = (int)ts.TotalMinutes;
                int s = ts.Seconds;
                return $"{m}m {s}s";
            }
            return $"{ts.Seconds}s";
        }

    }
}
