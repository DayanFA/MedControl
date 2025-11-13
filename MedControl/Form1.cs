using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using MedControl.UI;
using System.IO;

namespace MedControl
{
    public partial class Form1 : Form
    {
    public static Form1? Instance; // referência estática
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
    private ToolStripLabel _statusDot;
    private ToolStripLabel _statusText;
    private System.Windows.Forms.Timer? _statusTimer;
    // Usado para reduzir flicker: evita reconstruir a UI se nada relevante mudou
    private string? _lastKeysSignature;
    // Último RTT medido (ms) para exibir na barra de status
    private int _lastRttMs = -1;
    private System.Windows.Forms.Timer? _heartbeatTimer;
    private bool _allowExit = false; // controla saída real
    private NotifyIcon? _trayIcon; // ícone de bandeja
    private ContextMenuStrip? _trayMenu; // menu da bandeja

        public Form1()
        {
            InitializeComponent();
            Instance = this;
            Text = "Sistema de Empréstimo de Chaves";
            Width = 1000;
            Height = 700;

            // Carrega ícone único do aplicativo (janela + bandeja) se existir
            try
            {
                var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "app.ico");
                if (File.Exists(iconPath))
                {
                    // Usa um único objeto Icon para manter consistência
                    var ico = new Icon(iconPath);
                    this.Icon = ico;
                }
            }
            catch { }

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
            configuracoes.DropDownItems.Add("Conexões", null, (_, __) =>
            {
                new Views.ConexoesForm().ShowDialog(this);
            });

            var ajuda = new ToolStripMenuItem("Ajuda");
            ajuda.DropDownItems.Add("Sobre", null, (_, __) => new Views.AjudaForm().ShowDialog(this));

            _menu.Items.AddRange(new ToolStripItem[] { cadastro, chave, relatorio, configuracoes, ajuda });
            _statusText = new ToolStripLabel(" verificando...")
            {
                Alignment = ToolStripItemAlignment.Right,
                Margin = new Padding(2, 0, 0, 0),
                BackColor = Color.Transparent,
                ForeColor = Color.Black
            };
            _statusDot = new ToolStripLabel("●")
            {
                Alignment = ToolStripItemAlignment.Right,
                Margin = new Padding(12, 0, 0, 0),
                BackColor = Color.Transparent,
                ForeColor = Color.Gray
            };
            _menu.Items.Add(_statusText);
            _menu.Items.Add(_statusDot);
            MainMenuStrip = _menu;
            Controls.Add(_menu);

            // Bandeja (NotifyIcon) para manter aplicação acessível quando minimizada
            InitTrayIcon();

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
            // Mantém o timer em 1s, mas vamos evitar rebuild quando nada mudou
            _refreshTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _refreshTimer.Tick += (_, __) => RefreshKeysUi();
            _refreshTimer.Start();

            Shown += (_, __) =>
            {
                ApplyTheme();
                RefreshKeysUi();
                try
                {
                    _statusTimer = new System.Windows.Forms.Timer { Interval = 7000 };
                    _statusTimer.Tick += (_, __) => UpdateConnectionStatus(false);
                    _statusTimer.Start();
                    UpdateConnectionStatus(true);
                    // Heartbeat de saúde a cada 5 minutos para diagnosticar fechamentos inesperados
                    _heartbeatTimer = new System.Windows.Forms.Timer { Interval = 5 * 60 * 1000 };
                    _heartbeatTimer.Tick += (_, __) => WriteHeartbeat();
                    _heartbeatTimer.Start();
                    WriteHeartbeat();
                    // Listener para evento de restauração (segunda instância)
                    Task.Run(() =>
                    {
                        try
                        {
                            var restoreEvt = new EventWaitHandle(false, EventResetMode.AutoReset, "MedControl_Restore_Event");
                            while (!Program.IsShuttingDown())
                            {
                                restoreEvt.WaitOne();
                                if (Program.IsShuttingDown()) break;
                                try
                                {
                                    if (IsHandleCreated)
                                    {
                                        BeginInvoke(new Action(() =>
                                        {
                                            RestoreFromTray();
                                            Activate();
                                            BringToFront();
                                        }));
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    });
                }
                catch { }
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
            try
            {
                // 1) Coleta dados e calcula uma assinatura estável (ignora variações por segundo)
                var reservas = Database.GetReservas();
                var chaves = Database.GetChaves();

                int countDisp = 0, countRes = 0, countUso = 0;
                var now = DateTime.Now;
                var nowDate = now.Date;

                System.Text.StringBuilder sig = new System.Text.StringBuilder();
                foreach (var c in chaves)
                {
                    var reservasDaChave = reservas.Where(x => x.Chave == c.Nome).ToList();
                    var ativa = reservasDaChave
                        .Where(r => !r.Devolvido && r.EmUso)
                        .OrderByDescending(r => r.DataHora)
                        .FirstOrDefault();
                    var reservadaHoje = reservasDaChave
                        .Where(r => !r.Devolvido && !r.EmUso && r.DataHora.Date == nowDate)
                        .OrderBy(r => r.DataHora)
                        .FirstOrDefault();

                    string statusText = "Disponível";
                    foreach (var r in reservasDaChave)
                    {
                        if (!r.Devolvido)
                        {
                            if (r.EmUso) { statusText = "Em Uso"; break; }
                            else if (r.DataHora.Date == nowDate) { statusText = "Reservado"; }
                        }
                    }

                    if (statusText == "Disponível") countDisp++; else if (statusText == "Reservado") countRes++; else if (statusText == "Em Uso") countUso++;

                    // Disponibilidade por cópias atuais
                    int emUsoCount = reservasDaChave.Count(r => !r.Devolvido && r.EmUso);
                    int disponiveis = Math.Max(0, (c.NumCopias <= 0 ? 0 : c.NumCopias) - emUsoCount);

                    // Para a assinatura, ignorar segundos: usar minutos de duração
                    int minutos = 0;
                    string nomeRef = string.Empty;
                    if (ativa != null)
                    {
                        nomeRef = string.IsNullOrWhiteSpace(ativa.Aluno) ? (ativa.Professor ?? string.Empty) : (ativa.Aluno ?? string.Empty);
                        minutos = (int)Math.Floor((now - ativa.DataHora).TotalMinutes);
                    }
                    else if (reservadaHoje != null)
                    {
                        nomeRef = string.IsNullOrWhiteSpace(reservadaHoje.Aluno) ? (reservadaHoje.Professor ?? string.Empty) : (reservadaHoje.Aluno ?? string.Empty);
                        minutos = reservadaHoje.DataHora.Hour * 60 + reservadaHoje.DataHora.Minute;
                    }

                    sig.Append('|').Append(c.Nome).Append(':').Append(statusText).Append(':').Append(disponiveis).Append('/').Append(c.NumCopias).Append(':').Append(nomeRef).Append(':').Append(minutos);
                }

                // Acrescentar contadores para evitar rebuild quando iguais
                sig.Append("|tot:").Append(countDisp).Append(',').Append(countRes).Append(',').Append(countUso);

                var newSig = sig.ToString();
                if (string.Equals(newSig, _lastKeysSignature, StringComparison.Ordinal))
                {
                    // Nada relevante mudou: evita rebuild para não piscar
                    return;
                }

                _lastKeysSignature = newSig;

                // 2) Atualiza a UI incrementalmente somente quando a assinatura mudou
                _keysPanel.SuspendLayout();

                // Mapa de painéis existentes por chave
                var existing = _keysPanel.Controls
                    .OfType<Panel>()
                    .Where(p => p.Tag is string)
                    .ToDictionary(p => (string)p.Tag!, StringComparer.OrdinalIgnoreCase);

                var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var c in chaves)
                {
                    seen.Add(c.Nome);
                    // Determine status + detalhes
                    var reservasDaChave = reservas.Where(x => x.Chave == c.Nome).ToList();
                    var statusColor = Color.LightGreen; // Disponível
                    var statusText = "Disponível";

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
                            if (r.EmUso) { statusColor = Color.LightCoral; statusText = "Em Uso"; break; }
                            else if (r.DataHora.Date == nowDate) { statusColor = Color.Khaki; statusText = "Reservado"; }
                        }
                    }

                    // Card panel: reusa se existir, senão cria
                    int tileSize = 180;
                    Panel panel;
                    Label lbl;
                    Label status;
                    Label details;
                    Label disponibilidade;

                    if (existing.TryGetValue(c.Nome, out var reuse))
                    {
                        panel = reuse;
                        lbl = (Label)panel.Controls["titleLabel"]!;
                        status = (Label)panel.Controls["statusLabel"]!;
                        details = (Label)panel.Controls["detailsLabel"]!;
                        disponibilidade = (Label)panel.Controls["availLabel"]!;
                    }
                    else
                    {
                        panel = new MedControl.UI.SquareCardPanel
                        {
                            SizeHint = tileSize,
                            Width = tileSize,
                            Height = tileSize,
                            BackColor = GetCardColor(),
                            BorderStyle = BorderStyle.FixedSingle,
                            Padding = new Padding(6),
                            Margin = new Padding(10),
                            Tag = c.Nome
                        };

                        lbl = new Label
                        {
                            Name = "titleLabel",
                            Dock = DockStyle.Top,
                            Height = 24,
                            TextAlign = ContentAlignment.MiddleCenter,
                            BackColor = GetCardColor(),
                            Font = new Font("Segoe UI", 10, FontStyle.Bold),
                            Tag = "keep-backcolor keep-font"
                        };
                        lbl.AutoEllipsis = true;
                        lbl.AutoSize = false;

                        status = new Label
                        {
                            Name = "statusLabel",
                            Dock = DockStyle.Top,
                            Height = 24,
                            TextAlign = ContentAlignment.MiddleCenter,
                            Font = new Font("Segoe UI", 10, FontStyle.Bold)
                        };
                        status.Tag = "keep-backcolor keep-font";
                        details = new Label
                        {
                            Name = "detailsLabel",
                            Dock = DockStyle.Fill,
                            TextAlign = ContentAlignment.MiddleCenter,
                            BackColor = GetCardColor(),
                            Font = new Font("Segoe UI", 9, FontStyle.Regular),
                            Tag = "keep-backcolor keep-font"
                        };
                        details.AutoSize = false;
                        details.Padding = new Padding(2);
                        details.UseMnemonic = false;
                        details.AutoEllipsis = false;
                        details.MaximumSize = new Size(int.MaxValue, int.MaxValue);

                        disponibilidade = new Label
                        {
                            Name = "availLabel",
                            Dock = DockStyle.Bottom,
                            Height = 18,
                            TextAlign = ContentAlignment.MiddleCenter,
                            BackColor = GetCardColor(),
                            Font = new Font("Segoe UI", 8, FontStyle.Italic),
                            Tag = "keep-backcolor keep-font"
                        };

                        // Montagem
                        panel.Controls.Add(disponibilidade);
                        panel.Controls.Add(details);
                        panel.Controls.Add(status);
                        panel.Controls.Add(lbl);
                        _keysPanel.Controls.Add(panel);

                        // Clique abre relação
                        var chaveNome = c.Nome;
                        EventHandler open = (_, __) => { new Views.EntregasForm(chaveNome).ShowDialog(this); RefreshKeysUi(); };
                        void WireClick(Control ctrl) { ctrl.Cursor = Cursors.Hand; ctrl.Click += open; }
                        panel.Cursor = Cursors.Hand;
                        WireClick(panel); WireClick(lbl); WireClick(status); WireClick(details); WireClick(disponibilidade);
                    }

                    // details line: who and since when
                    string detailsText = string.Empty;
                    if ( ativa != null )
                    {
                        var nome = string.IsNullOrWhiteSpace(ativa.Aluno) ? ativa.Professor : ativa.Aluno;
                        var dur = now - ativa.DataHora;
                        string durStr = FormatDurationShort(dur);
                        detailsText = $"Com: {nome}\nDesde: {ativa.DataHora:dd/MM HH:mm:ss} • {durStr}";
                    }
                    else if ( reservadaHoje != null )
                    {
                        var nome = string.IsNullOrWhiteSpace(reservadaHoje.Aluno) ? reservadaHoje.Professor : reservadaHoje.Aluno;
                        detailsText = $"Reserva: {nome}\nHorário: {reservadaHoje.DataHora:HH:mm:ss}";
                    }

                    // Update contents (sem recriar)
                    lbl.Text = c.Nome;
                    status.Text = statusText;
                    status.BackColor = statusColor;
                    details.Text = detailsText;
                    // availability based on num_copias
                    int emUsoCount = reservasDaChave.Count(r => !r.Devolvido && r.EmUso);
                    int disponiveis = Math.Max(0, (c.NumCopias <= 0 ? 0 : c.NumCopias) - emUsoCount);
                    disponibilidade.Text = $"Disponíveis: {disponiveis} de {c.NumCopias}";

                    // Tooltip
                    var reservasText = string.Join("\n",
                        reservasDaChave.Select(r =>
                        {
                            var nome = string.IsNullOrWhiteSpace(r.Aluno) ? r.Professor : r.Aluno;
                            return $"{nome} - {r.DataHora:dd/MM/yyyy HH:mm:ss}";
                        }));
                    var statusInfo = $"Status: {statusText}" + (string.IsNullOrEmpty(reservasText) ? string.Empty : $"\nReservas:\n{reservasText}");
                    _toolTip.SetToolTip(status, statusInfo);
                    _toolTip.SetToolTip(panel, statusInfo);

                    // Ajuste tipográfico
                    FitCardTypography(panel, lbl, details);
                }

                // Remove painéis de chaves que não existem mais
                foreach (var ctrl in _keysPanel.Controls.OfType<Panel>().ToArray())
                {
                    if (ctrl.Tag is string key && !seen.Contains(key))
                    {
                        _keysPanel.Controls.Remove(ctrl);
                        try { ctrl.Dispose(); } catch { }
                    }
                }

                // update badges
                _badgeDisponiveis.Text = $"Disponíveis: {countDisp}";
                _badgeReservadas.Text = $"Reservadas: {countRes}";
                _badgeEmUso.Text = $"Em uso: {countUso}";

                _keysPanel.ResumeLayout();
            }
            catch { }
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

        private void UpdateConnectionStatus(bool immediate)
        {
            try
            {
                var mode = GroupConfig.Mode;
                if (mode == GroupMode.Client)
                {
                    SetStatusLabel(null, "Cliente");
                    // Non-blocking ping + RTT
                    Task.Run(() =>
                    {
                        bool ok; string? msg; int rtt;
                        try { ok = GroupClient.Ping(out msg, out rtt); }
                        catch (Exception ex) { ok = false; msg = ex.Message; rtt = -1; }
                        if (ok && rtt >= 0) _lastRttMs = rtt; // guarda último RTT válido
                        try
                        {
                            if (IsHandleCreated)
                            {
                                BeginInvoke(new Action(() =>
                                {
                                    string detail;
                                    if (ok)
                                    {
                                        // Exibe modo + RTT
                                        detail = _lastRttMs >= 0 ? $"Cliente – {_lastRttMs} ms" : "Cliente";
                                    }
                                    else
                                    {
                                        detail = "Offline" + (string.IsNullOrWhiteSpace(msg) ? string.Empty : " – " + msg);
                                    }
                                    SetStatusLabel(ok, detail);
                                }));
                            }
                        }
                        catch { }
                    });
                }
                else if (mode == GroupMode.Host)
                {
                    // Host local: considera RTT ~0ms
                    _lastRttMs = 0;
                    SetStatusLabel(true, "Host – 0 ms");
                }
                else
                {
                    SetStatusLabel(false, "Offline");
                }
            }
            catch { }
        }

        private void SetStatusLabel(bool? online, string? detail)
        {
            try
            {
                if (_statusDot == null || _statusText == null) return;
                if (online == null)
                {
                    _statusDot.Text = "●";
                    _statusDot.ForeColor = Color.Gray;
                    _statusDot.BackColor = Color.Transparent;
                    _statusText.Text = " verificando...";
                    _statusText.ForeColor = Color.Black;
                    _statusText.BackColor = Color.Transparent;
                    return;
                }
                var isOn = online.Value;
                _statusDot.Text = "●";
                _statusDot.ForeColor = isOn ? Color.LimeGreen : Color.Red;
                _statusDot.BackColor = Color.Transparent;
                _statusText.Text = $" {(isOn ? "Online" : "Offline")}" + (string.IsNullOrWhiteSpace(detail) ? string.Empty : $" ({detail})");
                _statusText.ForeColor = Color.Black;
                _statusText.BackColor = Color.Transparent;
            }
            catch { }
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Se o usuário clicou no X e não pediu saída real, apenas minimiza
            if (!_allowExit && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                WindowState = FormWindowState.Minimized;
                ShowInTaskbar = false; // oculta da barra; acessível pela bandeja
                return;
            }
            // Saída real
            try
            {
                Program.MarkShuttingDown($"FormClosing Reason={e.CloseReason}; AllowExit={_allowExit}");
                WriteHeartbeat("closing");
                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
                Instance = null; // limpa referência estática
            }
            catch { }
            base.OnFormClosing(e);
        }

        private void WriteHeartbeat(string state = "alive")
        {
            try
            {
                var logPath = Path.Combine(AppContext.BaseDirectory, "AppErrors.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Heartbeat: {state}; Mode={GroupConfig.Mode}; RTT={_lastRttMs}\n");
            }
            catch { }
        }

        private void InitTrayIcon()
        {
            try
            {
                if (_trayIcon != null) return; // já inicializado
                _trayMenu = new ContextMenuStrip();
                _trayMenu.Items.Add("Abrir", null, (_, __) => RestoreFromTray());
                _trayMenu.Items.Add("Sair", null, (_, __) => { _allowExit = true; Close(); });
                _trayIcon = new NotifyIcon
                {
                    Icon = this.Icon ?? SystemIcons.Application,
                    Text = "MedControl",
                    Visible = true,
                    ContextMenuStrip = _trayMenu
                };
                _trayIcon.DoubleClick += (_, __) => RestoreFromTray();
            }
            catch { }
        }

        private void UpdateTrayIcon()
        {
            try
            {
                if (_trayIcon != null && this.Icon != null)
                {
                    _trayIcon.Icon = this.Icon; // garante que sejam iguais
                }
            }
            catch { }
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            UpdateTrayIcon();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            try
            {
                if (WindowState == FormWindowState.Minimized && !_allowExit)
                {
                    ShowInTaskbar = false;
                    if (_trayIcon != null) _trayIcon.Visible = true;
                }
            }
            catch { }
        }

        private void RestoreFromTray()
        {
            try
            {
                ShowInTaskbar = true;
                if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
                Activate();
            }
            catch { }
        }

        public static void RequestRestore() => Instance?.RestoreFromTray();

    }
}
