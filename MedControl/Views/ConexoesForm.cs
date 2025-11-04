using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class ConexoesForm : Form
    {
        // Config UI
        private ComboBox _groupMode;
        private TextBox _groupName;
        private TextBox _groupHost;
        private TextBox _groupPort;
    private TextBox _nodeAlias;
        private Button _testConn;
        private Button _saveBtn;
        private ListView _peersList;
        private RichTextBox _chatLog;
        private TextBox _messageBox;
        private Button _sendBtn;
    private Button _emojiBtn;
    private ContextMenuStrip _emojiMenu;
        private Label _lblMode;
        private Label _lblGroup;
        private Label _lblHost;
    private Label _lblSelf;
    private Label _lblStatus;
        private System.Windows.Forms.Timer _uiTimer;
    private System.Windows.Forms.Timer? _statusTimer;
    private Panel _loadingPanel;
    private ProgressBar _loadingBar;
    private Label _loadingLabel;
    private volatile bool _pendingPeersUpdate;
    private bool _firstPeersLoad = true;

        public ConexoesForm()
        {
            Text = "Conexões e Chat";
            StartPosition = FormStartPosition.CenterParent;
            Width = 980;
            Height = 720;
            MinimumSize = new Size(900, 600);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(12),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 65));

            // Info header
            var infoPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 4,
                RowCount = 2,
                AutoSize = true,
            };
            infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            _lblMode = MakeInfoLabel("Modo: ...");
            _lblGroup = MakeInfoLabel("Grupo: ...");
            _lblHost = MakeInfoLabel("Host/Porta: ...");
            _lblSelf = MakeInfoLabel($"Este nó: {MedControl.SyncService.LocalNodeName()}");
            infoPanel.Controls.Add(_lblMode, 0, 0);
            infoPanel.Controls.Add(_lblGroup, 1, 0);
            infoPanel.Controls.Add(_lblHost, 2, 0);
            infoPanel.Controls.Add(_lblSelf, 3, 0);
            _lblStatus = MakeInfoLabel("Status: verificando...");
            infoPanel.Controls.Add(_lblStatus, 0, 1);
            try { infoPanel.SetColumnSpan(_lblStatus, 4); } catch { }

            // Config section
            var cfg = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true, Padding = new Padding(0, 4, 0, 6) };
            cfg.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            cfg.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            cfg.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            Label Mk(string t) => new Label { Text = t, AutoSize = true, Anchor = AnchorStyles.Right, Margin = new Padding(3, 8, 12, 8) };
            _groupMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
            _groupMode.Items.AddRange(new object[] { "Offline", "Host", "Cliente" });
            _groupName = new TextBox { Dock = DockStyle.Fill };
            _groupHost = new TextBox { Dock = DockStyle.Fill };
            _groupPort = new TextBox { Dock = DockStyle.Fill };
            _nodeAlias = new TextBox { Dock = DockStyle.Fill };
            _testConn = new Button { Text = "Testar Conexão", AutoSize = true };
            _saveBtn = new Button { Text = "Salvar", AutoSize = true };

            cfg.Controls.Add(Mk("Modo de Grupo:"), 0, 0);
            cfg.Controls.Add(_groupMode, 1, 0);
            cfg.Controls.Add(new Label { Text = " " }, 2, 0);

            cfg.Controls.Add(Mk("Nome do Grupo:"), 0, 1);
            cfg.Controls.Add(_groupName, 1, 1);
            cfg.Controls.Add(new Label { Text = "(ex.: lab1)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 1);

            cfg.Controls.Add(Mk("Host (Cliente → host:porta):"), 0, 2);
            cfg.Controls.Add(_groupHost, 1, 2);
            cfg.Controls.Add(_testConn, 2, 2);

            cfg.Controls.Add(Mk("Porta do Host:"), 0, 3);
            cfg.Controls.Add(_groupPort, 1, 3);
            cfg.Controls.Add(new Label { Text = "(padrão 49383)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 3);

            cfg.Controls.Add(Mk("Apelido deste dispositivo:"), 0, 4);
            cfg.Controls.Add(_nodeAlias, 1, 4);
            cfg.Controls.Add(new Label { Text = "(opcional p/ testar em 1 PC)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 4);

            cfg.Controls.Add(new Label { Text = " " }, 0, 5);
            cfg.Controls.Add(new Label { Text = " " }, 1, 5);
            cfg.Controls.Add(_saveBtn, 2, 5);

            // Peers list
            _peersList = new ListView
            {
                View = View.Details,
                FullRowSelect = true,
                Dock = DockStyle.Fill,
            };
            _peersList.Columns.Add("Usuário", 180);
            _peersList.Columns.Add("Endereço", 160);
            _peersList.Columns.Add("Dica", 260);
            _peersList.Columns.Add("Visto", 140);

            var peersGroup = new GroupBox { Text = "Usuários próximos", Dock = DockStyle.Fill };
            // Loading panel (top)
            _loadingPanel = new Panel { Dock = DockStyle.Top, Height = 32, Padding = new Padding(6), Visible = true };
            _loadingBar = new ProgressBar { Style = ProgressBarStyle.Marquee, Dock = DockStyle.Right, Width = 180, MarqueeAnimationSpeed = 30 };
            _loadingLabel = new Label { Text = "Carregando...", AutoSize = true, Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft };
            _loadingPanel.Controls.Add(_loadingBar);
            _loadingPanel.Controls.Add(_loadingLabel);
            peersGroup.Controls.Add(_peersList);
            peersGroup.Controls.Add(_loadingPanel);

            // Chat
            var chatGroup = new GroupBox { Text = "Chat do grupo", Dock = DockStyle.Fill };
            var chatLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            chatLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            chatLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _chatLog = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, DetectUrls = true, BackColor = Color.White };
            // Prefer an emoji-capable font
            try { _chatLog.Font = new Font("Segoe UI Emoji", 10f, FontStyle.Regular); } catch { }

            var inputPanel = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true };
            inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _messageBox = new TextBox { Dock = DockStyle.Fill };
            try { _messageBox.Font = new Font("Segoe UI Emoji", 10f, FontStyle.Regular); } catch { }
            _messageBox.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; SendChat(); } };
            _emojiBtn = new Button { Text = "😊", AutoSize = true, Dock = DockStyle.Right, Margin = new Padding(6, 0, 6, 0) };
            _emojiMenu = BuildEmojiMenu();
            _emojiBtn.Click += (s, e) =>
            {
                try
                {
                    var pt = new Point(0, _emojiBtn.Height);
                    _emojiMenu.Show(_emojiBtn, pt);
                }
                catch { }
            };
            _sendBtn = new Button { Text = "Enviar", AutoSize = true, Dock = DockStyle.Right };
            _sendBtn.Click += (_, __) => SendChat();
            inputPanel.Controls.Add(_messageBox, 0, 0);
            inputPanel.Controls.Add(_emojiBtn, 1, 0);
            inputPanel.Controls.Add(_sendBtn, 2, 0);
            chatLayout.Controls.Add(_chatLog, 0, 0);
            chatLayout.Controls.Add(inputPanel, 0, 1);
            chatGroup.Controls.Add(chatLayout);

            root.Controls.Add(infoPanel);
            root.Controls.Add(cfg);
            root.Controls.Add(peersGroup);
            root.Controls.Add(chatGroup);
            Controls.Add(root);

            Load += (_, __) => OnLoaded();
            FormClosed += (_, __) => OnClosed();

            // UI refresh timer to update peers timestamp smoothly
            _uiTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            _uiTimer.Tick += (_, __) =>
            {
                try
                {
                    // Atualiza visto e aplica refresh apenas quando necessário ou no primeiro load
                    if (_firstPeersLoad || _pendingPeersUpdate)
                    {
                        RefreshPeers();
                        _pendingPeersUpdate = false;
                    }
                }
                catch { }
            };

            try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
        }

        private Label MakeInfoLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                Margin = new Padding(0, 0, 16, 8)
            };
        }

        private void OnLoaded()
        {
            // Ensure sync service is running
            try { SyncService.Start(); } catch { }

            // Fill header
            try
            {
                var mode = ModeDisplay(GroupConfig.Mode);
                var group = GroupConfig.GroupName;
                var host = GroupConfig.Mode == GroupMode.Host ? $"localhost:{GroupConfig.HostPort}" : GroupConfig.HostAddress;
                if (string.IsNullOrWhiteSpace(host)) host = $"porta {GroupConfig.HostPort}";
                _lblMode.Text = $"Modo: {mode}";
                _lblGroup.Text = $"Grupo: {group}";
                _lblHost.Text = $"Host/Porta: {host}";
            }
            catch { }

            // Init config fields
            try
            {
                var gm = GroupConfig.Mode;
                _groupMode.SelectedIndex = gm == GroupMode.Host ? 1 : gm == GroupMode.Client ? 2 : 0;
                _groupName.Text = GroupConfig.GroupName;
                _groupHost.Text = GroupConfig.HostAddress;
                _groupPort.Text = GroupConfig.HostPort.ToString();
                _nodeAlias.Text = Database.GetConfig("node_alias") ?? string.Empty;
            }
            catch { }

            _groupMode.SelectedIndexChanged += (_, __) => { ToggleGroupUi(); UpdateStatusTimerMode(); };
            _testConn.Click += (_, __) => DoTestConn();
            _saveBtn.Click += (_, __) => DoSave();
            
            void DoSave()
            {
                try
                {
                    var m = _groupMode.SelectedIndex switch { 1 => GroupMode.Host, 2 => GroupMode.Client, _ => GroupMode.Solo };
                    GroupConfig.Mode = m;
                    GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                    if (int.TryParse(_groupPort.Text?.Trim(), out var port) && port > 0) GroupConfig.HostPort = port;
                    GroupConfig.HostAddress = _groupHost.Text?.Trim() ?? string.Empty;
                    var alias = _nodeAlias.Text?.Trim() ?? string.Empty;
                    Database.SetConfig("node_alias", alias);
                    try { if (m == GroupMode.Host) GroupHost.Start(); else GroupHost.Stop(); } catch { }

                    // Update header labels
                    _lblMode.Text = $"Modo: {ModeDisplay(GroupConfig.Mode)}";
                    _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                    var host = GroupConfig.Mode == GroupMode.Host ? $"localhost:{GroupConfig.HostPort}" : GroupConfig.HostAddress;
                    if (string.IsNullOrWhiteSpace(host)) host = $"porta {GroupConfig.HostPort}";
                    _lblHost.Text = $"Host/Porta: {host}";
                    _lblSelf.Text = $"Este nó: {MedControl.SyncService.LocalNodeName()}";

                    MessageBox.Show(this, "Conexões salvas.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Restart status/ping policy based on new mode
                    UpdateStatusTimerMode();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Erro ao salvar conexões: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            ToggleGroupUi();
            UpdateStatusTimerMode();

            // Subscribe events
            SyncService.PeersChanged += SyncService_PeersChanged;
            SyncService.OnChat += SyncService_OnChat;

            RefreshPeers();
            _uiTimer.Start();
        }

        private void OnClosed()
        {
            try { SyncService.PeersChanged -= SyncService_PeersChanged; } catch { }
            try { SyncService.OnChat -= SyncService_OnChat; } catch { }
            try { _uiTimer?.Stop(); } catch { }
            try { _statusTimer?.Stop(); } catch { }
        }

        private void SyncService_PeersChanged()
        {
            try { _pendingPeersUpdate = true; } catch { }
        }

        private void ToggleGroupUi()
        {
            try
            {
                var isClient = _groupMode.SelectedIndex == 2;
                _groupHost.Enabled = isClient;
                _groupPort.Enabled = isClient;
                _testConn.Enabled = isClient;
            }
            catch { }
        }

        private void DoTestConn()
        {
            try
            {
                if (_groupMode.SelectedIndex != 2)
                {
                    MessageBox.Show(this, "Mude o modo para Cliente para testar.");
                    return;
                }
                if (int.TryParse(_groupPort.Text?.Trim(), out var p) && p > 0) GroupConfig.HostPort = p;
                GroupConfig.HostAddress = _groupHost.Text?.Trim() ?? string.Empty;
                var ok = GroupClient.Ping(out var msg);
                MessageBox.Show(this, ok ? ($"Conectado ao Host. Resposta: {msg}") : ($"Falha: {msg}"));
                UpdateStatusBadge(ok, ok ? "" : msg ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao testar conexão: " + ex.Message);
            }
        }

        private void DoSave()
        {
            try
            {
                var m = _groupMode.SelectedIndex switch { 1 => GroupMode.Host, 2 => GroupMode.Client, _ => GroupMode.Solo };
                GroupConfig.Mode = m;
                GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                if (int.TryParse(_groupPort.Text?.Trim(), out var port) && port > 0) GroupConfig.HostPort = port;
                GroupConfig.HostAddress = _groupHost.Text?.Trim() ?? string.Empty;
                try { if (m == GroupMode.Host) GroupHost.Start(); else GroupHost.Stop(); } catch { }

                // Update header labels
                _lblMode.Text = $"Modo: {ModeDisplay(GroupConfig.Mode)}";
                _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                var host = GroupConfig.Mode == GroupMode.Host ? $"localhost:{GroupConfig.HostPort}" : GroupConfig.HostAddress;
                if (string.IsNullOrWhiteSpace(host)) host = $"porta {GroupConfig.HostPort}";
                _lblHost.Text = $"Host/Porta: {host}";

                MessageBox.Show(this, "Conexões salvas.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Restart status/ping policy based on new mode
                UpdateStatusTimerMode();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao salvar conexões: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshPeers()
        {
            try
            {
                var peers = SyncService.GetPeers();
                _peersList.BeginUpdate();
                _peersList.Items.Clear();
                foreach (var p in peers)
                {
                    var item = new ListViewItem(new[]
                    {
                        p.Node,
                        p.Address,
                        string.IsNullOrWhiteSpace(p.HostHint)? "" : p.HostHint,
                        p.LastSeen.ToLocalTime().ToString("HH:mm:ss")
                    });
                    _peersList.Items.Add(item);
                }
                _peersList.EndUpdate();
                // Hide loading once first refresh completes
                if (_firstPeersLoad)
                {
                    _firstPeersLoad = false;
                    try { _loadingPanel.Visible = false; } catch { }
                }
            }
            catch { }
        }

        private void UpdateStatusTimerMode()
        {
            try
            {
                _statusTimer ??= new System.Windows.Forms.Timer { Interval = 7000 };
                _statusTimer.Stop();
                _statusTimer.Tick -= StatusTimer_Tick;
                _statusTimer.Tick += StatusTimer_Tick;

                var mode = GroupConfig.Mode;
                if (mode == GroupMode.Client)
                {
                    UpdateStatusBadge(null); // show checking
                    _statusTimer.Start();
                    // Trigger an immediate ping
                    TriggerPing();
                }
                else if (mode == GroupMode.Host)
                {
                    _statusTimer.Stop();
                    UpdateStatusBadge(true, "Host ativo");
                }
                else // Offline
                {
                    _statusTimer.Stop();
                    UpdateStatusBadge(false, "Modo Offline");
                }
            }
            catch { }
        }

        private void StatusTimer_Tick(object? sender, EventArgs e)
        {
            try { TriggerPing(); } catch { }
        }

        private void TriggerPing()
        {
            try
            {
                // Don’t block UI
                Task.Run(() =>
                {
                    bool ok;
                    string? msg;
                    try { ok = GroupClient.Ping(out msg); }
                    catch (Exception ex) { ok = false; msg = ex.Message; }
                    try
                    {
                        if (IsHandleCreated)
                            BeginInvoke(new Action(() => UpdateStatusBadge(ok, ok ? "" : msg ?? "")));
                    }
                    catch { }
                });
            }
            catch { }
        }

        private void UpdateStatusBadge(bool? online, string? detail = null)
        {
            try
            {
                if (online == null)
                {
                    _lblStatus.Text = "Status: verificando...";
                    _lblStatus.ForeColor = Color.FromArgb(64, 64, 64);
                    _lblStatus.BackColor = Color.Transparent;
                    return;
                }
                var isOn = online.Value;
                var dot = isOn ? "●" : "●";
                var color = isOn ? Color.LightGreen : Color.LightCoral;
                var text = isOn ? "Online" : "Offline";
                if (!string.IsNullOrWhiteSpace(detail)) text += $" – {detail}";
                _lblStatus.Text = $"Status: {dot} {text}";
                _lblStatus.BackColor = color;
                _lblStatus.ForeColor = Color.Black;
                _lblStatus.Tag = "keep-backcolor keep-font";
            }
            catch { }
        }

        private void SyncService_OnChat(string sender, string message, DateTime utc, string address)
        {
            try
            {
                if (IsHandleCreated)
                {
                    BeginInvoke(new Action(() =>
                    {
                        AppendChatLine(sender, message, utc.ToLocalTime(), address);
                    }));
                }
            }
            catch { }
        }

        private static string ModeDisplay(GroupMode mode)
        {
            return mode switch
            {
                GroupMode.Solo => "Offline",
                GroupMode.Host => "Host",
                GroupMode.Client => "Cliente",
                _ => mode.ToString()
            };
        }

        private void AppendChatLine(string sender, string message, DateTime localTime, string address)
        {
            try
            {
                var time = localTime.ToString("HH:mm:ss");
                _chatLog.SelectionColor = Color.DimGray;
                _chatLog.AppendText($"[{time}] ");
                _chatLog.SelectionColor = Color.DarkBlue;
                _chatLog.AppendText(sender);
                _chatLog.SelectionColor = Color.DimGray;
                _chatLog.AppendText($" ({address})");
                _chatLog.SelectionColor = Color.Black;
                _chatLog.AppendText($": {message}\n");
                _chatLog.ScrollToCaret();
            }
            catch { }
        }

        private void SendChat()
        {
            try
            {
                var text = _messageBox.Text?.Trim();
                if (string.IsNullOrEmpty(text)) return;
                _messageBox.Clear();
                SyncService.SendChat(text);
                AppendChatLine("Você", text, DateTime.Now, "local");
            }
            catch { }
        }

        private ContextMenuStrip BuildEmojiMenu()
        {
            var menu = new ContextMenuStrip();
            string[] emojis = new[]
            {
                "😀","😂","😊","😉","😍","😘","😎","🤩","🤔","🙄","😴","😢","😭","😡","😅","🤗","👍","👏","🙏","👌","🤝","🎉","🔥","💯","✅","❌","⚠️","❤️","💔","✨","📌","📎","📣","📷","📝","🔒","🔑","🕒","📅"
            };
            foreach (var e in emojis)
            {
                var item = new ToolStripMenuItem(e) { Font = new Font("Segoe UI Emoji", 12f, FontStyle.Regular) };
                item.Click += (_, __) => InsertEmoji(e);
                menu.Items.Add(item);
            }
            return menu;
        }

        private void InsertEmoji(string emoji)
        {
            try
            {
                _messageBox.Focus();
                var selStart = _messageBox.SelectionStart;
                var txt = _messageBox.Text ?? string.Empty;
                _messageBox.Text = txt.Substring(0, selStart) + emoji + txt.Substring(selStart);
                _messageBox.SelectionStart = selStart + emoji.Length;
            }
            catch { }
        }
    }
}
