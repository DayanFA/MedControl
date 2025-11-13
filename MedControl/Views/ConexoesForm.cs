using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MedControl.Views
{
    public class ConexoesForm : Form
    {
        // Config UI
        private ComboBox _groupMode;
        private TextBox _groupName;
        private ComboBox _groupSelect; // lista de grupos detectados
        private TextBox _groupHost;
        private TextBox _groupPort;
        private TextBox _groupPassword;
    private Button _testConn;
        private Button _saveBtn;
        private Button _connectBtn;
    private Button _createGroupBtn;
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
    // Config labels to toggle visibility
    private Label _cfgLblHost;
    private Label _cfgLblPort;
    private Label _cfgLblPwd;
    private Label _cfgLblGroupName;
    private Label _cfgLblGroupSelect;
        private System.Windows.Forms.Timer _uiTimer;
        private System.Windows.Forms.Timer? _statusTimer;
        private Panel _loadingPanel;
        private ProgressBar _loadingBar;
        private Label _loadingLabel;
        private volatile bool _pendingPeersUpdate;
        private bool _firstPeersLoad = true;
        private bool _autoReconnectInProgress = false;
        private bool _wasOnline = false;

        public ConexoesForm()
        {
            Text = "Conexões";
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
            _groupPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            _groupSelect = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _testConn = new Button { Text = "Testar Conexão", AutoSize = true, Visible = false, Enabled = false };
            _saveBtn = new Button { Text = "Salvar e Conectar", AutoSize = true };
            _connectBtn = new Button { Text = "Conectar", AutoSize = true, Visible = false, Enabled = false };
            _createGroupBtn = new Button { Text = "Criar grupo (Host)", AutoSize = true };

            cfg.Controls.Add(Mk("Modo de Grupo:"), 0, 0);
            cfg.Controls.Add(_groupMode, 1, 0);
            cfg.Controls.Add(new Label { Text = " " }, 2, 0);

            _cfgLblGroupName = Mk("Nome do Grupo:");
            cfg.Controls.Add(_cfgLblGroupName, 0, 1);
            cfg.Controls.Add(_groupName, 1, 1);
            cfg.Controls.Add(new Label { Text = "(ex.: lab1)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 1);

            _cfgLblGroupSelect = Mk("Grupos próximos:");
            cfg.Controls.Add(_cfgLblGroupSelect, 0, 2);
            cfg.Controls.Add(_groupSelect, 1, 2);
            cfg.Controls.Add(new Label { Text = "(auto)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 2);

            _cfgLblHost = Mk("Host (modo Host):");
            cfg.Controls.Add(_cfgLblHost, 0, 3);
            cfg.Controls.Add(_groupHost, 1, 3);
            cfg.Controls.Add(new Label { Text = " ", AutoSize = true }, 2, 3); // espaço: botão de teste removido

            _cfgLblPort = Mk("Porta do Host:");
            cfg.Controls.Add(_cfgLblPort, 0, 4);
            cfg.Controls.Add(_groupPort, 1, 4);
            cfg.Controls.Add(new Label { Text = "(padrão 49383)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 4);

            _cfgLblPwd = Mk("Senha do Grupo:");
            cfg.Controls.Add(_cfgLblPwd, 0, 5);
            cfg.Controls.Add(_groupPassword, 1, 5);
            cfg.Controls.Add(new Label { Text = "(opcional)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 5);

            // Removido campo de apelido do dispositivo a pedido do usuário

            // Botão de criar grupo/ser Host agora
            cfg.Controls.Add(_createGroupBtn, 2, 6);

            cfg.Controls.Add(new Label { Text = " " }, 0, 7);
            cfg.Controls.Add(new Label { Text = " " }, 1, 7); // espaço onde ficava Conectar
            cfg.Controls.Add(_saveBtn, 2, 7);

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
            var chatGroup = new GroupBox { Text = "Chat do grupo", Dock = DockStyle.Fill }; // mantém título
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

        private string _chatHistoryPath = string.Empty;

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
                _groupPassword.Text = GroupConfig.GroupPassword;
                // Campo de apelido removido; mantemos configuração existente sem expor no UI
            }
            catch { }

            _groupMode.SelectedIndexChanged += (_, __) => { ToggleGroupUi(); UpdateStatusTimerMode(); };
            _saveBtn.Click += (_, __) => DoSave();
            _createGroupBtn.Click += (_, __) => DoCreateGroup();
            _groupSelect.DropDown += (_, __) => RefreshGroupsList();
            _groupHost.TextChanged += (_, __) => ToggleGroupUi();
            _groupPassword.TextChanged += (_, __) => ToggleGroupUi();
            
            void DoSave()
            {
                try
                {
                    var m = _groupMode.SelectedIndex switch { 1 => GroupMode.Host, 2 => GroupMode.Client, _ => GroupMode.Solo };
                    // Atualiza config base conforme modo desejado
                    if (m == GroupMode.Host)
                    {
                        GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                        GroupConfig.GroupPassword = _groupPassword.Text ?? string.Empty;
                    }
                    else if (m == GroupMode.Client)
                    {
                        // Cliente: somente IP (host) e senha; porta padrão
                        ApplyHostPortFromFields();
                        GroupConfig.GroupPassword = _groupPassword.Text ?? string.Empty;
                        // Nome do grupo pode ser descoberto via Hello mais adiante
                        if (string.IsNullOrWhiteSpace(GroupConfig.GroupName)) GroupConfig.GroupName = "default";
                    }
                    else
                    {
                        // Offline
                        GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                        GroupConfig.GroupPassword = _groupPassword.Text ?? string.Empty;
                    }
                    GroupConfig.Mode = m;

                    if (m == GroupMode.Host)
                    {
                        try { GroupHost.Start(); } catch { }
                        FinalizeUi("Conexões salvas.");
                        try { _saveBtn.Enabled = false; } catch { }
                        return;
                    }
                    if (m == GroupMode.Solo)
                    {
                        try { GroupHost.Stop(); } catch { }
                        FinalizeUi("Modo Offline salvo.");
                        try { _saveBtn.Enabled = false; } catch { }
                        return;
                    }

                    // Modo Cliente: fazer Hello (descobrir grupo) + Ping + Sync
                    using var dlg = new SyncProgressForm();
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.Show(this);
                    dlg.Refresh();
                    try { _saveBtn.Enabled = false; } catch { }

                    Task.Run(() =>
                    {
                        try
                        {
                            dlg.BeginInvoke(new Action(() => dlg.SetProgress(5, "Preparando...")));
                            // Força beacon (não obrigatório, mas mantém consistência)
                            try { SyncService.ForceBeacon(); } catch { }
                            // Tenta Hello se endereço informado para descobrir nome do grupo/porta
                            if (!string.IsNullOrWhiteSpace(GroupConfig.HostAddress))
                            {
                                if (GroupClient.Hello(out var gname, out var hport, out var herr))
                                {
                                    GroupConfig.GroupName = gname ?? "default";
                                    GroupConfig.HostPort = hport;
                                    try { BeginInvoke(new Action(() => _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}")); } catch { }
                                }
                            }
                            dlg.BeginInvoke(new Action(() => dlg.SetProgress(12, "Verificando host...")));
                            if (!GroupClient.Ping(out var msg, connectTimeoutMs: 1500, ioTimeoutMs: 2000))
                                throw new Exception(string.IsNullOrWhiteSpace(msg) ? "Host indisponível" : msg);

                            int step = 0; int total = 5;
                            string syncLogPath = Path.Combine(AppContext.BaseDirectory, "sync_download.log");
                            void LogSync(string dataset, int count)
                            {
                                try { File.AppendAllText(syncLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {dataset}: {count} registros baixados\n"); } catch { }
                            }
                            void Step(string label, int count)
                            {
                                step++;
                                var percent = Math.Min(100, Math.Max(0, (int)Math.Round(step * 100.0 / total)));
                                var display = count >= 0 ? $"{label} (#{count})" : label;
                                dlg.BeginInvoke(new Action(() => dlg.SetProgress(percent, display)));
                                LogSync(label, count);
                            }
                            // Sincronização
                            var chaves = GroupClient.PullChaves(); Database.ReplaceAllChavesLocal(chaves); Step("Chaves", chaves.Count);
                            var reservas = GroupClient.PullReservas(); Database.ReplaceAllReservas(reservas); Step("Reservas", reservas.Count);
                            var rels = GroupClient.PullRelatorio(); Database.ReplaceAllRelatorios(rels); Step("Relatórios", rels.Count);
                            var alunos = GroupClient.PullAlunos(); Database.ReplaceAllAlunos(alunos); Step("Alunos", alunos.Rows.Count);
                            var profs = GroupClient.PullProfessores(); Database.ReplaceAllProfessores(profs); Step("Professores", profs.Rows.Count);

                            dlg.BeginInvoke(new Action(() => dlg.Close()));
                            BeginInvoke(new Action(() =>
                            {
                                FinalizeUi("Conectado e sincronizado.");
                                try { MedControl.SyncService.ForceBeacon(); } catch { }
                                try { MedControl.SyncService.TryAddOrUpdateSelfPeer(); } catch { }
                                try { _saveBtn.Enabled = true; } catch { }
                            }));
                        }
                        catch (Exception ex)
                        {
                            try { dlg.BeginInvoke(new Action(() => dlg.Close())); } catch { }
                            BeginInvoke(new Action(() => FinalizeUi("Falha: " + ex.Message, isError:true)));
                        }
                    });

                    // Loop de mensagens até terminar
                    while (dlg.Visible)
                    {
                        Application.DoEvents();
                        System.Threading.Thread.Sleep(50);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Erro: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            void FinalizeUi(string message, bool isError = false)
            {
                try
                {
                    _lblMode.Text = $"Modo: {ModeDisplay(GroupConfig.Mode)}";
                    _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                    var host = GroupConfig.Mode == GroupMode.Host ? $"localhost:{GroupConfig.HostPort}" : GroupConfig.HostAddress;
                    if (string.IsNullOrWhiteSpace(host)) host = $"porta {GroupConfig.HostPort}";
                    _lblHost.Text = $"Host/Porta: {host}";
                    _lblSelf.Text = $"Este nó: {MedControl.SyncService.LocalNodeName()}";
                    UpdateStatusTimerMode();
                    if (!isError) MessageBox.Show(this, message, "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else MessageBox.Show(this, message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            ToggleGroupUi();
            RefreshGroupsList();
            UpdateStatusTimerMode();

            // Define caminho do histórico de chat (por grupo)
            try
            {
                var safeGroup = (GroupConfig.GroupName ?? "default").Trim();
                foreach (var c in System.IO.Path.GetInvalidFileNameChars()) safeGroup = safeGroup.Replace(c, '_');
                _chatHistoryPath = System.IO.Path.Combine(AppContext.BaseDirectory, $"chat_{safeGroup}.log");
                LoadChatHistory();
            }
            catch { }

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
                var isHost = _groupMode.SelectedIndex == 1;
                // Visibilidade por modo
                // Host: só Nome do Grupo e Senha
                _groupName.Visible = isHost;
                if (_cfgLblGroupName != null) _cfgLblGroupName.Visible = isHost;
                _groupHost.Visible = isClient;
                _cfgLblHost.Visible = isClient;
                _groupPort.Visible = false;
                _cfgLblPort.Visible = false;
                _groupSelect.Visible = false;
                if (_cfgLblGroupSelect != null) _cfgLblGroupSelect.Visible = false;
                try { if (isClient) _cfgLblHost.Text = "Host/IP do servidor:"; } catch { }
                // Senha visível em Host e Cliente
                _groupPassword.Visible = true;
                _cfgLblPwd.Visible = true;
                _testConn.Enabled = isClient;
                _connectBtn.Enabled = isClient;
                _createGroupBtn.Enabled = isHost; // criar grupo só faz sentido em Host

                // Habilitação do botão Salvar
                if (isClient)
                {
                    var hasHost = !string.IsNullOrWhiteSpace(_groupHost.Text);
                    var hasPwd = !string.IsNullOrWhiteSpace(_groupPassword.Text);
                    _saveBtn.Enabled = hasHost && hasPwd;
                }
                else
                {
                    _saveBtn.Enabled = false;
                }
            }
            catch { }
        }

        private void RefreshGroupsList()
        {
            try
            {
                var adverts = SyncService.GetGroupAdverts();
                var curGroup = GroupConfig.GroupName ?? "default";
                var curHost = GroupConfig.HostAddress ?? string.Empty;
                var curPort = GroupConfig.HostPort;
                var localNode = SyncService.LocalNodeName();
                // Apenas hosts (role=host) e ignora este nó caso seja host
                var hostAdverts = adverts
                    .Where(a => string.Equals(a.Role, "host", StringComparison.OrdinalIgnoreCase) && !string.Equals(a.Node, localNode, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var items = hostAdverts.Select(a => new GroupItem
                {
                    Group = a.Group,
                    Address = a.Address,
                    Port = a.HostPort,
                    Role = a.Role,
                    Node = a.Node
                }).ToList();
                // Garante que pelo menos o grupo atual aparece (mesmo sem anúncios)
                // Não adiciona fallback; se nenhum host foi encontrado deixa lista vazia
                bool emptyHosts = !items.Any();

                _groupSelect.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        _groupSelect.Items.Clear();
                        foreach (var it in items) _groupSelect.Items.Add(it);
                        // Seleciona o matching do host atual, senão o matching de grupo, senão primeiro
                        GroupItem? sel = null;
                        sel = items.FirstOrDefault(i => !string.IsNullOrWhiteSpace(i.Address) && string.Equals(i.Address, curHost, StringComparison.OrdinalIgnoreCase) && i.Port == curPort)
                              ?? items.FirstOrDefault(i => string.Equals(i.Group, curGroup, StringComparison.OrdinalIgnoreCase))
                              ?? null;
                        if (sel != null) _groupSelect.SelectedItem = sel; else _groupSelect.SelectedIndex = -1;
                        // Desabilita se não há hosts
                        _groupSelect.Enabled = !emptyHosts;
                    }
                    catch { }
                }));
            }
            catch { }
        }

        private class GroupItem
        {
            public string Group { get; set; } = "default";
            public string Address { get; set; } = string.Empty;
            public int Port { get; set; } = 0;
            public string Role { get; set; } = string.Empty;
            public string Node { get; set; } = string.Empty;
            public override string ToString()
            {
                var addr = !string.IsNullOrWhiteSpace(Address) && Port > 0 ? $"{Address}:{Port}" : (Port > 0 ? $"porta {Port}" : "");
                var nodePart = string.IsNullOrWhiteSpace(Node) ? "" : $" [{Node}]";
                var rolePart = string.Equals(Role, "host", StringComparison.OrdinalIgnoreCase) ? " (host)" : (Role == "(sem hosts)" ? "" : "");
                if (!string.IsNullOrWhiteSpace(addr)) return $"{Group} — {addr}{nodePart}{rolePart}";
                return $"{Group}{nodePart}{rolePart}";
            }
        }

        private void DoConnect() => DoConnect(false);

        private void DoConnect(bool auto)
        {
            try
            {
                if (_groupMode.SelectedIndex != 2)
                {
                    MessageBox.Show(this, "Selecione o modo Cliente antes de conectar.");
                    return;
                }

                // Exigir senha do grupo para Cliente
                if (string.IsNullOrWhiteSpace(_groupPassword.Text))
                {
                    MessageBox.Show(this, "Informe a senha do grupo para conectar como Cliente.", "Senha obrigatória", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Aplica senha ao config (nome pode ser capturado automaticamente)
                GroupConfig.GroupPassword = _groupPassword.Text ?? string.Empty;

                // Aplica temporariamente host/porta informados
                ApplyHostPortFromFields();

                // Se houver Host informado, tenta descobrir automaticamente o nome do grupo/porta
                if (!string.IsNullOrWhiteSpace(GroupConfig.HostAddress))
                {
                    if (GroupClient.Hello(out var gname, out var hport, out var herr))
                    {
                        GroupConfig.GroupName = gname ?? "default";
                        GroupConfig.HostPort = hport;
                        _groupName.Text = GroupConfig.GroupName;
                        _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                    }
                }

                // Se ainda não houver nome, usa o campo ou default
                GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) || string.Equals(_groupName.Text.Trim(), "(automático)", StringComparison.OrdinalIgnoreCase)
                    ? (GroupConfig.GroupName ?? "default")
                    : _groupName.Text.Trim();

                using var dlg = new SyncProgressForm();
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.Show(this);
                dlg.Refresh();

                var total = 5; // chaves, reservas, relatórios, alunos, profs
                int step = 0;

                // Executa sincronização em background
                Task.Run(() =>
                {
                    void Progress(string label)
                    {
                        try
                        {
                            var percent = Math.Min(100, Math.Max(0, (int)Math.Round((step * 100.0) / total)));
                            dlg.BeginInvoke(new Action(() => dlg.SetProgress(percent, label)));
                        }
                        catch { }
                    }

                    try
                    {
                        // Verifica host
                        Progress("Verificando host...");
                        if (!GroupClient.Ping(out var msg, connectTimeoutMs: 1500, ioTimeoutMs: 2000))
                            throw new Exception(string.IsNullOrWhiteSpace(msg) ? "Host indisponível" : msg);

                        // Baixa e importa
                        string syncLogPath2 = Path.Combine(AppContext.BaseDirectory, "sync_download.log");
                        void LogSync2(string dataset, int count)
                        {
                            try { File.AppendAllText(syncLogPath2, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {dataset}: {count} registros baixados\n"); } catch { }
                        }
                        step = 1; Progress("Sincronizando 20% – Baixando chaves...");
                        var chaves = GroupClient.PullChaves();
                        Database.ReplaceAllChavesLocal(chaves);
                        LogSync2("Chaves", chaves.Count);

                        step = 2; Progress("Sincronizando 40% – Baixando reservas...");
                        var reservas = GroupClient.PullReservas();
                        Database.ReplaceAllReservas(reservas);
                        LogSync2("Reservas", reservas.Count);

                        step = 3; Progress("Sincronizando 60% – Baixando relatórios...");
                        var rels = GroupClient.PullRelatorio();
                        Database.ReplaceAllRelatorios(rels);
                        LogSync2("Relatórios", rels.Count);

                        step = 4; Progress("Sincronizando 80% – Baixando alunos...");
                        var alunos = GroupClient.PullAlunos();
                        Database.ReplaceAllAlunos(alunos);
                        LogSync2("Alunos", alunos.Rows.Count);

                        step = 5; Progress("Sincronizando 100% – Baixando professores...");
                        var profs = GroupClient.PullProfessores();
                        Database.ReplaceAllProfessores(profs);
                        LogSync2("Professores", profs.Rows.Count);

                        // Conclui
                        dlg.BeginInvoke(new Action(() => dlg.Close()));

                        // Ativa modo cliente e salva configs
                        BeginInvoke(new Action(() =>
                        {
                            // Campo de apelido removido; não alterar node_alias aqui
                            GroupConfig.Mode = GroupMode.Client;
                            GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                            // Header
                            _lblMode.Text = $"Modo: {ModeDisplay(GroupConfig.Mode)}";
                            _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                            var host = GroupConfig.HostAddress;
                            if (string.IsNullOrWhiteSpace(host)) host = $"porta {GroupConfig.HostPort}";
                            _lblHost.Text = $"Host/Porta: {host}";
                            _lblSelf.Text = $"Este nó: {MedControl.SyncService.LocalNodeName()}";
                            UpdateStatusTimerMode();
                            _wasOnline = true;
                            try { MedControl.SyncService.ForceBeacon(); } catch { }
                            try { MedControl.SyncService.TryAddOrUpdateSelfPeer(); } catch { }
                            if (!auto)
                                MessageBox.Show(this, "Conectado e sincronizado com sucesso.", "Cliente", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }));
                    }
                    catch (Exception ex)
                    {
                        try { dlg.BeginInvoke(new Action(() => dlg.Close())); } catch { }
                        BeginInvoke(new Action(() =>
                        {
                            if (!auto)
                                MessageBox.Show(this, "Falha ao conectar/sincronizar: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            UpdateStatusTimerMode();
                        }));
                    }
                });

                // Loop de mensagens até fechar
                while (dlg.Visible)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao conectar: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DoCreateGroup()
        {
            try
            {
                // Define grupo e senha, torna-se Host e anuncia
                GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                GroupConfig.GroupPassword = _groupPassword.Text ?? string.Empty;
                GroupConfig.Mode = GroupMode.Host;
                try { GroupHost.Start(); } catch { }
                try { SyncService.ForceBeacon(); } catch { }
                // Atualiza cabeçalho
                _lblMode.Text = $"Modo: {ModeDisplay(GroupConfig.Mode)}";
                _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                var host = $"localhost:{GroupConfig.HostPort}";
                _lblHost.Text = $"Host/Porta: {host}";
                _lblSelf.Text = $"Este nó: {MedControl.SyncService.LocalNodeName()}";
                UpdateStatusTimerMode();
                MessageBox.Show(this, "Grupo criado neste dispositivo. Outros clientes podem entrar pelo nome e senha do grupo.", "Host ativo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Falha ao criar grupo: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                ApplyHostPortFromFields();
                var ok = GroupClient.Ping(out var msg, connectTimeoutMs: 1200, ioTimeoutMs: 1500);
                    // Tenta descobrir automaticamente o nome do grupo do Host informado se o Host estiver preenchido
                    if (!string.IsNullOrWhiteSpace(GroupConfig.HostAddress))
                    {
                        if (GroupClient.Hello(out var gname, out var hport, out var herr))
                        {
                            GroupConfig.GroupName = gname ?? "default";
                            GroupConfig.HostPort = hport;
                            _groupName.Text = GroupConfig.GroupName;
                            _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                        }
                        else if (!string.IsNullOrWhiteSpace(herr))
                        {
                            // Não bloqueia o teste, apenas informa em detalhe
                        }
                    }
                UpdateStatusBadge(ok, ok ? "" : msg ?? "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao testar conexão: " + ex.Message);
            }
        }

        // Removido método DoSave separado; usamos a função local em OnLoaded

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
                    int rtt = -1;
                    try { ok = GroupClient.Ping(out msg, out rtt); }
                    catch (Exception ex) { ok = false; msg = ex.Message; }
                    try
                    {
                        if (IsHandleCreated)
                            BeginInvoke(new Action(() =>
                            {
                                var detail = ok ? (rtt >= 0 ? $"{rtt} ms" : "") : (msg ?? "");
                                UpdateStatusBadge(ok, detail);
                            }));
                    }
                    catch { }

                    try
                    {
                        if (GroupConfig.Mode == GroupMode.Client)
                        {
                            if (ok)
                            {
                                // Se voltamos a ficar online e ainda não estávamos marcados como online, tenta reconectar/sincronizar automaticamente
                                if (!_wasOnline && !_autoReconnectInProgress)
                                {
                                    _autoReconnectInProgress = true;
                                    if (IsHandleCreated)
                                    {
                                        try
                                        {
                                            BeginInvoke(new Action(() =>
                                            {
                                                try { DoConnect(true); }
                                                finally { _autoReconnectInProgress = false; }
                                            }));
                                        }
                                        catch { _autoReconnectInProgress = false; }
                                    }
                                    else
                                    {
                                        _autoReconnectInProgress = false;
                                    }
                                }
                                _wasOnline = true;
                            }
                            else
                            {
                                _wasOnline = false;
                            }
                        }
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

        private void AppendChatLine(string sender, string message, DateTime localTime, string address, bool persist = true)
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

                // Evita crescimento infinito: mantém até ~500 linhas
                try
                {
                    var lines = _chatLog.Lines;
                    if (lines != null && lines.Length > 500)
                    {
                        var keep = lines.Skip(lines.Length - 500).ToArray();
                        _chatLog.Lines = keep;
                        _chatLog.SelectionStart = _chatLog.TextLength;
                        _chatLog.ScrollToCaret();
                    }
                    if (persist) SaveChatLine(time, sender, address, message);
                }
                catch { }
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

        private void SaveChatLine(string time, string sender, string address, string message)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_chatHistoryPath)) return;
                var line = $"[{time}] {sender} ({address}): {message}";
                System.IO.File.AppendAllText(_chatHistoryPath, line + Environment.NewLine);
            }
            catch { }
        }

        private void LoadChatHistory()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(_chatHistoryPath)) return;
                if (!System.IO.File.Exists(_chatHistoryPath)) return;
                var lines = System.IO.File.ReadAllLines(_chatHistoryPath);
                foreach (var raw in lines)
                {
                    // Não tenta re-colorir historico: append direto
                    _chatLog.AppendText(raw + Environment.NewLine);
                }
                _chatLog.SelectionStart = _chatLog.TextLength;
                _chatLog.ScrollToCaret();
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

        private void ApplyHostPortFromFields()
        {
            try
            {
                var hostField = _groupHost.Text?.Trim() ?? string.Empty;
                var portField = _groupPort.Text?.Trim() ?? string.Empty;
                // Remove qualquer prefixo ':' ou espaços extras no campo de porta
                while (portField.StartsWith(":")) portField = portField.Substring(1).Trim();
                // Se usuário colocou host:porta no campo de host, separar corretamente
                // Se o usuário digitou host:porta diretamente no campo host, separar
                var idx = hostField.IndexOf(':');
                if (idx > 0 && idx < hostField.Length - 1)
                {
                    var hostPart = hostField.Substring(0, idx).Trim();
                    var portPart = hostField.Substring(idx + 1).Trim();
                    while (portPart.StartsWith(":")) portPart = portPart.Substring(1).Trim();
                    if (int.TryParse(portPart, out var pColon) && pColon > 0)
                    {
                        hostField = hostPart;
                        if (string.IsNullOrWhiteSpace(portField)) portField = pColon.ToString();
                    }
                }
                // Sanitiza porta caso venha com host again (ex: ":49383" ou "192.168.0.5:49383")
                if (portField.Contains(':'))
                {
                    var parts = portField.Split(':');
                    var last = parts.Last();
                    if (int.TryParse(last, out var pSplit) && pSplit > 0) portField = pSplit.ToString();
                }
                // Caso host contenha novamente ":porta" remover
                var idx2 = hostField.IndexOf(':');
                if (idx2 > 0)
                {
                    var maybePort = hostField.Substring(idx2 + 1).Trim();
                    if (int.TryParse(maybePort, out var pAgain) && pAgain > 0)
                    {
                        hostField = hostField.Substring(0, idx2).Trim();
                        if (string.IsNullOrWhiteSpace(portField)) portField = pAgain.ToString();
                    }
                }
                // Porta padrão se não informada ou inválida
                if (int.TryParse(portField, out var p) && p > 0) GroupConfig.HostPort = p; else GroupConfig.HostPort = GroupConfig.HostPort > 0 ? GroupConfig.HostPort : 49383;
                GroupConfig.HostAddress = hostField; // sempre host puro
                _groupHost.Text = hostField;
                _groupPort.Text = GroupConfig.HostPort.ToString();
            }
            catch { }
        }
    }
}
