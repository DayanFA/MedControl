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
            _groupPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            _testConn = new Button { Text = "Testar Conexão", AutoSize = true };
            _saveBtn = new Button { Text = "Salvar", AutoSize = true };
            _connectBtn = new Button { Text = "Conectar", AutoSize = true };
            _createGroupBtn = new Button { Text = "Criar grupo (Host)", AutoSize = true };

            cfg.Controls.Add(Mk("Modo de Grupo:"), 0, 0);
            cfg.Controls.Add(_groupMode, 1, 0);
            cfg.Controls.Add(new Label { Text = " " }, 2, 0);

            cfg.Controls.Add(Mk("Nome do Grupo:"), 0, 1);
            cfg.Controls.Add(_groupName, 1, 1);
            cfg.Controls.Add(new Label { Text = "(ex.: lab1)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 1);

            _cfgLblHost = Mk("Host (modo Host):");
            cfg.Controls.Add(_cfgLblHost, 0, 2);
            cfg.Controls.Add(_groupHost, 1, 2);
            cfg.Controls.Add(_testConn, 2, 2);

            _cfgLblPort = Mk("Porta do Host:");
            cfg.Controls.Add(_cfgLblPort, 0, 3);
            cfg.Controls.Add(_groupPort, 1, 3);
            cfg.Controls.Add(new Label { Text = "(padrão 49383)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 3);

            _cfgLblPwd = Mk("Senha do Grupo:");
            cfg.Controls.Add(_cfgLblPwd, 0, 4);
            cfg.Controls.Add(_groupPassword, 1, 4);
            cfg.Controls.Add(new Label { Text = "(opcional)", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(6,8,3,8) }, 2, 4);

            // Removido campo de apelido do dispositivo a pedido do usuário

            // Botão de criar grupo/ser Host agora
            cfg.Controls.Add(_createGroupBtn, 2, 5);

            cfg.Controls.Add(new Label { Text = " " }, 0, 6);
            cfg.Controls.Add(_connectBtn, 1, 6);
            cfg.Controls.Add(_saveBtn, 2, 6);

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
                _groupPassword.Text = GroupConfig.GroupPassword;
                // Campo de apelido removido; mantemos configuração existente sem expor no UI
            }
            catch { }

            _groupMode.SelectedIndexChanged += (_, __) => { ToggleGroupUi(); UpdateStatusTimerMode(); };
            _testConn.Click += (_, __) => DoTestConn();
            _saveBtn.Click += (_, __) => DoSave();
            _connectBtn.Click += (_, __) => DoConnect();
            _createGroupBtn.Click += (_, __) => DoCreateGroup();
            
            void DoSave()
            {
                try
                {
                    var m = _groupMode.SelectedIndex switch { 1 => GroupMode.Host, 2 => GroupMode.Client, _ => GroupMode.Solo };
                    // Sempre atualiza nome/porta/host/senha a partir dos campos
                    GroupConfig.GroupName = string.IsNullOrWhiteSpace(_groupName.Text) ? "default" : _groupName.Text.Trim();
                    if (int.TryParse(_groupPort.Text?.Trim(), out var port) && port > 0) GroupConfig.HostPort = port;
                    GroupConfig.HostAddress = _groupHost.Text?.Trim() ?? string.Empty;
                    GroupConfig.GroupPassword = _groupPassword.Text ?? string.Empty;
                    // Campo de apelido removido; não alterar node_alias aqui

                    GroupConfig.Mode = m; // Persistir modo imediatamente, inclusive Cliente
                    if (m == GroupMode.Client)
                    {
                        // No modo Cliente, ao salvar já executa a sincronização com loading (sem exigir IP)
                        DoConnect();
                        return;
                    }

                    // Para Host/Offline: aplica e salva normalmente
                    try { if (m == GroupMode.Host) GroupHost.Start(); else GroupHost.Stop(); } catch { }

                    _lblMode.Text = $"Modo: {ModeDisplay(GroupConfig.Mode)}";
                    _lblGroup.Text = $"Grupo: {GroupConfig.GroupName}";
                    var host = GroupConfig.Mode == GroupMode.Host ? $"localhost:{GroupConfig.HostPort}" : GroupConfig.HostAddress;
                    if (string.IsNullOrWhiteSpace(host)) host = $"porta {GroupConfig.HostPort}";
                    _lblHost.Text = $"Host/Porta: {host}";
                    _lblSelf.Text = $"Este nó: {MedControl.SyncService.LocalNodeName()}";

                    MessageBox.Show(this, "Conexões salvas.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                // Em modo Cliente é opcional informar IP/Porta do Host para forçar conexão manual
                _groupHost.Visible = isClient;
                _cfgLblHost.Visible = isClient;
                _groupPort.Visible = true; // Porta é sempre relevante (Host usa para escutar; Cliente para conectar)
                _cfgLblPort.Visible = true;
                try { if (isClient) _cfgLblHost.Text = "Host (opcional):"; } catch { }
                // Senha do grupo é relevante em Cliente e Host
                _groupPassword.Visible = true;
                _cfgLblPwd.Visible = true;
                _testConn.Enabled = isClient;
                _connectBtn.Enabled = isClient;
                _createGroupBtn.Enabled = !isClient; // criar grupo faz sentido em Host/Offline

                // No modo Cliente, capturamos o nome do grupo automaticamente; desabilita a edição
                _groupName.Enabled = !isClient;
                if (isClient && string.IsNullOrWhiteSpace(_groupName.Text))
                {
                    _groupName.Text = "(automático)";
                }
            }
            catch { }
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
                if (int.TryParse(_groupPort.Text?.Trim(), out var p) && p > 0) GroupConfig.HostPort = p;
                GroupConfig.HostAddress = _groupHost.Text?.Trim() ?? string.Empty;

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
                        step = 1; Progress("Sincronizando 20% – Baixando chaves...");
                        var chaves = GroupClient.PullChaves();
                        Database.ReplaceAllChavesLocal(chaves);

                        step = 2; Progress("Sincronizando 40% – Baixando reservas...");
                        var reservas = GroupClient.PullReservas();
                        Database.ReplaceAllReservas(reservas);

                        step = 3; Progress("Sincronizando 60% – Baixando relatórios...");
                        var rels = GroupClient.PullRelatorio();
                        Database.ReplaceAllRelatorios(rels);

                        step = 4; Progress("Sincronizando 80% – Baixando alunos...");
                        var alunos = GroupClient.PullAlunos();
                        Database.ReplaceAllAlunos(alunos);

                        step = 5; Progress("Sincronizando 100% – Baixando professores...");
                        var profs = GroupClient.PullProfessores();
                        Database.ReplaceAllProfessores(profs);

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
                if (int.TryParse(_groupPort.Text?.Trim(), out var p) && p > 0) GroupConfig.HostPort = p;
                GroupConfig.HostAddress = _groupHost.Text?.Trim() ?? string.Empty;
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
