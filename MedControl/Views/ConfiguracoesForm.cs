using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class ConfiguracoesForm : Form
    {
    private readonly TextBox _alunos = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _profs = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly TextBox _chaves = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        private readonly Button _browseAlunos = new Button { Text = "📂 Selecionar Alunos.xlsx" };
        private readonly Button _browseProfs = new Button { Text = "📂 Selecionar Professores.xlsx" };
    private readonly Button _browseChaves = new Button { Text = "📂 Selecionar Chaves.xlsx" };
        private readonly Button _salvar = new Button { Text = "Salvar" };
        // Dados: Local | Online
        private readonly ComboBox _dadosModo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _urlAlunos = new TextBox { Dock = DockStyle.Fill };
        private readonly TextBox _urlProfs = new TextBox { Dock = DockStyle.Fill };
    private readonly TextBox _urlChaves = new TextBox { Dock = DockStyle.Fill };
        // Visual
        private readonly ComboBox _tema = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
    // Progresso (animação durante download)
    private readonly ProgressBar _progress = new ProgressBar { Style = ProgressBarStyle.Continuous, Visible = false, Width = 220, Height = 12, Anchor = AnchorStyles.Left };
    private readonly Label _progressLabel = new Label { Text = "Baixando e atualizando…", AutoSize = true, Visible = false, Anchor = AnchorStyles.Left, Font = new Font("Segoe UI", 9F) };
    private readonly Label _progressPercent = new Label { Text = "0%", AutoSize = true, Visible = false, Anchor = AnchorStyles.Left, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

        public ConfiguracoesForm()
        {
            Text = "Configurações";
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Width = 680;
            Height = 560;
            DoubleBuffered = true;

            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(18, 18, 18, 8), AutoSize = true };
            // A primeira coluna precisa comportar "Planilha de Professores:" sem quebra
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));

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
                b.Font = new Font("Segoe UI", 10F, primary ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Point);
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

            Controls.Add(table);

            // Linha 0: Dados (Local/Online)
            table.Controls.Add(MkLabel("Dados:"), 0, 0);
            _dadosModo.Items.AddRange(new object[] { "Local", "Online" });
            _dadosModo.Font = new Font("Segoe UI", 10F);
            _dadosModo.SelectedIndex = 0;
            table.Controls.Add(_dadosModo, 1, 0);
            table.Controls.Add(new Label { Text = " " }, 2, 0);

            // Linha 1: Alunos (Local)
            _alunos.Font = new Font("Segoe UI", 10F);
            try { _alunos.PlaceholderText = "Selecione o arquivo de alunos (xlsx)…"; } catch { }
            table.Controls.Add(MkLabel("Planilha de Alunos:"), 0, 1);
            table.Controls.Add(_alunos, 1, 1);
            StyleButton(_browseAlunos);
            table.Controls.Add(_browseAlunos, 2, 1);

            // Linha 2: Professores (Local)
            _profs.Font = new Font("Segoe UI", 10F);
            try { _profs.PlaceholderText = "Selecione o arquivo de professores (xlsx)…"; } catch { }
            table.Controls.Add(MkLabel("Planilha de Professores:"), 0, 2);
            table.Controls.Add(_profs, 1, 2);
            StyleButton(_browseProfs);
            table.Controls.Add(_browseProfs, 2, 2);

            // Linha 3: Chaves (Local)
            _chaves.Font = new Font("Segoe UI", 10F);
            try { _chaves.PlaceholderText = "Selecione o arquivo de chaves (xlsx)…"; } catch { }
            table.Controls.Add(MkLabel("Planilha de Chaves:"), 0, 3);
            table.Controls.Add(_chaves, 1, 3);
            StyleButton(_browseChaves);
            table.Controls.Add(_browseChaves, 2, 3);

            // Linha 4: URL Alunos (Online)
            table.Controls.Add(MkLabel("URL de Alunos:"), 0, 4);
            table.Controls.Add(_urlAlunos, 1, 4);
            // Linha 5: URL Professores (Online)
            table.Controls.Add(MkLabel("URL de Professores:"), 0, 5);
            table.Controls.Add(_urlProfs, 1, 5);
            // Linha 6: URL Chaves (Online)
            table.Controls.Add(MkLabel("URL de Chaves:"), 0, 6);
            table.Controls.Add(_urlChaves, 1, 6);
            table.Controls.Add(new Label { Text = " " }, 2, 6);

            // Linha 7: Tema
            table.Controls.Add(MkLabel("Tema:"), 0, 7);
            _tema.Items.AddRange(new object[] { "Padrão", "Claro", "Escuro", "Clássico", "Mica", "Alto Contraste", "Terminal" });
            _tema.SelectedIndex = 0;
            table.Controls.Add(_tema, 1, 7);

            // Linha 8: Salvar
            StyleButton(_salvar, primary: true);
            table.Controls.Add(new Label { Text = " " }, 0, 8);
            table.Controls.Add(new Label { Text = " " }, 1, 8);
            table.Controls.Add(_salvar, 2, 8);

            // Linha 9: Progresso (só aparece no modo Online ao salvar)
            table.Controls.Add(_progressLabel, 0, 9);
            table.Controls.Add(_progress, 1, 9);
            table.Controls.Add(_progressPercent, 2, 9);

            // Eventos
            _browseAlunos.Click += (_, __) =>
            {
                var last = _alunos.Text;
                if (string.IsNullOrWhiteSpace(last)) last = Database.GetConfig("caminho_alunos_local") ?? Database.GetConfig("caminho_alunos") ?? string.Empty;
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                try
                {
                    if (!string.IsNullOrWhiteSpace(last))
                    {
                        var dir = Path.GetDirectoryName(last);
                        if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir)) ofd.InitialDirectory = dir;
                        ofd.FileName = Path.GetFileName(last);
                    }
                }
                catch { }
                if (ofd.ShowDialog(this) == DialogResult.OK) _alunos.Text = ofd.FileName;
            };

            _browseProfs.Click += (_, __) =>
            {
                var last = _profs.Text;
                if (string.IsNullOrWhiteSpace(last)) last = Database.GetConfig("caminho_professores_local") ?? Database.GetConfig("caminho_professores") ?? string.Empty;
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                try
                {
                    if (!string.IsNullOrWhiteSpace(last))
                    {
                        var dir = Path.GetDirectoryName(last);
                        if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir)) ofd.InitialDirectory = dir;
                        ofd.FileName = Path.GetFileName(last);
                    }
                }
                catch { }
                if (ofd.ShowDialog(this) == DialogResult.OK) _profs.Text = ofd.FileName;
            };

            _browseChaves.Click += (_, __) =>
            {
                var last = _chaves.Text;
                if (string.IsNullOrWhiteSpace(last)) last = Database.GetConfig("caminho_chaves_local") ?? Database.GetConfig("caminho_chaves") ?? string.Empty;
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                try
                {
                    if (!string.IsNullOrWhiteSpace(last))
                    {
                        var dir = Path.GetDirectoryName(last);
                        if (!string.IsNullOrWhiteSpace(dir) && Directory.Exists(dir)) ofd.InitialDirectory = dir;
                        ofd.FileName = Path.GetFileName(last);
                    }
                }
                catch { }
                if (ofd.ShowDialog(this) == DialogResult.OK) _chaves.Text = ofd.FileName;
            };

            _dadosModo.SelectedIndexChanged += (_, __) => ToggleModo();
            _tema.SelectedIndexChanged += (_, __) => { /* Temas pré-definidos; aplicação no salvar */ };

            _salvar.Click += async (_, __) =>
            {
                // trava a UI durante a operação
                SetBusy(true);
                var modoSel = _dadosModo.SelectedIndex == 1 ? "online" : "local";
                Database.SetConfig("dados_modo", modoSel);
                Database.SetConfig("url_alunos", _urlAlunos.Text?.Trim() ?? string.Empty);
                Database.SetConfig("url_professores", _urlProfs.Text?.Trim() ?? string.Empty);
                Database.SetConfig("url_chaves", _urlChaves.Text?.Trim() ?? string.Empty);

                string? alunosPath = null;
                string? profsPath = null;
                string? chavesPath = null;

                try
                {
                    if (modoSel == "online")
                    {
                        // mostra animação de progresso
                        _progressLabel.Text = "Baixando alunos…";
                        _progress.Value = 0;
                        _progress.Visible = _progressLabel.Visible = true;
                        _progressPercent.Text = "0%";
                        _progressPercent.Visible = true;
                        UseWaitCursor = true;
                        var targetDir = EnsureDataDir();
                        if (!string.IsNullOrWhiteSpace(_urlAlunos.Text))
                        {
                            var p = Path.Combine(targetDir, "alunos.xlsx");
                            await DownloadToAsync(
                                _urlAlunos.Text.Trim(),
                                p,
                                percent: new Progress<int>(v =>
                                {
                                    if (v >= 0)
                                    {
                                        _progress.Style = ProgressBarStyle.Continuous;
                                        _progress.Value = Math.Min(100, Math.Max(0, v));
                                        _progressPercent.Text = _progress.Value + "%";
                                    }
                                }),
                                setIndeterminate: ind =>
                                {
                                    _progress.Style = ind ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
                                    _progressPercent.Visible = !ind;
                                }
                            );
                            alunosPath = p;
                            // Não altera o caminho local exibido; mantém separado
                            Database.SetConfig("caminho_alunos_online", p);
                        }
                        if (!string.IsNullOrWhiteSpace(_urlProfs.Text))
                        {
                            _progressLabel.Text = "Baixando professores…";
                            _progress.Value = 0;
                            _progressPercent.Text = "0%";
                            var p = Path.Combine(targetDir, "professores.xlsx");
                            await DownloadToAsync(
                                _urlProfs.Text.Trim(),
                                p,
                                percent: new Progress<int>(v =>
                                {
                                    if (v >= 0)
                                    {
                                        _progress.Style = ProgressBarStyle.Continuous;
                                        _progress.Value = Math.Min(100, Math.Max(0, v));
                                        _progressPercent.Text = _progress.Value + "%";
                                    }
                                }),
                                setIndeterminate: ind =>
                                {
                                    _progress.Style = ind ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
                                    _progressPercent.Visible = !ind;
                                }
                            );
                            profsPath = p;
                            // Não altera o caminho local exibido; mantém separado
                            Database.SetConfig("caminho_professores_online", p);
                        }
                        if (!string.IsNullOrWhiteSpace(_urlChaves.Text))
                        {
                            _progressLabel.Text = "Baixando chaves…";
                            _progress.Value = 0;
                            _progressPercent.Text = "0%";
                            var p = Path.Combine(targetDir, "chaves.xlsx");
                            await DownloadToAsync(
                                _urlChaves.Text.Trim(),
                                p,
                                percent: new Progress<int>(v =>
                                {
                                    if (v >= 0)
                                    {
                                        _progress.Style = ProgressBarStyle.Continuous;
                                        _progress.Value = Math.Min(100, Math.Max(0, v));
                                        _progressPercent.Text = _progress.Value + "%";
                                    }
                                }),
                                setIndeterminate: ind =>
                                {
                                    _progress.Style = ind ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
                                    _progressPercent.Visible = !ind;
                                }
                            );
                            chavesPath = p;
                            Database.SetConfig("caminho_chaves_online", p);
                        }
                    }
                    else
                    {
                        // Usa textbox ou, se vazio, usa caminho salvo do modo Local
                        var savedAlunos = Database.GetConfig("caminho_alunos_local") ?? Database.GetConfig("caminho_alunos") ?? string.Empty;
                        var savedProfs = Database.GetConfig("caminho_professores_local") ?? Database.GetConfig("caminho_professores") ?? string.Empty;
                        var savedChaves = Database.GetConfig("caminho_chaves_local") ?? Database.GetConfig("caminho_chaves") ?? string.Empty;
                        alunosPath = !string.IsNullOrWhiteSpace(_alunos.Text) ? _alunos.Text : savedAlunos;
                        profsPath = !string.IsNullOrWhiteSpace(_profs.Text) ? _profs.Text : savedProfs;
                        chavesPath = !string.IsNullOrWhiteSpace(_chaves.Text) ? _chaves.Text : savedChaves;
                        if (!string.IsNullOrWhiteSpace(alunosPath)) _alunos.Text = alunosPath;
                        if (!string.IsNullOrWhiteSpace(profsPath)) _profs.Text = profsPath;
                        if (!string.IsNullOrWhiteSpace(chavesPath)) _chaves.Text = chavesPath;
                    }

                    // Persiste somente os caminhos do modo Local, não sobrescrevendo quando for Online
                    if (modoSel == "local")
                    {
                        if (!string.IsNullOrWhiteSpace(alunosPath)) { Database.SetConfig("caminho_alunos_local", alunosPath); Database.SetConfig("caminho_alunos", alunosPath); }
                        if (!string.IsNullOrWhiteSpace(profsPath))  { Database.SetConfig("caminho_professores_local", profsPath); Database.SetConfig("caminho_professores", profsPath); }
                        if (!string.IsNullOrWhiteSpace(chavesPath)) { Database.SetConfig("caminho_chaves_local", chavesPath); Database.SetConfig("caminho_chaves", chavesPath); }
                    }

                    // Importa automaticamente para os respectivos cadastros (DB)
                    try { Database.Setup(); } catch { }
                    if (!string.IsNullOrEmpty(alunosPath) && File.Exists(alunosPath))
                    {
                        var dtAlunos = ExcelHelper.LoadToDataTable(alunosPath);
                        EnsureInternalId(dtAlunos);
                        await Task.Run(() => Database.ReplaceAllAlunos(dtAlunos));
                    }
                    if (!string.IsNullOrEmpty(profsPath) && File.Exists(profsPath))
                    {
                        var dtProfs = ExcelHelper.LoadToDataTable(profsPath);
                        EnsureInternalId(dtProfs);
                        await Task.Run(() => Database.ReplaceAllProfessores(dtProfs));
                    }
                    if (!string.IsNullOrEmpty(chavesPath) && File.Exists(chavesPath))
                    {
                        var dtChaves = ExcelHelper.LoadToDataTable(chavesPath);
                        var list = ParseChaves(dtChaves);
                        await Task.Run(() => Database.ReplaceAllChaves(list));
                    }

                    // Se telas de cadastro estiverem abertas, tenta recarregar
                    try
                    {
                        foreach (Form f in Application.OpenForms)
                        {
                            var t = f.GetType();
                            if (t.Name == "CadastroAlunosForm" || t.Name == "CadastroProfessoresForm" || t.Name == "CadastroChavesForm")
                            {
                                var mi = t.GetMethod("ReloadFromDatabase", BindingFlags.Instance | BindingFlags.NonPublic);
                                mi?.Invoke(f, null);
                            }
                        }
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Falha ao atualizar dados: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    // esconde animação e destrava UI
                    _progress.Visible = _progressLabel.Visible = _progressPercent.Visible = false;
                    UseWaitCursor = false;
                    SetBusy(false);
                }

                // Salva e aplica tema
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
                try { MedControl.UI.ThemeHelper.ApplyVisualStyleStateForCurrentTheme(); MedControl.UI.ThemeHelper.ApplyToAllOpenForms(); } catch { }

                MessageBox.Show(this, "Configurações salvas e dados atualizados.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            };

            // Inicializa valores persistidos
            try
            {
                var modo = (Database.GetConfig("dados_modo") ?? "local").ToLowerInvariant();
                _dadosModo.SelectedIndex = modo == "online" ? 1 : 0;
                _urlAlunos.Text = Database.GetConfig("url_alunos") ?? string.Empty;
                _urlProfs.Text = Database.GetConfig("url_professores") ?? string.Empty;
                _urlChaves.Text = Database.GetConfig("url_chaves") ?? string.Empty;
                var savedAlunosLocal = Database.GetConfig("caminho_alunos_local") ?? Database.GetConfig("caminho_alunos") ?? string.Empty;
                var savedProfsLocal = Database.GetConfig("caminho_professores_local") ?? Database.GetConfig("caminho_professores") ?? string.Empty;
                var savedChavesLocal = Database.GetConfig("caminho_chaves_local") ?? Database.GetConfig("caminho_chaves") ?? string.Empty;
                _alunos.Text = savedAlunosLocal;
                _profs.Text = savedProfsLocal;
                _chaves.Text = savedChavesLocal;
                var theme = (Database.GetConfig("theme") ?? "padrao").ToLowerInvariant();
                _tema.SelectedIndex = theme switch { "padrao" => 0, "claro" => 1, "escuro" => 2, "classico" => 3, "mica" => 4, "alto_contraste" => 5, "terminal" => 6, _ => 0 };
            }
            catch { }

            ToggleModo();

            // Aplicar tema atual
            try { MedControl.UI.ThemeHelper.ApplyCurrentTheme(this); } catch { }
        }

        private static string EnsureDataDir()
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var dir = Path.Combine(appData, "MedControl", "Dados");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return dir;
            }
            catch
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "Dados");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return dir;
            }
        }

        private static async Task DownloadToAsync(string url, string targetPath, IProgress<int>? percent = null, Action<bool>? setIndeterminate = null)
        {
            // Normaliza links do Google Sheets/Drive para baixar XLSX diretamente
            var normalizedUrl = NormalizeSpreadsheetUrl(url);

            using var http = new HttpClient();
            http.DefaultRequestHeaders.UserAgent.ParseAdd("MedControl/1.0 (+https://github.com/)");
            using var resp = await http.GetAsync(normalizedUrl, HttpCompletionOption.ResponseHeadersRead);
            resp.EnsureSuccessStatusCode();

            var total = resp.Content.Headers.ContentLength;
            if (total == null || total <= 0)
            {
                setIndeterminate?.Invoke(true);
                await using var fs = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await resp.Content.CopyToAsync(fs);
                await fs.FlushAsync();
            }
            else
            {
                setIndeterminate?.Invoke(false);
                await using var stream = await resp.Content.ReadAsStreamAsync();
                await using var fs = new FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None);
                var buffer = new byte[81920];
                long readTotal = 0;
                int read;
                int lastReported = -1;
                while ((read = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    await fs.WriteAsync(buffer, 0, read);
                    readTotal += read;
                    var p = (int)(readTotal * 100 / total.Value);
                    if (p != lastReported)
                    {
                        percent?.Report(p);
                        lastReported = p;
                    }
                }
                await fs.FlushAsync();
            }

            // Valida assinatura ZIP (XLSX) para evitar arquivos HTML salvos como .xlsx
            EnsureLooksLikeXlsx(targetPath, normalizedUrl);
        }

        private static string NormalizeSpreadsheetUrl(string url)
        {
            try
            {
                var u = new Uri(url);
                var host = u.Host.ToLowerInvariant();

                if (host.Contains("docs.google.com") && u.AbsolutePath.Contains("/spreadsheets/d/"))
                {
                    var m = Regex.Match(u.AbsolutePath, @"/spreadsheets/d/([^/]+)");
                    if (m.Success)
                    {
                        var id = m.Groups[1].Value;
                        return $"https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx";
                    }
                }

                if (host.Contains("drive.google.com") && u.AbsolutePath.Contains("/file/d/"))
                {
                    var m = Regex.Match(u.AbsolutePath, @"/file/d/([^/]+)");
                    if (m.Success)
                    {
                        var id = m.Groups[1].Value;
                        return $"https://drive.google.com/uc?export=download&id={id}";
                    }
                }
            }
            catch { }
            return url;
        }

        private static void EnsureLooksLikeXlsx(string path, string sourceUrl)
        {
            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                var sig = new byte[2];
                var n = fs.Read(sig, 0, 2);
                if (n != 2 || sig[0] != (byte)'P' || sig[1] != (byte)'K')
                {
                    throw new InvalidDataException($"O arquivo baixado não parece um .xlsx válido. Se o link for do Google Sheets, use: https://docs.google.com/spreadsheets/d/<ID>/export?format=xlsx (seu link foi normalizado para: {sourceUrl}).");
                }
            }
            catch (Exception ex)
            {
                // Repropaga com mensagem amigável
                throw new InvalidDataException("Arquivo inválido para XLSX: " + ex.Message, ex);
            }
        }

        private void ToggleModo()
        {
            var online = _dadosModo.SelectedIndex == 1;
            // Local controls
            _alunos.Enabled = !online;
            _browseAlunos.Enabled = !online;
            _profs.Enabled = !online;
            _browseProfs.Enabled = !online;
            _chaves.Enabled = !online;
            _browseChaves.Enabled = !online;
            // Online controls
            _urlAlunos.Enabled = online;
            _urlProfs.Enabled = online;
            _urlChaves.Enabled = online;
            // Mostra/oculta linha de progresso apenas quando necessário (exibida ao salvar)
            _progressLabel.Visible = false;
            _progress.Visible = false;
            // Ao mudar para Local, preenche com últimos caminhos salvos se estiver vazio
            if (!online)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(_alunos.Text))
                        _alunos.Text = Database.GetConfig("caminho_alunos_local") ?? Database.GetConfig("caminho_alunos") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(_profs.Text))
                        _profs.Text = Database.GetConfig("caminho_professores_local") ?? Database.GetConfig("caminho_professores") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(_chaves.Text))
                        _chaves.Text = Database.GetConfig("caminho_chaves_local") ?? Database.GetConfig("caminho_chaves") ?? string.Empty;
                }
                catch { }
            }
        }

        private void SetBusy(bool busy)
        {
            try
            {
                _dadosModo.Enabled = !busy;
                _urlAlunos.Enabled = !busy && _dadosModo.SelectedIndex == 1;
                _urlProfs.Enabled = !busy && _dadosModo.SelectedIndex == 1;
                _urlChaves.Enabled = !busy && _dadosModo.SelectedIndex == 1;
                _alunos.Enabled = !busy && _dadosModo.SelectedIndex != 1;
                _profs.Enabled = !busy && _dadosModo.SelectedIndex != 1;
                _chaves.Enabled = !busy && _dadosModo.SelectedIndex != 1;
                _browseAlunos.Enabled = !busy && _dadosModo.SelectedIndex != 1;
                _browseProfs.Enabled = !busy && _dadosModo.SelectedIndex != 1;
                _browseChaves.Enabled = !busy && _dadosModo.SelectedIndex != 1;
                _salvar.Enabled = !busy;
            }
            catch { }
        }

        private static void EnsureInternalId(DataTable table)
        {
            if (table == null) return;
            if (!table.Columns.Contains("_id"))
            {
                table.Columns.Add("_id", typeof(string));
                foreach (DataRow r in table.Rows) r["_id"] = Guid.NewGuid().ToString();
            }
            else
            {
                var idCol = table.Columns["_id"];
                if (idCol != null && idCol.DataType != typeof(string))
                {
                    idCol.ColumnName = "_id_old";
                    table.Columns.Add("_id", typeof(string));
                    foreach (DataRow r in table.Rows)
                    {
                        var v = r.Table.Columns.Contains("_id_old") ? r["_id_old"] : null;
                        r["_id"] = v == null || v == DBNull.Value ? Guid.NewGuid().ToString() : v.ToString();
                    }
                    table.Columns.Remove("_id_old");
                }
            }
        }

        private static string NormalizeHeader(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.Trim().ToLowerInvariant();
            s = s
                .Replace("á", "a").Replace("à", "a").Replace("ã", "a").Replace("â", "a")
                .Replace("é", "e").Replace("ê", "e")
                .Replace("í", "i")
                .Replace("ó", "o").Replace("õ", "o").Replace("ô", "o")
                .Replace("ú", "u")
                .Replace("ç", "c");
            s = new string(s.Where(ch => char.IsLetterOrDigit(ch) || ch == ' ').ToArray());
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            return s;
        }

        private static System.Collections.Generic.List<Chave> ParseChaves(DataTable dt)
        {
            var list = new System.Collections.Generic.List<Chave>();
            if (dt == null || dt.Columns.Count == 0) return list;
            int? idxNome = null, idxNum = null, idxDesc = null;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var norm = NormalizeHeader(dt.Columns[i].ColumnName);
                if (idxNome == null && (norm == "nome")) idxNome = i;
                else if (idxNum == null && (norm == "numero de copias" || norm == "numerodecopias" || norm == "num de copias" || norm == "qtd de copias")) idxNum = i;
                else if (idxDesc == null && (norm == "descricao" || norm == "descrição" || norm == "descricao observacao")) idxDesc = i;
            }
            if (idxNome == null) return list;
            foreach (DataRow r in dt.Rows)
            {
                var nome = Convert.ToString(r[(int)idxNome])?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(nome)) continue;
                int num = 0; if (idxNum != null) int.TryParse(Convert.ToString(r[(int)idxNum])?.Trim() ?? "0", out num);
                var desc = idxDesc != null ? (Convert.ToString(r[(int)idxDesc]) ?? string.Empty) : string.Empty;
                list.Add(new Chave { Nome = nome, NumCopias = num, Descricao = desc });
            }
            return list;
        }
    }
}
