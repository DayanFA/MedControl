using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl
{
    internal static class UpdateService
    {
        private static readonly HttpClient _http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(8)
        };

        // Configure your repo here
        private const string GithubOwner = "DayanFA";
        private const string GithubRepo = "MedControl";
        private static DateTime _lastCheckUtc = DateTime.MinValue;

        public static void CheckSilentlyInBackground()
        {
            try
            {
                // Avoid checking too often (min 6 hours)
                if ((DateTime.UtcNow - _lastCheckUtc).TotalHours < 6) return;
                _lastCheckUtc = DateTime.UtcNow;
                Task.Run(async () =>
                {
                    try { await CheckAndPromptAsync(); } catch { }
                });
            }
            catch { }
        }

        private static Version GetCurrentVersion()
        {
            // Try ProductVersion first (may contain "+commit" metadata)
            try
            {
                var v = Application.ProductVersion;
                if (!string.IsNullOrWhiteSpace(v))
                {
                    if (TryParseVersionFlexible(v, out var ver)) return Normalize(ver);
                }
            }
            catch { }

            // Fallback: AssemblyInformationalVersion (may contain metadata)
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
                if (!string.IsNullOrWhiteSpace(info) && TryParseVersionFlexible(info!, out var ver)) return Normalize(ver);
            }
            catch { }

            // Fallback: AssemblyFileVersion (numeric)
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var fileVer = asm.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
                if (!string.IsNullOrWhiteSpace(fileVer) && Version.TryParse(fileVer, out var ver)) return Normalize(ver);
            }
            catch { }

            return new Version(1, 0, 0, 0);
        }

        private static bool TryParseVersionFlexible(string input, out Version ver)
        {
            // Strip any "+build" or "-prerelease" metadata and keep only the first numeric x[.y][.z][.w]
            // e.g., "1.1.0+7f8fdb0" -> "1.1.0", "v1.2.3-beta" -> "1.2.3"
            ver = new Version(1, 0, 0, 0);
            if (string.IsNullOrWhiteSpace(input)) return false;

            // Remove leading v/V and split off metadata (+ or -)
            var cleaned = input.Trim().TrimStart('v', 'V');
            var metaIdx = cleaned.IndexOfAny(new[] { '+', '-' });
            if (metaIdx > 0) cleaned = cleaned.Substring(0, metaIdx);

            // Keep only digits and dots from the beginning
            var m = Regex.Match(cleaned, "^\\d+(?:\\.\\d+){0,3}");
            if (!m.Success) return false;

            if (Version.TryParse(m.Value, out var parsed))
            {
                ver = parsed;
                return true;
            }
            return false;
        }

        private static Version Normalize(Version v)
        {
            // Keep 3 components for comparison (major.minor.patch)
            var patch = v.Build >= 0 ? v.Build : 0;
            return new Version(v.Major, Math.Max(0, v.Minor), Math.Max(0, patch));
        }

        private class GithubRelease
        {
            public string? tag_name { get; set; }
            public GithubAsset[]? assets { get; set; }
        }
        private class GithubAsset
        {
            public string? name { get; set; }
            public string? browser_download_url { get; set; }
        }

        private static async Task<(Version? version, string? url)> GetLatestReleaseAsync()
        {
            try
            {
                var api = $"https://api.github.com/repos/{GithubOwner}/{GithubRepo}/releases/latest";
                var req = new HttpRequestMessage(HttpMethod.Get, api);
                req.Headers.UserAgent.ParseAdd("MedControl-Updater");
                using var resp = await _http.SendAsync(req);
                resp.EnsureSuccessStatusCode();
                await using var s = await resp.Content.ReadAsStreamAsync();
                var rel = await JsonSerializer.DeserializeAsync<GithubRelease>(s);
                if (rel == null || string.IsNullOrWhiteSpace(rel.tag_name)) return (null, null);
                var tag = rel.tag_name.Trim().TrimStart('v', 'V');
                if (!Version.TryParse(tag, out var ver)) return (null, null);
                ver = Normalize(ver);
                // pick first exe asset
                var asset = rel.assets != null ? Array.Find(rel.assets, a => (a.name ?? "").EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) : null;
                var url = asset?.browser_download_url;
                return (ver, url);
            }
            catch
            {
                return (null, null);
            }
        }

        private static async Task CheckAndPromptAsync()
        {
            var cur = GetCurrentVersion();
            var (latest, url) = await GetLatestReleaseAsync();
            if (latest == null || string.IsNullOrWhiteSpace(url)) return;
            if (latest <= cur) return; // up-to-date

            try
            {
                var msg = $"Uma nova versão do MedControl está disponível (atual: {cur}, nova: {latest}).\n\nDeseja baixar e instalar agora?";
                var dr = MessageBox.Show(msg, "Atualização disponível", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;

                await DownloadAndRunInstallerAsync(url);
            }
            catch { }
        }

        // Public method to check now and inform the user even when up-to-date
        public static async Task CheckNowAsync(IWin32Window? owner = null)
        {
            try
            {
                var cur = GetCurrentVersion();
                var (latest, url) = await GetLatestReleaseAsync();
                if (latest == null)
                {
                    MessageBox.Show(owner ?? new Form(), "Não foi possível verificar atualizações agora.", "Atualização", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (latest <= cur)
                {
                    MessageBox.Show(owner ?? new Form(), $"Você já está na versão mais recente ({cur}).", "Atualização", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (string.IsNullOrWhiteSpace(url))
                {
                    MessageBox.Show(owner ?? new Form(), $"Nova versão disponível ({latest}), mas não há instalador anexado à release.", "Atualização", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var dr = MessageBox.Show(owner ?? new Form(), $"Uma nova versão do MedControl está disponível (atual: {cur}, nova: {latest}).\n\nDeseja baixar e instalar agora?", "Atualização disponível", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;
                await DownloadAndRunInstallerAsync(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner ?? new Form(), "Erro ao verificar atualização: " + ex.Message, "Atualização", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static async Task DownloadAndRunInstallerAsync(string url)
        {
            string tempFile = Path.Combine(Path.GetTempPath(), $"MedControl-Setup-{Guid.NewGuid():N}.exe");
            Form? progressForm = null;
            Label? progressLabel = null;
            ProgressBar? progressBar = null;
            try
            {
                // Cria mini janela de progresso (não modal) na UI thread
                var owner = Application.OpenForms.Count > 0 ? Application.OpenForms[0] : null;
                owner?.Invoke(new Action(() =>
                {
                    progressForm = new Form
                    {
                        Text = "Atualizando MedControl",
                        Width = 420,
                        Height = 140,
                        StartPosition = FormStartPosition.CenterScreen,
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        MaximizeBox = false,
                        MinimizeBox = false,
                        ShowInTaskbar = false,
                        TopMost = true
                    };
                    progressLabel = new Label
                    {
                        Dock = DockStyle.Top,
                        Height = 40,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Segoe UI", 10, FontStyle.Regular),
                        Text = "Preparando download..."
                    };
                    progressBar = new ProgressBar
                    {
                        Dock = DockStyle.Top,
                        Height = 28,
                        Style = ProgressBarStyle.Marquee,
                        MarqueeAnimationSpeed = 30
                    };
                    var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(18) };
                    panel.Controls.Add(progressBar);
                    panel.Controls.Add(progressLabel);
                    progressForm.Controls.Add(panel);
                    progressForm.Show(owner);
                }));
            }
            catch { }
            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.UserAgent.ParseAdd("MedControl-Updater");
                using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead);
                resp.EnsureSuccessStatusCode();
                var total = resp.Content.Headers.ContentLength ?? -1L;

                // Ajusta barra se soubermos o tamanho
                if (progressBar != null && progressLabel != null)
                {
                    progressBar.Invoke(new Action(() =>
                    {
                        if (total > 0)
                        {
                            progressBar.Style = ProgressBarStyle.Continuous;
                            progressBar.Minimum = 0; progressBar.Maximum = 100; progressBar.Value = 0;
                        }
                        progressLabel.Text = total > 0 ? "Baixando atualização..." : "Baixando (tamanho desconhecido)...";
                    }));
                }

                using var src = await resp.Content.ReadAsStreamAsync();
                using var dst = File.Create(tempFile);
                var buffer = new byte[81920];
                long read = 0;
                int n;
                while ((n = await src.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    await dst.WriteAsync(buffer.AsMemory(0, n));
                    read += n;
                    if (progressBar != null && progressLabel != null && total > 0)
                    {
                        var pct = (int)Math.Round(read * 100d / total);
                        var mbRead = read / 1024d / 1024d;
                        var mbTot = total / 1024d / 1024d;
                        try
                        {
                            progressBar.Invoke(new Action(() =>
                            {
                                progressBar.Value = Math.Min(100, Math.Max(0, pct));
                                progressLabel.Text = $"Baixando atualização... {pct}% ({mbRead:F1} MB de {mbTot:F1} MB)";
                            }));
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                try { File.Delete(tempFile); } catch { }
                try
                {
                    progressForm?.Invoke(new Action(() =>
                    {
                        progressLabel!.Text = "Falha no download.";
                        progressBar!.Style = ProgressBarStyle.Blocks;
                        progressBar.Value = 0;
                        progressForm!.Close();
                    }));
                }
                catch { }
                MessageBox.Show("Falha ao baixar a atualização: " + ex.Message, "Atualização", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                if (progressForm != null && progressLabel != null)
                {
                    progressForm.Invoke(new Action(() =>
                    {
                        progressLabel.Text = "Iniciando instalador...";
                    }));
                }
                // Executa instalador Inno Setup em modo silencioso. Ele cuidará da atualização (mesmo AppId).
                var psi = new ProcessStartInfo
                {
                    FileName = tempFile,
                    Arguments = "/VERYSILENT /NORESTART",
                    UseShellExecute = true,
                };
                Process.Start(psi);
                try
                {
                    progressForm?.Invoke(new Action(() => progressForm!.Close()));
                }
                catch { }
                // Fecha o app para liberar arquivos
                try { Application.Exit(); } catch { }
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                try
                {
                    progressForm?.Invoke(new Action(() =>
                    {
                        progressLabel!.Text = "Falha ao iniciar instalador.";
                        progressBar!.Style = ProgressBarStyle.Blocks;
                    }));
                }
                catch { }
                MessageBox.Show("Falha ao iniciar o instalador: " + ex.Message, "Atualização", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
