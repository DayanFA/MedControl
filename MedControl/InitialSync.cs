using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MedControl
{
    internal static class InitialSync
    {
        public static void RunWithUiSplash()
        {
            try
            {
                // Só exibe quando for cliente – hosts e solo já têm fonte local
                if (GroupConfig.Mode != GroupMode.Client) return;

                using var splash = new MedControl.Views.PreloadForm();
                splash.Show();
                splash.Refresh();

                // Roda a sincronização em background para não travar a UI
                Task.Run(async () =>
                {
                    try
                    {
                        await RunSyncStepsAsync(splash);
                    }
                    catch { }
                    try { splash.BeginInvoke(new Action(() => splash.Close())); } catch { }
                });

                // Processa eventos até fechar
                while (splash.Visible)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(50);
                }
            }
            catch { }
        }

        private static async Task RunSyncStepsAsync(MedControl.Views.PreloadForm splash)
        {
            // Pequena janela para o coordenador descobrir um host (se houver)
            await Task.Delay(500);

            if (!GroupClient.ShouldTryRemote()) return; // sem remoto – mantém local

            try
            {
                // Se há HostAddress configurado, tenta capturar automaticamente nome do grupo e porta do Host
                try
                {
                    if (!string.IsNullOrWhiteSpace(GroupConfig.HostAddress))
                    {
                        if (GroupClient.Hello(out var gname, out var hport, out var herr))
                        {
                            GroupConfig.GroupName = gname ?? "default";
                            GroupConfig.HostPort = hport;
                        }
                    }
                }
                catch { }

                splash.SetStatus("Verificando host...");
                string? _;
                if (!GroupClient.Ping(out _)) return; // sem host – segue local
            }
            catch { return; }

            try
            {
                splash.SetStatus("Baixando chaves...");
                var chaves = GroupClient.PullChaves();
                Database.ReplaceAllChavesLocal(chaves);
            }
            catch { }

            try
            {
                splash.SetStatus("Baixando reservas...");
                var reservas = GroupClient.PullReservas();
                Database.ReplaceAllReservas(reservas);
            }
            catch { }

            try
            {
                splash.SetStatus("Baixando relatórios...");
                var rels = GroupClient.PullRelatorio();
                Database.ReplaceAllRelatorios(rels);
            }
            catch { }

            try
            {
                splash.SetStatus("Baixando alunos...");
                var alunos = GroupClient.PullAlunos();
                Database.ReplaceAllAlunos(alunos);
            }
            catch { }

            try
            {
                splash.SetStatus("Baixando professores...");
                var profs = GroupClient.PullProfessores();
                Database.ReplaceAllProfessores(profs);
            }
            catch { }

            await Task.Delay(200);

            // Anuncia presença e atualiza cache de peers para que a UI já mostre online
            try { MedControl.SyncService.ForceBeacon(); } catch { }
            try { MedControl.SyncService.TryAddOrUpdateSelfPeer(); } catch { }
        }
    }
}
