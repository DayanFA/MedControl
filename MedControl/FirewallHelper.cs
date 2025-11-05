using System;
using System.Diagnostics;

namespace MedControl
{
    internal static class FirewallHelper
    {
        // Tenta abrir exceções no Firewall do Windows para a porta TCP do Host e a porta UDP de beacons
        public static void TryOpenFirewall(int tcpPort)
        {
            try
            {
                // Porta UDP usada pelos beacons (igual ao SyncService; default 49382)
                int udpPort = 49382;
                try
                {
                    var s = Database.GetConfig("sync_udp_port");
                    if (int.TryParse(s, out var p) && p > 0) udpPort = p;
                }
                catch { }

                // Regras Inbound para perfis (Domain/Private/Public)
                RunNetsh($"advfirewall firewall add rule name=\"MedControl Host TCP {tcpPort}\" dir=in action=allow protocol=TCP localport={tcpPort} profile=any");
                RunNetsh($"advfirewall firewall add rule name=\"MedControl Beacons UDP {udpPort}\" dir=in action=allow protocol=UDP localport={udpPort} profile=any");
            }
            catch { }
        }

        private static void RunNetsh(string args)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "netsh",
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                };
                using var p = Process.Start(psi);
                if (p != null)
                {
                    try { p.WaitForExit(3000); } catch { }
                }
            }
            catch { }
        }
    }
}
