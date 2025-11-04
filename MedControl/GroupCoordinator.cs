using System;
using System.Linq;
using System.Threading;

namespace MedControl
{
    // Coordena o papel Host/Client automaticamente:
    // - Se existir um host visível via beacon recente, opera como Client e aponta para ele
    // - Se não existir host, e este nó for o primeiro na ordenação determinística, promove-se a Host
    // - Sempre mantém dados locais utilizáveis (Database já é offline-first)
    internal static class GroupCoordinator
    {
    private static System.Threading.Timer? _timer;
        private static readonly object _lock = new();
        private static DateTime _lastRoleChange = DateTime.MinValue;
        private static readonly TimeSpan RoleChangeCooldown = TimeSpan.FromSeconds(10);
        private static readonly TimeSpan HostStaleAfter = TimeSpan.FromSeconds(35);

        public static void Start()
        {
            try
            {
                _timer = new System.Threading.Timer(_ => Evaluate(), null, TimeSpan.FromSeconds(3), TimeSpan.FromSeconds(7));
            }
            catch { }
        }

        public static void Stop()
        {
            try { _timer?.Dispose(); } catch { }
        }

        private static void Evaluate()
        {
            if (!Monitor.TryEnter(_lock)) return;
            try
            {
                var now = DateTime.UtcNow;

                // 1) Procura por um host atual válido via beacons recentes
                var peers = SyncService.GetPeers();
                var hostPeer = peers
                    .Where(p => string.Equals(p.Role, "host", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(p => p.LastSeen)
                    .FirstOrDefault();

                bool hasFreshHost = hostPeer != null && (now - hostPeer.LastSeen) < HostStaleAfter && hostPeer.HostPort > 0;
                var meName = Environment.MachineName;

                if (hasFreshHost)
                {
                    // Se eu sou o host anunciado, garante modo Host ativo
                    if (string.Equals(hostPeer!.Node, meName, StringComparison.OrdinalIgnoreCase))
                    {
                        EnsureMode(GroupMode.Host, null);
                        return;
                    }
                    else
                    {
                        // Caso contrário, opera como Client apontando para o host
                        var addr = hostPeer.Address;
                        var hostAddr = $"{addr}:{hostPeer.HostPort}";
                        EnsureMode(GroupMode.Client, hostAddr);
                        return;
                    }
                }

                // 2) Não há host atual: eleição determinística pelo nome do nó
                // Inclui a si mesmo entre candidatos
                var candidateNames = peers.Select(p => p.Node)
                    .Concat(new[] { meName })
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (candidateNames.Length == 0)
                {
                    // Sem candidatos? segue solo
                    EnsureMode(GroupMode.Solo, null);
                    return;
                }
                var winner = candidateNames[0];
                if (string.Equals(winner, meName, StringComparison.OrdinalIgnoreCase))
                {
                    // Eu sou o eleito -> viro Host
                    EnsureMode(GroupMode.Host, null);
                }
                else
                {
                    // Outro nó é o eleito -> Cliente apontando para a melhor estimativa de endereço
                    // Tenta achar o peer desse vencedor para obter endereço; se não achar, mantém apenas porta local
                    var peer = peers.FirstOrDefault(p => string.Equals(p.Node, winner, StringComparison.OrdinalIgnoreCase));
                    var port = GroupConfig.HostPort;
                    if (peer != null)
                    {
                        var hostAddr = $"{peer.Address}:{(peer.HostPort > 0 ? peer.HostPort : port)}";
                        EnsureMode(GroupMode.Client, hostAddr);
                    }
                    else
                    {
                        // Sem endereço conhecido; ainda assim marca Client para evitar colisão de Hosts
                        EnsureMode(GroupMode.Client, GroupConfig.HostAddress);
                    }
                }
            }
            catch { }
            finally
            {
                try { Monitor.Exit(_lock); } catch { }
            }
        }

        private static void EnsureMode(GroupMode desired, string? hostAddress)
        {
            try
            {
                var now = DateTime.UtcNow;
                if ((now - _lastRoleChange) < RoleChangeCooldown)
                {
                    // evita oscillation
                    return;
                }

                var current = GroupConfig.Mode;
                bool changed = false;
                if (desired != current)
                {
                    GroupConfig.Mode = desired;
                    changed = true;
                }
                if (!string.IsNullOrWhiteSpace(hostAddress) && !string.Equals(GroupConfig.HostAddress, hostAddress, StringComparison.Ordinal))
                {
                    GroupConfig.HostAddress = hostAddress;
                    changed = true;
                }

                // Inicia/para servidor conforme necessário
                if (desired == GroupMode.Host)
                {
                    try { GroupHost.Start(); } catch { }
                }
                else
                {
                    try { GroupHost.Stop(); } catch { }
                }

                if (changed)
                {
                    _lastRoleChange = now;
                    // Notifica mudança para telas atualizarem
                    try { Database.SetConfig("last_change_reason", "role_change"); } catch { }
                    try { SyncService.NotifyChange(); } catch { }
                }
            }
            catch { }
        }
    }
}
