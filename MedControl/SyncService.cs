using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Collections.Concurrent;
using System.Net.NetworkInformation;

namespace MedControl
{
    // Serviço de descoberta e notificação LAN simples via UDP multicast/broadcast.
    // Objetivos:
    // - Descobrir "usuários próximos" (nós escutando no mesmo grupo/porta)
    // - Enviar beacons de presença ocasionais
    // - Notificar mudanças (para refresh quase imediato nas telas)
    // Não replica dados nem substitui o banco central; complementa a sincronização.
    public static class SyncService
    {
        private static UdpClient? _listener;
        private static IPEndPoint? _multicastEp;
        private static Thread? _listenThread;
    private static System.Threading.Timer? _beaconTimer;
        private static volatile bool _running;
    private static readonly ConcurrentDictionary<string, PeerInfo> _peers = new();
    private static readonly ConcurrentDictionary<string, DateTime> _groupsSeen = new();
    private static readonly ConcurrentDictionary<string, GroupAdvert> _groupAdverts = new();
    private static readonly ConcurrentDictionary<string, DateTime> _seenChatIds = new();
    private static DateTime _lastPrune = DateTime.UtcNow;
    private static DateTime _lastBgPullUtc = DateTime.MinValue;

        private static int Port => GetIntConfig("sync_udp_port", 49382);
        private static string Group => Database.GetConfig("sync_group") ?? "default";
        // Multicast reservado para redes privadas; pode cair para broadcast se necessário
        private static readonly IPAddress MulticastIp = IPAddress.Parse("239.0.0.222");

        public class PeerInfo
        {
            public string Node { get; set; } = "?";
            public string Address { get; set; } = "?";
            public string HostHint { get; set; } = string.Empty;
            public string Role { get; set; } = string.Empty; // "host" | "client" | "solo" | ""
            public int HostPort { get; set; } = 0; // porta TCP do host, se Role==host
            public string Group { get; set; } = "default"; // grupo do beacon
            public DateTime LastSeen { get; set; }
        }

        public class GroupAdvert
        {
            public string Group { get; set; } = "default";
            public string Node { get; set; } = "?";
            public string Address { get; set; } = "?";
            public int HostPort { get; set; } = 0;
            public string Role { get; set; } = string.Empty; // host|client|solo
            public DateTime LastSeen { get; set; }
        }

        // Nome local do nó (permite override por variável de ambiente ou config para testes)
        public static string LocalNodeName()
        {
            try
            {
                var env = Environment.GetEnvironmentVariable("MEDCONTROL_NODE");
                if (!string.IsNullOrWhiteSpace(env)) return env.Trim();
                var alias = Database.GetConfig("node_alias");
                if (!string.IsNullOrWhiteSpace(alias)) return alias.Trim();
            }
            catch { }
            return Environment.MachineName;
        }

        // Eventos para UI
        public static event Action? PeersChanged;
        public static event Action<string, string, DateTime, string>? OnChat; // sender, message, utc, address

        public static PeerInfo[] GetPeers()
        {
            try { return _peers.Values.OrderByDescending(p => p.LastSeen).ToArray(); } catch { return Array.Empty<PeerInfo>(); }
        }

        // Lista de grupos vistos recentemente via beacons (últimos 90s)
        public static string[] GetSeenGroups()
        {
            try
            {
                var now = DateTime.UtcNow;
                var cut = now.AddSeconds(-90);
                var list = _groupsSeen.Where(kv => kv.Value >= cut).Select(kv => kv.Key).Distinct().ToList();
                // Garante inclusão do grupo atual
                var cur = Group;
                if (!list.Contains(cur)) list.Add(cur);
                return list.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToArray();
            }
            catch { return new[] { Group }; }
        }

        // Retorna anúncios (recém vistos) por grupo, priorizando quem é host e com porta válida
        public static GroupAdvert[] GetGroupAdverts()
        {
            try
            {
                var now = DateTime.UtcNow;
                var cut = now.AddSeconds(-90);
                return _groupAdverts
                    .Select(kv => kv.Value)
                    .Where(a => a.LastSeen >= cut)
                    .OrderByDescending(a => string.Equals(a.Role, "host", StringComparison.OrdinalIgnoreCase))
                    .ThenByDescending(a => a.LastSeen)
                    .ToArray();
            }
            catch { return Array.Empty<GroupAdvert>(); }
        }

        // Melhor estimativa de host atual baseado nos beacons recentes
        public static bool TryGetBestHost(out string address, out int port)
        {
            try
            {
                var now = DateTime.UtcNow;
                var host = _peers.Values
                    .Where(p => string.Equals(p.Role, "host", StringComparison.OrdinalIgnoreCase) && p.HostPort > 0)
                    .OrderByDescending(p => p.LastSeen)
                    .FirstOrDefault();
                if (host != null && (now - host.LastSeen) < TimeSpan.FromSeconds(15))
                {
                    address = host.Address;
                    port = host.HostPort;
                    return true;
                }
            }
            catch { }
            address = string.Empty;
            port = GroupConfig.HostPort;
            return false;
        }

        public static void Start()
        {
            if (_running) return;
            _running = true;
            try
            {
                _multicastEp = new IPEndPoint(MulticastIp, Port);
                _listener = new UdpClient();
                _listener.ExclusiveAddressUse = false;
                _listener.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
                _listener.Client.Bind(new IPEndPoint(IPAddress.Any, Port));
                try { _listener.JoinMulticastGroup(MulticastIp); } catch { /* fallback se não suportado */ }

                _listenThread = new Thread(ListenLoop) { IsBackground = true, Name = "SyncService.Listen" };
                _listenThread.Start();

                _beaconTimer = new System.Threading.Timer(_ => SendBeacon(), null, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(5));
                // envia um beacon inicial
                SendBeacon();
                // garante que este nó apareça imediatamente em "Usuários próximos"
                TryAddOrUpdateSelfPeer();
            }
            catch { /* ignore */ }
        }

        public static void Stop()
        {
            _running = false;
            try { _beaconTimer?.Dispose(); } catch { }
            try { _listener?.DropMulticastGroup(MulticastIp); } catch { }
            try { _listener?.Dispose(); } catch { }
        }

        public static void NotifyChange()
        {
            try
            {
                var payload = $"MC|CHANGE|{Group}|{DateTime.UtcNow:o}";
                // Dispara em background para não bloquear UI
                System.Threading.Tasks.Task.Run(() =>
                {
                    try { Send(payload); } catch { }
                });
            }
            catch { }
        }

        public static void SendChat(string message)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(message)) return;
                var node = LocalNodeName();
                var id = Guid.NewGuid().ToString("N");
                var bytes = Encoding.UTF8.GetBytes(message);
                var b64 = Convert.ToBase64String(bytes);
                var payload = $"MC|CHAT|{Group}|{node}|{id}|{b64}";
                // Dispara em background para não bloquear UI
                System.Threading.Tasks.Task.Run(() =>
                {
                    try { Send(payload); } catch { }
                });
            }
            catch { }
        }

        private static void SendBeacon()
        {
            try
            {
                var hostHint = GetHostHint();
                var node = LocalNodeName();
                var role = GroupConfig.Mode == GroupMode.Host ? "host" : GroupConfig.Mode == GroupMode.Client ? "client" : "solo";
                var hostPort = GroupConfig.HostPort;
                var payload = $"MC|BEACON|{Group}|{node}|{hostHint}|{role}|{hostPort}";
                Send(payload);
            }
            catch { }
        }

        // Força o envio imediato de um beacon (usado após mudança de papel Host/Client)
        public static void ForceBeacon()
        {
            try { SendBeacon(); TryAddOrUpdateSelfPeer(); } catch { }
        }

        private static void Send(string payload)
        {
            try
            {
                var data = Encoding.UTF8.GetBytes(payload);
                using var c = new UdpClient();
                c.EnableBroadcast = true;
                // Não é necessário ingressar no grupo para enviar para multicast
                try { c.Send(data, data.Length, new IPEndPoint(MulticastIp, Port)); } catch { }
                // opcional: também por broadcast
                try { c.Send(data, data.Length, new IPEndPoint(IPAddress.Broadcast, Port)); } catch { }
            }
            catch { }
        }

        private static void ListenLoop()
        {
            var ep = new IPEndPoint(IPAddress.Any, 0);
            while (_running && _listener != null)
            {
                try
                {
                    var data = _listener.Receive(ref ep);
                    var text = Encoding.UTF8.GetString(data);
                    // formato: MC|TYPE|group|...
                    if (!text.StartsWith("MC|")) continue;
                    var parts = text.Split('|');
                    if (parts.Length < 3) continue;
                    var type = parts[1];
                    var grp = parts[2];
                    // Sempre registra o grupo visto, mesmo que não seja o atual
                    try { _groupsSeen[grp] = DateTime.UtcNow; } catch { }
                    if (!string.Equals(grp, Group, StringComparison.Ordinal))
                    {
                        // Ignora processamento detalhado (peers/chat) para outros grupos
                        continue;
                    }

                    switch (type)
                    {
                        case "CHANGE":
                            // Ao receber, sinaliza refresh e, se for Cliente, busca diffs do Host em background
                            try { Database.SetConfig("last_change_at", DateTime.UtcNow.ToString("o")); } catch { }
                            try
                            {
                                if (GroupConfig.Mode == GroupMode.Client && ShouldTryBackgroundPull())
                                {
                                    System.Threading.Tasks.Task.Run(() =>
                                    {
                                        try
                                        {
                                            var reservas = GroupClient.PullReservas(connectTimeoutMs: 600, ioTimeoutMs: 1200);
                                            Database.ReplaceAllReservasLocalSilent(reservas);
                                        }
                                        catch { }
                                    });
                                }
                            }
                            catch { }
                            break;
                        case "BEACON":
                            try
                            {
                                var node = parts.Length > 3 ? parts[3] : "?";
                                if (string.IsNullOrWhiteSpace(node)) node = "?";
                                var hint = parts.Length > 4 ? parts[4] : string.Empty;
                                var role = parts.Length > 5 ? parts[5] : string.Empty;
                                int hostPort = 0;
                                if (parts.Length > 6) int.TryParse(parts[6], out hostPort);
                                var addr = ep.Address.ToString();
                                // Registra anúncio por grupo para listagem no UI (mesmo quando group != atual acima)
                                try
                                {
                                    _groupAdverts.AddOrUpdate(grp, _ => new GroupAdvert
                                    {
                                        Group = grp,
                                        Node = node!,
                                        Address = addr,
                                        HostPort = hostPort,
                                        Role = role,
                                        LastSeen = DateTime.UtcNow
                                    }, (_, existing) =>
                                    {
                                        // preferir host; caso contrário, apenas atualiza se for mais recente
                                        bool prefer = string.Equals(role, "host", StringComparison.OrdinalIgnoreCase) && !string.Equals(existing.Role, "host", StringComparison.OrdinalIgnoreCase);
                                        if (prefer || existing.LastSeen < DateTime.UtcNow)
                                        {
                                            existing.Node = node!;
                                            existing.Address = addr;
                                            existing.HostPort = hostPort;
                                            existing.Role = role;
                                            existing.LastSeen = DateTime.UtcNow;
                                        }
                                        return existing;
                                    });
                                }
                                catch { }
                                var key = (node ?? "?").Trim().ToLowerInvariant();
                                _peers.AddOrUpdate(key, _ => new PeerInfo
                                {
                                    Node = node!,
                                    Address = addr,
                                    HostHint = hint,
                                    Role = role,
                                    HostPort = hostPort,
                                    Group = grp,
                                    LastSeen = DateTime.UtcNow
                                }, (_, existing) =>
                                {
                                    existing.Node = node!;
                                    // Evita duplicação por 127.0.0.1: prefere endereço não-loopback quando disponível
                                    if (PreferAddress(addr, existing.Address)) existing.Address = addr;
                                    existing.HostHint = hint;
                                    existing.Role = role;
                                    existing.HostPort = hostPort;
                                    existing.Group = grp;
                                    existing.LastSeen = DateTime.UtcNow;
                                    return existing;
                                });
                                PeersChanged?.Invoke();
                            }
                            catch { }
                            break;
                        case "CHAT":
                            try
                            {
                                var sender = parts.Length > 3 ? parts[3] : "?";
                                // formato novo: MC|CHAT|group|sender|id|b64
                                string msgId = string.Empty;
                                string b64;
                                if (parts.Length > 5)
                                {
                                    msgId = parts[4];
                                    b64 = parts[5];
                                }
                                else
                                {
                                    // compat: antigo não tinha id -> coloca hash do conteúdo como id
                                    b64 = parts.Length > 4 ? parts[4] : string.Empty;
                                    msgId = $"legacy:{sender}:{b64.GetHashCode()}";
                                }

                                // ignora eco das próprias mensagens (já mostramos localmente)
                                if (!string.IsNullOrWhiteSpace(sender) &&
                                    string.Equals(sender, LocalNodeName(), StringComparison.OrdinalIgnoreCase))
                                {
                                    break;
                                }

                                if (IsDuplicateChat(msgId)) break;
                                RememberChat(msgId);

                                string msg;
                                try { msg = Encoding.UTF8.GetString(Convert.FromBase64String(b64)); }
                                catch { msg = "(mensagem inválida)"; }
                                OnChat?.Invoke(sender, msg, DateTime.UtcNow, ep.Address.ToString());
                            }
                            catch { }
                            break;
                    }
                }
                catch { Thread.Sleep(200); }
            }
        }

        // Limita a frequência de pulls em resposta a eventos CHANGE para evitar tempestades
        private static bool ShouldTryBackgroundPull()
        {
            try
            {
                var now = DateTime.UtcNow;
                if ((now - _lastBgPullUtc).TotalSeconds >= 2)
                {
                    _lastBgPullUtc = now;
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static bool IsDuplicateChat(string id)
        {
            try
            {
                return _seenChatIds.ContainsKey(id);
            }
            catch { return false; }
        }

        private static void RememberChat(string id)
        {
            try
            {
                _seenChatIds[id] = DateTime.UtcNow;
                // prune occasionally
                var now = DateTime.UtcNow;
                if ((now - _lastPrune).TotalSeconds > 60 && _seenChatIds.Count > 128)
                {
                    _lastPrune = now;
                    foreach (var kv in _seenChatIds)
                    {
                        if ((now - kv.Value).TotalMinutes > 5)
                        {
                            _seenChatIds.TryRemove(kv.Key, out _);
                        }
                    }
                }
            }
            catch { }
        }

        private static int GetIntConfig(string key, int def)
        {
            try { return int.TryParse(Database.GetConfig(key), out var v) ? v : def; } catch { return def; }
        }

        // Divulga pista do host do banco quando usando MySQL
        private static string GetHostHint()
        {
            try
            {
                if ((AppConfig.Instance.DbProvider ?? "sqlite").Equals("mysql", StringComparison.OrdinalIgnoreCase))
                {
                    var cs = AppConfig.Instance.MySqlConnectionString ?? string.Empty;
                    var b = new MySqlConnector.MySqlConnectionStringBuilder(cs);
                    var server = string.IsNullOrWhiteSpace(b.Server) ? "?" : b.Server;
                    var db = string.IsNullOrWhiteSpace(b.Database) ? "?" : b.Database;
                    var port = b.Port;
                    return $"mysql://{server}:{port}/{db}"; // sem credenciais
                }
            }
            catch { }
            return "";
        }

        // Heurística de preferência de endereço: evita loopback (127.0.0.1), prefere IPv4 privada
        private static bool PreferAddress(string candidate, string current)
        {
            try
            {
                int Score(string a)
                {
                    if (!IPAddress.TryParse(a, out var ip)) return 0;
                    if (IPAddress.IsLoopback(ip)) return 1; // pior
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        // IPv4 privada: 10/8, 172.16/12, 192.168/16
                        var bytes = ip.GetAddressBytes();
                        bool isPrivate = bytes[0] == 10 || (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) || (bytes[0] == 192 && bytes[1] == 168);
                        return isPrivate ? 4 : 3;
                    }
                    if (ip.AddressFamily == AddressFamily.InterNetworkV6) return 2;
                    return 0;
                }
                return Score(candidate) >= Score(current);
            }
            catch { return false; }
        }

        // Tenta inserir/atualizar a presença do próprio nó na lista de peers (para aparecer de imediato no UI)
        public static void TryAddOrUpdateSelfPeer()
        {
            try
            {
                var node = LocalNodeName();
                var addr = GetLocalPreferredAddress();
                var role = GroupConfig.Mode == GroupMode.Host ? "host" : GroupConfig.Mode == GroupMode.Client ? "client" : "solo";
                var hostPort = GroupConfig.HostPort;
                var key = node.Trim().ToLowerInvariant();
                _peers.AddOrUpdate(key, _ => new PeerInfo
                {
                    Node = node,
                    Address = addr,
                    HostHint = GetHostHint(),
                    Role = role,
                    HostPort = hostPort,
                    LastSeen = DateTime.UtcNow
                }, (_, existing) =>
                {
                    existing.Node = node;
                    if (PreferAddress(addr, existing.Address)) existing.Address = addr;
                    existing.Role = role;
                    existing.HostPort = hostPort;
                    existing.HostHint = GetHostHint();
                    existing.LastSeen = DateTime.UtcNow;
                    return existing;
                });
                PeersChanged?.Invoke();
            }
            catch { }
        }

        private static string GetLocalPreferredAddress()
        {
            try
            {
                string? best = null;
                int bestScore = -1;
                foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (nic.OperationalStatus != OperationalStatus.Up) continue;
                    // Ignora loopback e interfaces de túnel desnecessárias
                    if (nic.NetworkInterfaceType == NetworkInterfaceType.Loopback) continue;
                    var ipProps = nic.GetIPProperties();
                    foreach (var unicast in ipProps.UnicastAddresses)
                    {
                        var ip = unicast.Address;
                        if (ip.AddressFamily != AddressFamily.InterNetwork) continue; // só IPv4
                        if (IPAddress.IsLoopback(ip)) continue;
                        var bytes = ip.GetAddressBytes();
                        bool is169 = bytes[0] == 169 && bytes[1] == 254; // link-local
                        if (is169) continue; // evita 169.254.x.x instável
                        bool isPrivate = bytes[0] == 10 || (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) || (bytes[0] == 192 && bytes[1] == 168);
                        // Score heurístico
                        int score = 0;
                        if (isPrivate) score += 5;
                        if (nic.NetworkInterfaceType == NetworkInterfaceType.Wireless80211) score += 2; // provável hotspot wifi
                        if (nic.Description.ToLowerInvariant().Contains("host") || nic.Description.ToLowerInvariant().Contains("wifi") || nic.Description.ToLowerInvariant().Contains("virtual")) score += 1;
                        // Se tem gateway, soma
                        if (ipProps.GatewayAddresses.Any(g => g.Address.AddressFamily == AddressFamily.InterNetwork)) score += 1;
                        if (score > bestScore)
                        {
                            bestScore = score;
                            best = ip.ToString();
                        }
                    }
                }
                if (!string.IsNullOrWhiteSpace(best)) return best!;
            }
            catch { }
            // Fallback antigo se heurística falhar
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork && !IPAddress.IsLoopback(ip))
                        return ip.ToString();
                }
            }
            catch { }
            return "127.0.0.1";
        }

        // Atualiza presença ao mudar endereços (ex: hotspot inicializa / cliente conecta)
        static SyncService()
        {
            try
            {
                NetworkChange.NetworkAddressChanged += (_, __) =>
                {
                    try { TryAddOrUpdateSelfPeer(); ForceBeacon(); } catch { }
                };
            }
            catch { }
        }
    }
}
