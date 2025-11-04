using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Collections.Concurrent;

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
    private static readonly ConcurrentDictionary<string, DateTime> _seenChatIds = new();
    private static DateTime _lastPrune = DateTime.UtcNow;

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

                _beaconTimer = new System.Threading.Timer(_ => SendBeacon(), null, TimeSpan.FromSeconds(3), TimeSpan.FromSeconds(20));
                // envia um beacon inicial
                SendBeacon();
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
                    if (!string.Equals(grp, Group, StringComparison.Ordinal)) continue; // outro grupo

                    switch (type)
                    {
                        case "CHANGE":
                            // Ao receber, apenas sinalizamos para telas fazerem refresh (elas já fazem polling; isso torna quase instantâneo)
                            try { Database.SetConfig("last_change_at", DateTime.UtcNow.ToString("o")); } catch { }
                            break;
                        case "BEACON":
                            try
                            {
                                var node = parts.Length > 3 ? parts[3] : "?";
                                var hint = parts.Length > 4 ? parts[4] : string.Empty;
                                var role = parts.Length > 5 ? parts[5] : string.Empty;
                                int hostPort = 0;
                                if (parts.Length > 6) int.TryParse(parts[6], out hostPort);
                                var addr = ep.Address.ToString();
                                var key = node + "@" + addr;
                                var info = _peers.AddOrUpdate(key, _ => new PeerInfo
                                {
                                    Node = node,
                                    Address = addr,
                                    HostHint = hint,
                                    Role = role,
                                    HostPort = hostPort,
                                    LastSeen = DateTime.UtcNow
                                }, (_, existing) =>
                                {
                                    existing.Node = node;
                                    existing.Address = addr;
                                    existing.HostHint = hint;
                                    existing.Role = role;
                                    existing.HostPort = hostPort;
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
    }
}
