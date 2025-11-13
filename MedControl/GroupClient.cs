using System;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Diagnostics;
using System.Text.Json;

namespace MedControl
{
    public static class GroupClient
    {
        // Garante que a UI não trave tentando host indisponível repetidamente
        private static DateTime _lastFailUtc = DateTime.MinValue;
        private static readonly TimeSpan RetryAfter = TimeSpan.FromSeconds(10);

        public static bool ShouldTryRemote()
        {
            try { return (DateTime.UtcNow - _lastFailUtc) > RetryAfter; } catch { return true; }
        }

        private static void MarkFail()
        {
            try { _lastFailUtc = DateTime.UtcNow; } catch { }
        }

        private static void MarkSuccess()
        {
            try { _lastFailUtc = DateTime.MinValue; } catch { }
        }

        private static TcpClient ConnectWithTimeout(string hostOnly, int port, int timeoutMs)
        {
            var c = new TcpClient();
            var task = c.ConnectAsync(hostOnly, port);
            if (!task.Wait(timeoutMs))
            {
                try { c.Close(); } catch { }
                throw new TimeoutException("Timeout ao conectar ao Host");
            }
            return c;
        }

        private static (bool ok, string? error, string? data) Send(object req, int connectTimeoutMs = 200, int ioTimeoutMs = 400, int attempts = 1)
        {
            try
            {
                // Resolve destino: prefere host descoberto pelo grupo; cai para HostAddress
                string hostOnly = string.Empty;
                int port = GroupConfig.HostPort;
                if (GroupConfig.Mode == GroupMode.Client && SyncService.TryGetBestHost(out var addr, out var prt))
                {
                    hostOnly = addr;
                    port = prt;
                }
                else
                {
                    var host = GroupConfig.HostAddress;
                    var parts = host.Split(':');
                    hostOnly = parts[0];
                    if (parts.Length > 1 && int.TryParse(parts[1], out var p)) port = p;
                }

                Exception? lastEx = null;
                for (int i = 0; i < Math.Max(1, attempts); i++)
                {
                    var ct = i == 0 ? connectTimeoutMs : (i == 1 ? Math.Max(600, connectTimeoutMs * 3) : 3000);
                    var io = i == 0 ? ioTimeoutMs : Math.Max(ioTimeoutMs, (int)(ct * 1.5));
                    try
                    {
                        using var c = ConnectWithTimeout(hostOnly, port, ct);
                        c.NoDelay = true;
                        c.ReceiveTimeout = io; c.SendTimeout = io;
                        using var ns = c.GetStream();
                        using var writer = new StreamWriter(ns, new UTF8Encoding(false)) { AutoFlush = true };
                        using var reader = new StreamReader(ns, Encoding.UTF8, false);
                        // Empacota grupo/senha em todas as requisições
                        var json = MergeWithAuth(req);
                        writer.WriteLine(json);
                        var line = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(line)) { lastEx = new IOException("empty response"); throw lastEx; }
                        using var doc = JsonDocument.Parse(line);
                        var root = doc.RootElement;
                        var ok = root.GetProperty("ok").GetBoolean();
                        string? err = root.TryGetProperty("error", out var e) ? e.GetString() : null;
                        string? data = root.TryGetProperty("data", out var d) ? d.GetRawText() : null;
                        if (ok) MarkSuccess(); else MarkFail();
                        return (ok, err, data);
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        // se ainda há tentativas, continua; caso contrário, propaga
                        if (i < attempts - 1)
                        {
                            try { System.Threading.Thread.Sleep(100); } catch { }
                            continue;
                        }
                        break;
                    }
                }
                MarkFail();
                throw lastEx ?? new IOException("Falha de comunicação");
            }
            catch
            {
                MarkFail();
                throw;
            }
        }

        public static bool Ping(out string? message, int connectTimeoutMs = 150, int ioTimeoutMs = 400)
        {
            try { var ok = Ping(out message, out var _rtt, connectTimeoutMs, ioTimeoutMs); return ok; } catch (Exception ex) { message = ex.Message; MarkFail(); return false; }
        }

        public static bool Ping(out string? message, out int rttMs, int connectTimeoutMs = 150, int ioTimeoutMs = 400)
        {
            var sw = Stopwatch.StartNew();
            try
            {
                var r = Send(new { type = "ping" }, connectTimeoutMs: connectTimeoutMs, ioTimeoutMs: ioTimeoutMs, attempts: 1);
                sw.Stop();
                rttMs = (int)Math.Max(0, sw.Elapsed.TotalMilliseconds);
                if (r.ok) MarkSuccess(); else MarkFail();
                message = r.data;
                return r.ok;
            }
            catch (Exception ex)
            {
                sw.Stop();
                rttMs = -1;
                message = ex.Message;
                MarkFail();
                return false;
            }
        }

        // Obtém informações básicas do Host sem exigir autenticação (nome do grupo e porta)
        public static bool Hello(out string? groupName, out int port, out string? error, int connectTimeoutMs = 300, int ioTimeoutMs = 500)
        {
            groupName = null; port = GroupConfig.HostPort; error = null;
            try
            {
                var r = Send(new { type = "hello" }, connectTimeoutMs: connectTimeoutMs, ioTimeoutMs: ioTimeoutMs, attempts: 1);
                if (!r.ok)
                {
                    error = r.error ?? "hello falhou";
                    return false;
                }
                if (string.IsNullOrWhiteSpace(r.data))
                {
                    error = "resposta vazia";
                    return false;
                }
                using var doc = JsonDocument.Parse(r.data);
                var root = doc.RootElement;
                groupName = root.TryGetProperty("group", out var g) ? (g.GetString() ?? "default") : "default";
                port = root.TryGetProperty("port", out var p) && p.TryGetInt32(out var pp) ? pp : GroupConfig.HostPort;
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
        // Encaminhadores de escrita
        public static void InsertReserva(Reserva r) => Send(new { type = "write", entity = "reserva", data = r });
        public static void UpdateReserva(Reserva r) => Send(new { type = "write", entity = "reserva_update", data = r });
        public static void DeleteReserva(Reserva r) => Send(new { type = "write", entity = "reserva_delete", data = r });
        public static void InsertRelatorio(Relatorio r) => Send(new { type = "write", entity = "relatorio", data = r });
        public static void ReplaceAllChaves(System.Collections.Generic.List<Chave> list) => Send(new { type = "write", entity = "chaves_replace", data = list });
        public static void ReplaceAllAlunos(System.Data.DataTable table) => Send(new { type = "write", entity = "alunos_replace", data = JsonTableHelper.SerializeDataTable(table) });
        public static void ReplaceAllProfessores(System.Data.DataTable table) => Send(new { type = "write", entity = "professores_replace", data = JsonTableHelper.SerializeDataTable(table) });

        // Pulls
        public static System.Collections.Generic.List<Chave> PullChaves(int connectTimeoutMs = 200, int ioTimeoutMs = 600)
        {
            var r = Send(new { type = "pull", entity = "chaves" }, connectTimeoutMs, ioTimeoutMs, attempts: 3);
            return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Chave>>(r.data) ?? new() : new();
        }

        public static System.Collections.Generic.List<Relatorio> PullRelatorio(int connectTimeoutMs = 200, int ioTimeoutMs = 600)
        {
            var r = Send(new { type = "pull", entity = "relatorio" }, connectTimeoutMs, ioTimeoutMs, attempts: 3);
            return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Relatorio>>(r.data) ?? new() : new();
        }

        public static System.Collections.Generic.List<Reserva> PullReservas(int connectTimeoutMs = 200, int ioTimeoutMs = 600)
        {
            var r = Send(new { type = "pull", entity = "reservas" }, connectTimeoutMs, ioTimeoutMs, attempts: 3);
            return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Reserva>>(r.data) ?? new() : new();
        }

        // Faz pull diretamente de um endereço/porta específicos, ignorando heurísticas de host atual.
        public static System.Collections.Generic.List<Reserva> PullReservasFrom(string hostOnly, int port, int connectTimeoutMs = 400, int ioTimeoutMs = 800)
        {
            try
            {
                var req = new { type = "pull", entity = "reservas" };
                var r = SendTo(hostOnly, port, req, connectTimeoutMs, ioTimeoutMs, attempts: 2);
                return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Reserva>>(r.data) ?? new() : new();
            }
            catch { return new System.Collections.Generic.List<Reserva>(); }
        }

        // Versão interna de Send que permite especificar host/porta de destino diretamente
        private static (bool ok, string? error, string? data) SendTo(string hostOnly, int port, object req, int connectTimeoutMs, int ioTimeoutMs, int attempts = 1)
        {
            try
            {
                Exception? lastEx = null;
                for (int i = 0; i < Math.Max(1, attempts); i++)
                {
                    var ct = i == 0 ? connectTimeoutMs : Math.Max(600, connectTimeoutMs * 2);
                    var io = i == 0 ? ioTimeoutMs : Math.Max(ioTimeoutMs, (int)(ct * 1.5));
                    try
                    {
                        using var c = ConnectWithTimeout(hostOnly, port, ct);
                        c.NoDelay = true;
                        c.ReceiveTimeout = io; c.SendTimeout = io;
                        using var ns = c.GetStream();
                        using var writer = new StreamWriter(ns, new UTF8Encoding(false)) { AutoFlush = true };
                        using var reader = new StreamReader(ns, Encoding.UTF8, false);
                        var json = MergeWithAuth(req);
                        writer.WriteLine(json);
                        var line = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(line)) { lastEx = new IOException("empty response"); throw lastEx; }
                        using var doc = JsonDocument.Parse(line);
                        var root = doc.RootElement;
                        var ok = root.GetProperty("ok").GetBoolean();
                        string? err = root.TryGetProperty("error", out var e) ? e.GetString() : null;
                        string? data = root.TryGetProperty("data", out var d) ? d.GetRawText() : null;
                        if (ok) MarkSuccess(); else MarkFail();
                        return (ok, err, data);
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        if (i < attempts - 1)
                        {
                            try { System.Threading.Thread.Sleep(100); } catch { }
                            continue;
                        }
                        break;
                    }
                }
                MarkFail();
                throw lastEx ?? new IOException("Falha de comunicação");
            }
            catch
            {
                MarkFail();
                throw;
            }
        }

        public static System.Data.DataTable PullAlunos(int connectTimeoutMs = 200, int ioTimeoutMs = 800)
        {
            var r = Send(new { type = "pull", entity = "alunos" }, connectTimeoutMs, ioTimeoutMs, attempts: 3);
            return r.ok && r.data != null ? JsonTableHelper.DeserializeToDataTable(r.data) : new System.Data.DataTable();
        }

        public static System.Data.DataTable PullProfessores(int connectTimeoutMs = 200, int ioTimeoutMs = 800)
        {
            var r = Send(new { type = "pull", entity = "professores" }, connectTimeoutMs, ioTimeoutMs, attempts: 3);
            return r.ok && r.data != null ? JsonTableHelper.DeserializeToDataTable(r.data) : new System.Data.DataTable();
        }

        private static string MergeWithAuth(object req)
        {
            try
            {
                using var doc = JsonDocument.Parse(JsonSerializer.Serialize(req));
                var root = doc.RootElement;
                using var stream = new MemoryStream();
                using (var writer = new Utf8JsonWriter(stream))
                {
                    writer.WriteStartObject();
                    foreach (var prop in root.EnumerateObject()) prop.WriteTo(writer);
                    writer.WriteString("group", GroupConfig.GroupName ?? "default");
                    var pass = GroupConfig.GroupPassword ?? string.Empty;
                    if (!string.IsNullOrEmpty(pass)) writer.WriteString("auth", pass);
                    writer.WriteEndObject();
                }
                return Encoding.UTF8.GetString(stream.ToArray());
            }
            catch
            {
                // fallback simples
                var pass = GroupConfig.GroupPassword ?? string.Empty;
                var grp = GroupConfig.GroupName ?? "default";
                return JsonSerializer.Serialize(new { req, group = grp, auth = pass });
            }
        }
    }
}
