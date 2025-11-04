using System;
using System.IO;
using System.Net.Sockets;
using System.Text;
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

        private static (bool ok, string? error, string? data) Send(object req, int connectTimeoutMs = 200)
        {
            try
            {
                var host = GroupConfig.HostAddress;
                var port = GroupConfig.HostPort;
                var parts = host.Split(':');
                string hostOnly = parts[0];
                if (parts.Length > 1 && int.TryParse(parts[1], out var p)) port = p;

                using var c = ConnectWithTimeout(hostOnly, port, connectTimeoutMs);
                c.NoDelay = true;
                c.ReceiveTimeout = 300; c.SendTimeout = 300;
                using var ns = c.GetStream();
                using var writer = new StreamWriter(ns, new UTF8Encoding(false)) { AutoFlush = true };
                using var reader = new StreamReader(ns, Encoding.UTF8, false);
                var json = JsonSerializer.Serialize(req);
                writer.WriteLine(json);
                var line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) { MarkFail(); return (false, "empty response", null); }
                using var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                var ok = root.GetProperty("ok").GetBoolean();
                string? err = root.TryGetProperty("error", out var e) ? e.GetString() : null;
                string? data = root.TryGetProperty("data", out var d) ? d.GetRawText() : null;
                if (ok) MarkSuccess(); else MarkFail();
                return (ok, err, data);
            }
            catch
            {
                MarkFail();
                throw;
            }
        }

        public static bool Ping(out string? message)
        {
            try { var r = Send(new { type = "ping" }, connectTimeoutMs: 150); if (r.ok) MarkSuccess(); else MarkFail(); message = r.data; return r.ok; } catch (Exception ex) { message = ex.Message; MarkFail(); return false; }
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
        public static System.Collections.Generic.List<Chave> PullChaves()
        {
            var r = Send(new { type = "pull", entity = "chaves" });
            return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Chave>>(r.data) ?? new() : new();
        }

        public static System.Collections.Generic.List<Relatorio> PullRelatorio()
        {
            var r = Send(new { type = "pull", entity = "relatorio" });
            return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Relatorio>>(r.data) ?? new() : new();
        }

        public static System.Collections.Generic.List<Reserva> PullReservas()
        {
            var r = Send(new { type = "pull", entity = "reservas" });
            return r.ok && r.data != null ? JsonSerializer.Deserialize<System.Collections.Generic.List<Reserva>>(r.data) ?? new() : new();
        }

        public static System.Data.DataTable PullAlunos()
        {
            var r = Send(new { type = "pull", entity = "alunos" });
            return r.ok && r.data != null ? JsonTableHelper.DeserializeToDataTable(r.data) : new System.Data.DataTable();
        }

        public static System.Data.DataTable PullProfessores()
        {
            var r = Send(new { type = "pull", entity = "professores" });
            return r.ok && r.data != null ? JsonTableHelper.DeserializeToDataTable(r.data) : new System.Data.DataTable();
        }
    }
}
