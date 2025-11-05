using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace MedControl
{
    // Servidor TCP simples para o "Host do Grupo".
    // Protocolo: uma linha JSON por requisição. Exemplo:
    // {"type":"ping"}
    // {"type":"write","entity":"reserva","data":{...}}
    // {"type":"pull","entity":"chaves"}
    // Resposta: JSON com { ok, error?, data? }
    public static class GroupHost
    {
        private static TcpListener? _listener;
        private static Thread? _thread;
        private static volatile bool _running;

        public static void Start()
        {
            if (_running || GroupConfig.Mode != GroupMode.Host) return;
            _running = true;
            try
            {
                var port = GroupConfig.HostPort;
                _listener = new TcpListener(IPAddress.Any, port);
                _listener.Start();
                _thread = new Thread(ListenLoop) { IsBackground = true, Name = "GroupHost.Listen" };
                _thread.Start();
            }
            catch { _running = false; }
        }

        public static void Stop()
        {
            _running = false;
            try { _listener?.Stop(); } catch { }
        }

        private static void ListenLoop()
        {
            while (_running && _listener != null)
            {
                try
                {
                    var client = _listener.AcceptTcpClient();
                    var _ = new Thread(() => HandleClient(client)) { IsBackground = true };
                    _.Start();
                }
                catch { Thread.Sleep(100); }
            }
        }

        private static void HandleClient(TcpClient c)
        {
            using (c)
            using (var ns = c.GetStream())
            using (var reader = new StreamReader(ns, Encoding.UTF8, leaveOpen: true))
            using (var writer = new StreamWriter(ns, new UTF8Encoding(false)) { AutoFlush = true })
            {
                try
                {
                    var line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) return;
                    using var doc = JsonDocument.Parse(line);
                    var root = doc.RootElement;
                    var type = root.GetProperty("type").GetString() ?? "";

                    // Atende "hello" sem exigir autenticação, para permitir que clientes descubram o nome do grupo/porta do Host
                    if (string.Equals(type, "hello", StringComparison.OrdinalIgnoreCase))
                    {
                        var payload = new
                        {
                            ok = true,
                            data = new { group = GroupConfig.GroupName ?? "default", port = GroupConfig.HostPort }
                        };
                        writer.WriteLine(JsonSerializer.Serialize(payload));
                        return;
                    }

                    // Autorização simples por grupo/senha para os demais tipos
                    if (!ValidateAuth(root, out var authError))
                    {
                        writer.WriteLine(JsonSerializer.Serialize(new { ok = false, error = authError }));
                        return;
                    }
                    switch (type)
                    {
                        case "ping":
                            writer.WriteLine("{\"ok\":true,\"data\":\"pong\"}");
                            break;
                        case "write":
                            HandleWrite(root, writer);
                            break;
                        case "pull":
                            HandlePull(root, writer);
                            break;
                        default:
                            writer.WriteLine("{\"ok\":false,\"error\":\"unknown type\"}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    var msg = JsonSerializer.Serialize(new { ok = false, error = ex.Message });
                    writer.WriteLine(msg);
                }
            }
        }

        private static bool ValidateAuth(JsonElement root, out string error)
        {
            try
            {
                var reqGroup = root.TryGetProperty("group", out var g) ? (g.GetString() ?? string.Empty) : string.Empty;
                var reqAuth = root.TryGetProperty("auth", out var a) ? (a.GetString() ?? string.Empty) : string.Empty;
                var groupOk = string.Equals(reqGroup, GroupConfig.GroupName ?? "default", StringComparison.Ordinal);
                var pass = GroupConfig.GroupPassword ?? string.Empty;
                var authOk = string.IsNullOrEmpty(pass) || string.Equals(reqAuth, pass, StringComparison.Ordinal);
                if (!groupOk)
                {
                    error = "grupo incorreto";
                    return false;
                }
                if (!authOk)
                {
                    error = "senha inválida";
                    return false;
                }
                error = string.Empty;
                return true;
            }
            catch
            {
                error = "erro de autenticação";
                return false;
            }
        }

        private static void HandleWrite(JsonElement root, StreamWriter writer)
        {
            try
            {
                var entity = root.GetProperty("entity").GetString() ?? "";
                var data = root.GetProperty("data");
                switch (entity)
                {
                    case "reserva":
                        var r = data.Deserialize<Reserva>();
                        if (r == null) throw new Exception("invalid reserva");
                        Database.InsertReserva(r);
                        break;
                    case "reserva_update":
                        var ru = data.Deserialize<Reserva>();
                        if (ru == null) throw new Exception("invalid reserva");
                        Database.UpdateReserva(ru);
                        break;
                    case "reserva_delete":
                        var rd = data.Deserialize<Reserva>();
                        if (rd == null) throw new Exception("invalid reserva");
                        Database.DeleteReserva(rd.Chave!, rd.Aluno, rd.Professor, rd.DataHora);
                        break;
                    case "relatorio":
                        var rel = data.Deserialize<Relatorio>();
                        if (rel == null) throw new Exception("invalid relatorio");
                        Database.InsertRelatorio(rel);
                        break;
                    case "chaves_replace":
                        var ch = data.Deserialize<System.Collections.Generic.List<Chave>>() ?? new();
                        Database.ReplaceAllChaves(ch);
                        break;
                    case "alunos_replace":
                        var aTable = JsonTableHelper.DeserializeToDataTable(data.GetRawText());
                        Database.ReplaceAllAlunos(aTable);
                        break;
                    case "professores_replace":
                        var pTable = JsonTableHelper.DeserializeToDataTable(data.GetRawText());
                        Database.ReplaceAllProfessores(pTable);
                        break;
                    default:
                        throw new Exception("unknown entity");
                }
                writer.WriteLine("{\"ok\":true}");
            }
            catch (Exception ex)
            {
                var msg = JsonSerializer.Serialize(new { ok = false, error = ex.Message });
                writer.WriteLine(msg);
            }
        }

        private static void HandlePull(JsonElement root, StreamWriter writer)
        {
            try
            {
                var entity = root.GetProperty("entity").GetString() ?? "";
                switch (entity)
                {
                    case "chaves":
                        var ch = Database.GetChaves();
                        writer.WriteLine(JsonSerializer.Serialize(new { ok = true, data = ch }));
                        break;
                    case "relatorio":
                        var rl = Database.GetRelatorios();
                        writer.WriteLine(JsonSerializer.Serialize(new { ok = true, data = rl }));
                        break;
                    case "reservas":
                        var rv = Database.GetReservas();
                        writer.WriteLine(JsonSerializer.Serialize(new { ok = true, data = rv }));
                        break;
                    case "alunos":
                        var at = JsonTableHelper.SerializeDataTable(Database.GetAlunosAsDataTable());
                        writer.WriteLine(JsonSerializer.Serialize(new { ok = true, data = at }));
                        break;
                    case "professores":
                        var pt = JsonTableHelper.SerializeDataTable(Database.GetProfessoresAsDataTable());
                        writer.WriteLine(JsonSerializer.Serialize(new { ok = true, data = pt }));
                        break;
                    default:
                        throw new Exception("unknown entity");
                }
            }
            catch (Exception ex)
            {
                var msg = JsonSerializer.Serialize(new { ok = false, error = ex.Message });
                writer.WriteLine(msg);
            }
        }
    }

    internal static class JsonTableHelper
    {
        public static System.Data.DataTable DeserializeToDataTable(string json)
        {
            // Espera um objeto com { columns: [name], rows: [[...]] }
            var model = JsonSerializer.Deserialize<TableModel>(json) ?? new TableModel();
            var dt = new System.Data.DataTable();
            foreach (var c in model.columns ?? Array.Empty<string>()) dt.Columns.Add(c);
            if (model.rows != null)
            {
                foreach (var row in model.rows)
                {
                    var dr = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count && i < row.Length; i++) dr[i] = row[i] ?? string.Empty;
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        public static object SerializeDataTable(System.Data.DataTable dt)
        {
            var cols = new string[dt.Columns.Count];
            for (int i = 0; i < cols.Length; i++) cols[i] = dt.Columns[i].ColumnName;
            var rows = new string[dt.Rows.Count][];
            for (int r = 0; r < rows.Length; r++)
            {
                var arr = new string[dt.Columns.Count];
                for (int c = 0; c < dt.Columns.Count; c++) arr[c] = System.Convert.ToString(dt.Rows[r][c]) ?? string.Empty;
                rows[r] = arr;
            }
            return new TableModel { columns = cols, rows = rows };
        }

        private class TableModel
        {
            public string[]? columns { get; set; }
            public string[][]? rows { get; set; }
        }
    }
}
