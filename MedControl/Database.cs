using Microsoft.Data.Sqlite;
using MySqlConnector;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;

namespace MedControl
{
    public static class Database
    {
        private static readonly string DbPath = GetSqlitePath();

        private static string GetSqlitePath()
        {
            // Persist the SQLite DB in %AppData%/MedControl/app.db to survive rebuilds/watch (bin cleanup)
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var dir = Path.Combine(appData, "MedControl");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                return Path.Combine(dir, "app.db");
            }
            catch
            {
                // Fallback to base directory if something goes wrong
                return Path.Combine(AppContext.BaseDirectory, "app.db");
            }
        }

        private static string Provider => (AppConfig.Instance.DbProvider ?? "sqlite").ToLowerInvariant();

        private static DbConnection CreateConnection()
        {
            if (Provider == "mysql")
            {
                var cs = AppConfig.Instance.MySqlConnectionString ?? string.Empty;
                return new MySqlConnection(cs);
            }
            return new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = DbPath }.ToString());
        }

        public static string GetCurrentDataSourceDescription()
        {
            if (Provider == "mysql")
            {
                try
                {
                    var cs = AppConfig.Instance.MySqlConnectionString ?? string.Empty;
                    var b = new MySqlConnectionStringBuilder(cs);
                    var server = string.IsNullOrWhiteSpace(b.Server) ? "(server?)" : b.Server;
                    var db = string.IsNullOrWhiteSpace(b.Database) ? "(db?)" : b.Database;
                    return $"MySQL: {server}/{db}";
                }
                catch { return "MySQL"; }
            }
            else
            {
                try { return "SQLite: " + DbPath; } catch { return "SQLite"; }
            }
        }

        public static void Setup()
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            if (Provider == "mysql")
            {
                // Ensure target database exists (CREATE DATABASE IF NOT EXISTS)
                try
                {
                    var raw = AppConfig.Instance.MySqlConnectionString ?? string.Empty;
                    var b = new MySqlConnectionStringBuilder(raw);
                    var dbName = b.Database;
                    if (!string.IsNullOrWhiteSpace(dbName))
                    {
                        var serverCsb = new MySqlConnectionStringBuilder(raw) { Database = string.Empty };
                        using var serverConn = new MySqlConnection(serverCsb.ConnectionString);
                        serverConn.Open();
                        using var createDb = serverConn.CreateCommand();
                        createDb.CommandText = $"CREATE DATABASE IF NOT EXISTS `{dbName}` CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci";
                        createDb.ExecuteNonQuery();
                    }
                }
                catch { /* ignore - may lack CREATE DATABASE privilege */ }

                var statements = new[]
                {
                    @"CREATE TABLE IF NOT EXISTS config (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        chave VARCHAR(255) UNIQUE,
                        valor TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS chaves (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        nome VARCHAR(255) UNIQUE,
                        num_copias INT,
                        descricao TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS alunos (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        uid VARCHAR(36) UNIQUE,
                        data JSON
                    )",
                    @"CREATE TABLE IF NOT EXISTS professores (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        uid VARCHAR(36) UNIQUE,
                        data JSON
                    )",
                    @"CREATE TABLE IF NOT EXISTS reservas (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        chave VARCHAR(255),
                        aluno VARCHAR(255),
                        professor VARCHAR(255),
                        data_hora VARCHAR(19),
                        em_uso TINYINT(1),
                        termo TEXT,
                        devolvido TINYINT(1),
                        data_devolucao VARCHAR(19)
                    )",
                    @"CREATE TABLE IF NOT EXISTS relatorio (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        chave VARCHAR(255),
                        aluno VARCHAR(255),
                        professor VARCHAR(255),
                        data_hora VARCHAR(19),
                        data_devolucao VARCHAR(19),
                        tempo_com_chave VARCHAR(255),
                        termo TEXT
                    )"
                };
                foreach (var sql in statements)
                {
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                var statements = new[]
                {
                    @"CREATE TABLE IF NOT EXISTS config (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        chave TEXT UNIQUE,
                        valor TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS chaves (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nome TEXT UNIQUE,
                        num_copias INTEGER,
                        descricao TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS alunos (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        uid TEXT UNIQUE,
                        data TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS professores (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        uid TEXT UNIQUE,
                        data TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS reservas (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        chave TEXT,
                        aluno TEXT,
                        professor TEXT,
                        data_hora TEXT,
                        em_uso INTEGER,
                        termo TEXT,
                        devolvido INTEGER,
                        data_devolucao TEXT
                    )",
                    @"CREATE TABLE IF NOT EXISTS relatorio (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        chave TEXT,
                        aluno TEXT,
                        professor TEXT,
                        data_hora TEXT,
                        data_devolucao TEXT,
                        tempo_com_chave TEXT,
                        termo TEXT
                    )"
                };
                foreach (var sql in statements)
                {
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private static void AddParam(DbCommand cmd, string name, object? value)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            p.Value = value ?? DBNull.Value;
            cmd.Parameters.Add(p);
        }

        public static List<Chave> GetChaves()
        {
            var list = new List<Chave>();
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT nome, num_copias, descricao FROM chaves ORDER BY nome";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                list.Add(new Chave
                {
                    Nome = rdr.GetString(0),
                    NumCopias = rdr.IsDBNull(1) ? 0 : rdr.GetInt32(1),
                    Descricao = rdr.IsDBNull(2) ? string.Empty : rdr.GetString(2)
                });
            }
            return list;
        }

        public static void UpsertChave(Chave c)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            if (Provider == "mysql")
            {
                cmd.CommandText = @"INSERT INTO chaves (nome, num_copias, descricao)
VALUES (@nome,@num,@desc)
ON DUPLICATE KEY UPDATE num_copias=VALUES(num_copias), descricao=VALUES(descricao)";
            }
            else
            {
                cmd.CommandText = @"INSERT INTO chaves (nome, num_copias, descricao)
VALUES (@nome,@num,@desc)
ON CONFLICT(nome) DO UPDATE SET num_copias=excluded.num_copias, descricao=excluded.descricao";
            }
            AddParam(cmd, "@nome", c.Nome);
            AddParam(cmd, "@num", c.NumCopias);
            AddParam(cmd, "@desc", c.Descricao ?? "");
            cmd.ExecuteNonQuery();
        }

        // Atualiza uma chave existente identificada pelo nome antigo; permite renomear a chave
        public static void UpdateChave(string oldNome, Chave c)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"UPDATE chaves SET nome=@newNome, num_copias=@num, descricao=@desc WHERE nome=@oldNome";
            AddParam(cmd, "@newNome", c.Nome);
            AddParam(cmd, "@num", c.NumCopias);
            AddParam(cmd, "@desc", c.Descricao ?? "");
            AddParam(cmd, "@oldNome", oldNome);
            cmd.ExecuteNonQuery();
        }

        public static void DeleteChave(string nome)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM chaves WHERE nome=@n";
            AddParam(cmd, "@n", nome);
            cmd.ExecuteNonQuery();
        }

        public static void ReplaceAllChaves(System.Collections.Generic.List<Chave> chaves)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var tx = conn.BeginTransaction();
            using (var del = conn.CreateCommand())
            {
                del.Transaction = tx;
                del.CommandText = "DELETE FROM chaves";
                del.ExecuteNonQuery();
            }

            foreach (var c in chaves)
            {
                using var ins = conn.CreateCommand();
                ins.Transaction = tx;
                ins.CommandText = "INSERT INTO chaves (nome, num_copias, descricao) VALUES (@nome,@num,@desc)";
                AddParam(ins, "@nome", c.Nome);
                AddParam(ins, "@num", c.NumCopias);
                AddParam(ins, "@desc", c.Descricao ?? "");
                ins.ExecuteNonQuery();
            }
            tx.Commit();
        }

        public static void InsertReserva(Reserva r)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT INTO reservas (chave, aluno, professor, data_hora, em_uso, termo, devolvido, data_devolucao)
                                VALUES (@ch,@al,@pr,@dt,@em,@te,@dev,@dd)";
            AddParam(cmd, "@ch", r.Chave);
            AddParam(cmd, "@al", r.Aluno ?? "");
            AddParam(cmd, "@pr", r.Professor ?? "");
            AddParam(cmd, "@dt", r.DataHora.ToString("dd/MM/yyyy HH:mm:ss"));
            AddParam(cmd, "@em", r.EmUso ? 1 : 0);
            AddParam(cmd, "@te", r.Termo ?? "");
            AddParam(cmd, "@dev", r.Devolvido ? 1 : 0);
            AddParam(cmd, "@dd", r.DataDevolucao.HasValue ? r.DataDevolucao.Value.ToString("dd/MM/yyyy HH:mm:ss") : (object?)DBNull.Value);
            cmd.ExecuteNonQuery();
        }

        public static void UpdateReserva(Reserva r)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"UPDATE reservas SET em_uso=@em, termo=@te, devolvido=@dev, data_devolucao=@dd
                                WHERE chave=@ch AND aluno=@al AND professor=@pr AND data_hora=@dt";
            AddParam(cmd, "@em", r.EmUso ? 1 : 0);
            AddParam(cmd, "@te", r.Termo ?? "");
            AddParam(cmd, "@dev", r.Devolvido ? 1 : 0);
            AddParam(cmd, "@dd", r.DataDevolucao.HasValue ? r.DataDevolucao.Value.ToString("dd/MM/yyyy HH:mm:ss") : (object?)DBNull.Value);
            AddParam(cmd, "@ch", r.Chave);
            AddParam(cmd, "@al", r.Aluno ?? "");
            AddParam(cmd, "@pr", r.Professor ?? "");
            AddParam(cmd, "@dt", r.DataHora.ToString("dd/MM/yyyy HH:mm:ss"));
            cmd.ExecuteNonQuery();
        }

        public static void DeleteReserva(string chave, string? aluno, string? professor, DateTime dataHora)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"DELETE FROM reservas WHERE chave=@ch AND aluno=@al AND professor=@pr AND data_hora=@dt";
            AddParam(cmd, "@ch", chave);
            AddParam(cmd, "@al", aluno ?? "");
            AddParam(cmd, "@pr", professor ?? "");
            AddParam(cmd, "@dt", dataHora.ToString("dd/MM/yyyy HH:mm:ss"));
            cmd.ExecuteNonQuery();
        }

        public static List<Reserva> GetReservas()
        {
            var list = new List<Reserva>();
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT chave, aluno, professor, data_hora, em_uso, termo, devolvido, data_devolucao FROM reservas";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                DateTime? dd = null;
                if (!rdr.IsDBNull(7))
                {
                    DateTime.TryParse(rdr.GetString(7), out var parsed);
                    dd = parsed;
                }
                DateTime.TryParse(rdr.GetString(3), out var data);
                list.Add(new Reserva
                {
                    Chave = rdr.GetString(0),
                    Aluno = rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1),
                    Professor = rdr.IsDBNull(2) ? string.Empty : rdr.GetString(2),
                    DataHora = data,
                    EmUso = !rdr.IsDBNull(4) && rdr.GetInt32(4) == 1,
                    Termo = rdr.IsDBNull(5) ? string.Empty : rdr.GetString(5),
                    Devolvido = !rdr.IsDBNull(6) && rdr.GetInt32(6) == 1,
                    DataDevolucao = dd
                });
            }
            return list;
        }

        public static void InsertRelatorio(Relatorio r)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"INSERT INTO relatorio (chave, aluno, professor, data_hora, data_devolucao, tempo_com_chave, termo)
                                VALUES (@c,@a,@p,@dh,@dd,@t,@te)";
            AddParam(cmd, "@c", r.Chave);
            AddParam(cmd, "@a", r.Aluno ?? "");
            AddParam(cmd, "@p", r.Professor ?? "");
            AddParam(cmd, "@dh", r.DataHora.ToString("dd/MM/yyyy HH:mm:ss"));
            AddParam(cmd, "@dd", r.DataDevolucao.HasValue ? r.DataDevolucao.Value.ToString("dd/MM/yyyy HH:mm:ss") : (object?)DBNull.Value);
            AddParam(cmd, "@t", r.TempoComChave ?? "");
            AddParam(cmd, "@te", r.Termo ?? "");
            cmd.ExecuteNonQuery();
        }

        public static List<Relatorio> GetRelatorios()
        {
            var list = new List<Relatorio>();
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT chave, aluno, professor, data_hora, data_devolucao, tempo_com_chave, termo FROM relatorio";
            using var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                DateTime.TryParse(rdr.GetString(3), out var dh);
                DateTime? dd = null;
                if (!rdr.IsDBNull(4))
                {
                    DateTime.TryParse(rdr.GetString(4), out var parsed);
                    dd = parsed;
                }
                list.Add(new Relatorio
                {
                    Chave = rdr.GetString(0),
                    Aluno = rdr.IsDBNull(1) ? string.Empty : rdr.GetString(1),
                    Professor = rdr.IsDBNull(2) ? string.Empty : rdr.GetString(2),
                    DataHora = dh,
                    DataDevolucao = dd,
                    TempoComChave = rdr.IsDBNull(5) ? string.Empty : rdr.GetString(5),
                    Termo = rdr.IsDBNull(6) ? string.Empty : rdr.GetString(6)
                });
            }
            return list;
        }

        public static void SetConfig(string chave, string valor)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            if (Provider == "mysql")
            {
                cmd.CommandText = @"INSERT INTO config (chave, valor) VALUES (@c,@v)
ON DUPLICATE KEY UPDATE valor=VALUES(valor)";
            }
            else
            {
                cmd.CommandText = @"INSERT INTO config (chave, valor) VALUES (@c,@v)
ON CONFLICT(chave) DO UPDATE SET valor=excluded.valor";
            }
            AddParam(cmd, "@c", chave);
            AddParam(cmd, "@v", valor);
            cmd.ExecuteNonQuery();
        }

        public static string? GetConfig(string chave)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT valor FROM config WHERE chave=@c";
            AddParam(cmd, "@c", chave);
            var result = cmd.ExecuteScalar();
            return result == null || result == DBNull.Value ? null : Convert.ToString(result);
        }

        // ===== Alunos (tabela flexível com dados em JSON) =====
        public static System.Data.DataTable GetAlunosAsDataTable()
        {
            var dt = new System.Data.DataTable();
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT uid, data FROM alunos";
            using var rdr = cmd.ExecuteReader();
            var rows = new List<(string uid, string json)>();
            while (rdr.Read())
            {
                var uid = rdr.IsDBNull(0) ? Guid.NewGuid().ToString() : rdr.GetString(0);
                var json = rdr.IsDBNull(1) ? "{}" : rdr.GetString(1);
                rows.Add((uid, json));
            }

            // Descobrir colunas (união das chaves do JSON)
            var allCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "_id" };
            foreach (var r in rows)
            {
                try
                {
                    var dict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(r.json) ?? new Dictionary<string, string>();
                    foreach (var k in dict.Keys) allCols.Add(k);
                }
                catch { }
            }
            foreach (var c in allCols) dt.Columns.Add(c);

            // Popular linhas
            foreach (var r in rows)
            {
                var dr = dt.NewRow();
                dr["_id"] = r.uid;
                try
                {
                    var dict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(r.json) ?? new Dictionary<string, string>();
                    foreach (var kv in dict)
                    {
                        if (!dt.Columns.Contains(kv.Key)) dt.Columns.Add(kv.Key);
                        dr[kv.Key] = kv.Value ?? string.Empty;
                    }
                }
                catch { }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public static void UpsertAluno(string uid, Dictionary<string, string> data)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            var json = System.Text.Json.JsonSerializer.Serialize(data);
            if (Provider == "mysql")
            {
                cmd.CommandText = @"INSERT INTO alunos (uid, data) VALUES (@u, CAST(@d AS JSON))
ON DUPLICATE KEY UPDATE data=VALUES(data)";
            }
            else
            {
                cmd.CommandText = @"INSERT INTO alunos (uid, data) VALUES (@u, @d)
ON CONFLICT(uid) DO UPDATE SET data=excluded.data";
            }
            AddParam(cmd, "@u", uid);
            AddParam(cmd, "@d", json);
            cmd.ExecuteNonQuery();
        }

        public static void DeleteAluno(string uid)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM alunos WHERE uid=@u";
            AddParam(cmd, "@u", uid);
            cmd.ExecuteNonQuery();
        }

        public static void ReplaceAllAlunos(System.Data.DataTable table)
        {
            using var conn = CreateConnection();
            conn.Open();
            using (var tx = conn.BeginTransaction())
            {
                // Limpa tudo
                using (var del = conn.CreateCommand())
                {
                    del.Transaction = tx;
                    del.CommandText = "DELETE FROM alunos";
                    del.ExecuteNonQuery();
                }

                // Insere todos
                foreach (System.Data.DataRow r in table.Rows)
                {
                    var uid = r.Table.Columns.Contains("_id") && r["_id"] != null && r["_id"] != DBNull.Value
                        ? Convert.ToString(r["_id"])!
                        : Guid.NewGuid().ToString();
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (System.Data.DataColumn c in table.Columns)
                    {
                        if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                        dict[c.ColumnName] = Convert.ToString(r[c]) ?? string.Empty;
                    }
                    using var ins = conn.CreateCommand();
                    ins.Transaction = tx;
                    if (Provider == "mysql")
                    {
                        ins.CommandText = "INSERT INTO alunos (uid, data) VALUES (@u, CAST(@d AS JSON))";
                    }
                    else
                    {
                        ins.CommandText = "INSERT INTO alunos (uid, data) VALUES (@u, @d)";
                    }
                    AddParam(ins, "@u", uid);
                    AddParam(ins, "@d", System.Text.Json.JsonSerializer.Serialize(dict));
                    ins.ExecuteNonQuery();
                }
                tx.Commit();
            }
        }

        // ===== Professores (mesmo modelo flexível em JSON) =====
        public static System.Data.DataTable GetProfessoresAsDataTable()
        {
            var dt = new System.Data.DataTable();
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT uid, data FROM professores";
            using var rdr = cmd.ExecuteReader();
            var rows = new List<(string uid, string json)>();
            while (rdr.Read())
            {
                var uid = rdr.IsDBNull(0) ? Guid.NewGuid().ToString() : rdr.GetString(0);
                var json = rdr.IsDBNull(1) ? "{}" : rdr.GetString(1);
                rows.Add((uid, json));
            }

            var allCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "_id" };
            foreach (var r in rows)
            {
                try
                {
                    var dict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(r.json) ?? new Dictionary<string, string>();
                    foreach (var k in dict.Keys) allCols.Add(k);
                }
                catch { }
            }
            foreach (var c in allCols) dt.Columns.Add(c);

            foreach (var r in rows)
            {
                var dr = dt.NewRow();
                dr["_id"] = r.uid;
                try
                {
                    var dict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(r.json) ?? new Dictionary<string, string>();
                    foreach (var kv in dict)
                    {
                        if (!dt.Columns.Contains(kv.Key)) dt.Columns.Add(kv.Key);
                        dr[kv.Key] = kv.Value ?? string.Empty;
                    }
                }
                catch { }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public static void UpsertProfessor(string uid, Dictionary<string, string> data)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            var json = System.Text.Json.JsonSerializer.Serialize(data);
            if (Provider == "mysql")
            {
                cmd.CommandText = @"INSERT INTO professores (uid, data) VALUES (@u, CAST(@d AS JSON))
ON DUPLICATE KEY UPDATE data=VALUES(data)";
            }
            else
            {
                cmd.CommandText = @"INSERT INTO professores (uid, data) VALUES (@u, @d)
ON CONFLICT(uid) DO UPDATE SET data=excluded.data";
            }
            AddParam(cmd, "@u", uid);
            AddParam(cmd, "@d", json);
            cmd.ExecuteNonQuery();
        }

        public static void DeleteProfessor(string uid)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM professores WHERE uid=@u";
            AddParam(cmd, "@u", uid);
            cmd.ExecuteNonQuery();
        }

        public static void ReplaceAllProfessores(System.Data.DataTable table)
        {
            using var conn = CreateConnection();
            conn.Open();
            using (var tx = conn.BeginTransaction())
            {
                using (var del = conn.CreateCommand())
                {
                    del.Transaction = tx;
                    del.CommandText = "DELETE FROM professores";
                    del.ExecuteNonQuery();
                }

                foreach (System.Data.DataRow r in table.Rows)
                {
                    var uid = r.Table.Columns.Contains("_id") && r["_id"] != null && r["_id"] != DBNull.Value
                        ? Convert.ToString(r["_id"])!
                        : Guid.NewGuid().ToString();
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (System.Data.DataColumn c in table.Columns)
                    {
                        if (string.Equals(c.ColumnName, "_id", StringComparison.OrdinalIgnoreCase)) continue;
                        dict[c.ColumnName] = Convert.ToString(r[c]) ?? string.Empty;
                    }
                    using var ins = conn.CreateCommand();
                    ins.Transaction = tx;
                    if (Provider == "mysql")
                    {
                        ins.CommandText = "INSERT INTO professores (uid, data) VALUES (@u, CAST(@d AS JSON))";
                    }
                    else
                    {
                        ins.CommandText = "INSERT INTO professores (uid, data) VALUES (@u, @d)";
                    }
                    AddParam(ins, "@u", uid);
                    AddParam(ins, "@d", System.Text.Json.JsonSerializer.Serialize(dict));
                    ins.ExecuteNonQuery();
                }
                tx.Commit();
            }
        }
    }
}
