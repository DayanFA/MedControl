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
        private static readonly string DbPath = Path.Combine(AppContext.BaseDirectory, "app.db");

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

        public static void Setup()
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            if (Provider == "mysql")
            {
                cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS config (
  id INT AUTO_INCREMENT PRIMARY KEY,
  chave VARCHAR(255) UNIQUE,
  valor TEXT
);
CREATE TABLE IF NOT EXISTS chaves (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nome VARCHAR(255) UNIQUE,
  num_copias INT,
  descricao TEXT
);
CREATE TABLE IF NOT EXISTS reservas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  chave VARCHAR(255),
  aluno VARCHAR(255),
  professor VARCHAR(255),
  data_hora VARCHAR(19),
  em_uso TINYINT(1),
  termo TEXT,
  devolvido TINYINT(1),
  data_devolucao VARCHAR(19)
);
CREATE TABLE IF NOT EXISTS relatorio (
  id INT AUTO_INCREMENT PRIMARY KEY,
  chave VARCHAR(255),
  aluno VARCHAR(255),
  professor VARCHAR(255),
  data_hora VARCHAR(19),
  data_devolucao VARCHAR(19),
  tempo_com_chave VARCHAR(255),
  termo TEXT
);";
            }
            else
            {
                cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS config (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chave TEXT UNIQUE,
    valor TEXT
);
CREATE TABLE IF NOT EXISTS chaves (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT UNIQUE,
    num_copias INTEGER,
    descricao TEXT
);
CREATE TABLE IF NOT EXISTS reservas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chave TEXT,
    aluno TEXT,
    professor TEXT,
    data_hora TEXT,
    em_uso INTEGER,
    termo TEXT,
    devolvido INTEGER,
    data_devolucao TEXT
);
CREATE TABLE IF NOT EXISTS relatorio (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chave TEXT,
    aluno TEXT,
    professor TEXT,
    data_hora TEXT,
    data_devolucao TEXT,
    tempo_com_chave TEXT,
    termo TEXT
);";
            }
            cmd.ExecuteNonQuery();
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

        public static void DeleteChave(string nome)
        {
            using var conn = CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM chaves WHERE nome=@n";
            AddParam(cmd, "@n", nome);
            cmd.ExecuteNonQuery();
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
    }
}
