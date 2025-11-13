using System;

namespace MedControl
{
    public class Chave
    {
        public string Nome { get; set; } = string.Empty;
        public int NumCopias { get; set; }
        public string Descricao { get; set; } = string.Empty;
    }

    public class Reserva
    {
        public string Chave { get; set; } = string.Empty;
        public string? Aluno { get; set; }
        public string? Professor { get; set; }
        public DateTime DataHora { get; set; }
        public bool EmUso { get; set; }
        public string? Termo { get; set; }
        public bool Devolvido { get; set; }
        public DateTime? DataDevolucao { get; set; }
        // Controle de sincronização
        public DateTime LastModifiedUtc { get; set; } = DateTime.MinValue;
        public bool Deleted { get; set; } = false;
    }

    public class Relatorio
    {
        public string Chave { get; set; } = string.Empty;
        public string? Aluno { get; set; }
        public string? Professor { get; set; }
        public DateTime DataHora { get; set; }
        public DateTime? DataDevolucao { get; set; }
        public string? TempoComChave { get; set; }
        public string? Termo { get; set; }
    }
}
