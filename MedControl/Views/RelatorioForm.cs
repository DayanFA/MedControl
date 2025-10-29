using ClosedXML.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class RelatorioForm : Form
    {
        private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false };
        private BindingSource _bs = new BindingSource();
        private Button _export = new Button { Text = "Exportar Relatório" };

        public RelatorioForm()
        {
            Text = "Relatórios";
            Width = 1000;
            Height = 600;

            var panelTop = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40 };
            panelTop.Controls.Add(_export);

            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Chave", DataPropertyName = nameof(Relatorio.Chave), Width = 120 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Aluno", DataPropertyName = nameof(Relatorio.Aluno), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Professor", DataPropertyName = nameof(Relatorio.Professor), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Data e Hora", DataPropertyName = nameof(Relatorio.DataHora), Width = 160 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Data de Devolução", DataPropertyName = nameof(Relatorio.DataDevolucao), Width = 160 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Tempo com a Chave", DataPropertyName = nameof(Relatorio.TempoComChave), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Termo de Compromisso", DataPropertyName = nameof(Relatorio.Termo), Width = 220 });

            Controls.Add(_grid);
            Controls.Add(panelTop);

            Load += (_, __) => RefreshGrid();
            _export.Click += (_, __) => Exportar();
        }

        private void RefreshGrid()
        {
            _bs.DataSource = Database.GetRelatorios();
            _grid.DataSource = _bs;
        }

        private void Exportar()
        {
            var dados = Database.GetRelatorios();
            if (dados.Count == 0)
            {
                MessageBox.Show("Não há dados no relatório para exportar.");
                return;
            }

            using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "Relatorio.xlsx" };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Relatório");
            // Cabeçalhos
            string[] headers = { "Chave", "Aluno", "Professor", "Data e Hora", "Data de Devolução", "Tempo com a Chave", "Termo de Compromisso", "Item Atribuído" };
            for (int i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];

            // Linhas
            int row = 2;
            foreach (var r in dados)
            {
                ws.Cell(row, 1).Value = r.Chave;
                ws.Cell(row, 2).Value = r.Aluno ?? "";
                ws.Cell(row, 3).Value = r.Professor ?? "";
                ws.Cell(row, 4).Value = r.DataHora;
                ws.Cell(row, 5).Value = r.DataDevolucao;
                ws.Cell(row, 6).Value = r.TempoComChave ?? "";

                var termo = r.Termo ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(termo) && File.Exists(termo))
                {
                    var cell = ws.Cell(row, 7);
                    cell.Value = "Abrir PDF";
                    cell.GetHyperlink().ExternalAddress = new Uri(termo);
                    ws.Cell(row, 8).Value = "Sim";
                }
                else
                {
                    ws.Cell(row, 7).Value = termo;
                    ws.Cell(row, 8).Value = "Não";
                }
                row++;
            }

            wb.SaveAs(sfd.FileName);
            MessageBox.Show($"Relatório exportado com sucesso para '{sfd.FileName}'.");
        }
    }
}
