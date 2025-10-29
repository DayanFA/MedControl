using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class CadastroProfessoresForm : Form
    {
        private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = true };
        private DataTable _table = new DataTable();
        private Button _import = new Button { Text = "Importar Excel" };
        private Button _export = new Button { Text = "Exportar Excel" };
        private Button _salvarCaminho = new Button { Text = "Salvar Caminho" };

        public CadastroProfessoresForm()
        {
            Text = "Cadastro de Professores";
            Width = 900; Height = 600;
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40 };
            top.Controls.AddRange(new Control[] { _import, _export, _salvarCaminho });
            Controls.Add(top);
            Controls.Add(_grid);

            Load += (_, __) =>
            {
                var path = Database.GetConfig("caminho_professores");
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    _table = ExcelHelper.LoadToDataTable(path);
                }
                else
                {
                    _table = new DataTable();
                    _table.Columns.Add("Nome");
                }
                _grid.DataSource = _table;
            };

            _import.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    _table = ExcelHelper.LoadToDataTable(ofd.FileName);
                    _grid.DataSource = _table;
                    Database.SetConfig("caminho_professores", ofd.FileName);
                }
            };
            _export.Click += (_, __) =>
            {
                using var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "cadastro_professores.xlsx" };
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ExcelHelper.SaveDataTable(sfd.FileName, _table);
                    MessageBox.Show("Exportado com sucesso.");
                }
            };
            _salvarCaminho.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    Database.SetConfig("caminho_professores", ofd.FileName);
                    MessageBox.Show("Caminho salvo.");
                }
            };
        }
    }
}
