using System;
using System.Linq;
using System.Windows.Forms;

namespace MedControl.Views
{
    public class CadastroChavesForm : Form
    {
        private DataGridView _grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false };
        private BindingSource _bs = new BindingSource();

        public CadastroChavesForm()
        {
            Text = "Cadastro de Chaves";
            Width = 800;
            Height = 600;

            var panelTop = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40 };
            var btnAdd = new Button { Text = "Adicionar" };
            var btnEdit = new Button { Text = "Editar" };
            var btnDel = new Button { Text = "Deletar" };
            btnAdd.Click += (_, __) => AddChave();
            btnEdit.Click += (_, __) => EditChave();
            btnDel.Click += (_, __) => DeleteChave();
            panelTop.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDel });

            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nome", DataPropertyName = nameof(Chave.Nome), Width = 200 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Número de Cópias", DataPropertyName = nameof(Chave.NumCopias), Width = 150 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Descrição", DataPropertyName = nameof(Chave.Descricao), Width = 300 });

            Controls.Add(_grid);
            Controls.Add(panelTop);

            Load += (_, __) => RefreshGrid();
        }

        private void RefreshGrid()
        {
            var data = Database.GetChaves();
            _bs.DataSource = data;
            _grid.DataSource = _bs;
        }

        private void AddChave()
        {
            var dlg = new ChaveDialog();
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                Database.UpsertChave(dlg.Result);
                RefreshGrid();
            }
        }

        private void EditChave()
        {
            if (_bs.Current is Chave c)
            {
                var dlg = new ChaveDialog(c);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    Database.UpsertChave(dlg.Result);
                    RefreshGrid();
                }
            }
        }

        private void DeleteChave()
        {
            if (_bs.Current is Chave c)
            {
                if (MessageBox.Show($"Deletar a chave '{c.Nome}'?", "Deletar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Database.DeleteChave(c.Nome);
                    RefreshGrid();
                }
            }
        }

        private class ChaveDialog : Form
        {
            private TextBox _nome = new TextBox();
            private NumericUpDown _num = new NumericUpDown { Minimum = 0, Maximum = 1000 };
            private TextBox _desc = new TextBox();
            public Chave Result { get; private set; } = new Chave();

            public ChaveDialog(Chave? c = null)
            {
                Text = c == null ? "Adicionar Chave" : "Editar Chave";
                Width = 400;
                Height = 240;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;

                var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 4, Padding = new Padding(10) };
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

                table.Controls.Add(new Label { Text = "Nome:" }, 0, 0);
                table.Controls.Add(_nome, 1, 0);
                table.Controls.Add(new Label { Text = "Número de Cópias:" }, 0, 1);
                table.Controls.Add(_num, 1, 1);
                table.Controls.Add(new Label { Text = "Descrição:" }, 0, 2);
                _desc.Multiline = true; _desc.Dock = DockStyle.Fill;
                table.Controls.Add(_desc, 1, 2);

                var panelButtons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
                var ok = new Button { Text = "Salvar", DialogResult = DialogResult.OK };
                var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel };
                panelButtons.Controls.AddRange(new Control[] { ok, cancel });
                table.Controls.Add(panelButtons, 0, 3); table.SetColumnSpan(panelButtons, 2);

                Controls.Add(table);

                if (c != null)
                {
                    _nome.Text = c.Nome;
                    _num.Value = c.NumCopias;
                    _desc.Text = c.Descricao;
                }

                AcceptButton = ok;
                CancelButton = cancel;

                ok.Click += (_, __) =>
                {
                    Result = new Chave { Nome = _nome.Text.Trim(), NumCopias = (int)_num.Value, Descricao = _desc.Text };
                    if (string.IsNullOrWhiteSpace(Result.Nome))
                    {
                        MessageBox.Show("O campo 'Nome' é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                    }
                };
            }
        }
    }
}
