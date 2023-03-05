using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.Common;
using System.IO;
using System.Diagnostics;

namespace Библиотека_МГЭПТК
{
    public partial class MainForm : Form
    {
        public OleDbConnection Conn;
        private bool LoadOrAction;
        public MainForm()
        {
            if (settings.Default.isPass)
            {
                using (login log = new login())
                {
                    if (log.ShowDialog() != DialogResult.OK)
                    {
                        Application.Exit();
                    }
                }
            }
            InitializeComponent();
            if (!File.Exists($"{Environment.CurrentDirectory}\\db.mdb"))
            {
                File.WriteAllBytes($"{Environment.CurrentDirectory}\\db.mdb", Properties.Resources.dbEMPTY);
            }
            Conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;");
            Conn.Open();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            TreeUpdate();
        }
        private void TreeUpdate()
        {
            string query = "SELECT * FROM [Специальности]";
            OleDbCommand command = new OleDbCommand(query, Conn);
            OleDbDataReader reader = command.ExecuteReader();
            treeView1.Nodes.Clear();
            treeView1.Nodes.Add(new TreeNode("Все книги"));
            while (reader.Read())
            {
                TreeNode lvl = new TreeNode(reader[1].ToString());
                if (treeView1.Nodes.Cast<TreeNode>().Where(x => x.Text == lvl.Text).Count() == 0)
                {
                    treeView1.Nodes.Add(lvl);
                }
                TreeNode spec = new TreeNode(reader[2].ToString());

                string subQuery = $"SELECT * FROM [Дисциплины] WHERE [Специальность] LIKE '{reader[0]}'";
                OleDbCommand subCommand = new OleDbCommand(subQuery, Conn);
                OleDbDataReader subReader = subCommand.ExecuteReader();
                while (subReader.Read())
                {
                    TreeNode disciplina = new TreeNode(subReader[2].ToString());
                    spec.Nodes.Add(disciplina);

                }
                subReader.Close();

                foreach (TreeNode node in treeView1.Nodes)
                {
                    if (node.Text == lvl.Text)
                    {
                        node.Nodes.Add(spec);
                    }
                }
            }
            reader.Close();
        }
        private void DataUpdate()
        {
            string query = "SELECT Книги.Код, Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Дисциплины.Название\r\n" +
                "FROM Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина\r\n" +
                "ORDER BY Книги.Код;\r\n";
            OleDbDataAdapter oleDb = new OleDbDataAdapter(query, Conn);
            DataSet ds = new DataSet();
            oleDb.Fill(ds, "books");
            ds.Tables[0].Columns[ds.Tables[0].Columns.Count - 1].ColumnName = "Дисциплина";
            ds.Tables[0].Columns[ds.Tables[0].Columns.IndexOf("Книги.Название")].ColumnName = "Название";
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "books";
            changeStatus(2, 0);
        }
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string[] spath = treeView1.SelectedNode.FullPath.Split(@"\".ToCharArray());
            this.Text = $"Главная [{treeView1.SelectedNode.FullPath}]";
            string query;
            DataSet ds = new DataSet();
            ds.Clear();
            if (spath[0] == "Все книги") { DataUpdate(); }
            else
            {
                switch (spath.Length)
                {
                    case 1:
                        query = $"SELECT [Код] FROM [Специальности] WHERE [Уровень] LIKE '{spath[0]}'";
                        OleDbCommand command = new OleDbCommand(query, Conn);
                        OleDbDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string subQuery = $"SELECT [Код] FROM [Дисциплины] WHERE [Специальность] LIKE '{reader[0]}'";
                            OleDbCommand subCommand = new OleDbCommand(subQuery, Conn);
                            OleDbDataReader subReader = subCommand.ExecuteReader();
                            while (subReader.Read())
                            {
                                string subQuery2 = $"SELECT Книги.Код, Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Дисциплины.Название\r\nFROM Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина\r\nWHERE (((Книги.Дисциплина) Like '{subReader[0]}'))\r\nORDER BY Книги.Код;\r\n";
                                OleDbDataAdapter oleDb = new OleDbDataAdapter(subQuery2, Conn);
                                oleDb.Fill(ds, "books");

                                dataGridView1.DataSource = ds;
                                dataGridView1.DataMember = "books";
                            }
                            subReader.Close();
                        }
                        reader.Close();
                        break;
                    case 2:
                        query = $"SELECT [Код] FROM [Специальности] WHERE [Название] LIKE '{spath[1]}' AND [Уровень] LIKE '{spath[0]}'";
                        command = new OleDbCommand(query, Conn);
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string subQuery = $"SELECT [Код] FROM [Дисциплины] WHERE [Специальность] LIKE '{reader[0]}'";
                            OleDbCommand subCommand = new OleDbCommand(subQuery, Conn);
                            OleDbDataReader subReader = subCommand.ExecuteReader();
                            while (subReader.Read())
                            {
                                string subQuery2 = $"SELECT Книги.Код, Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Дисциплины.Название\r\nFROM Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина\r\nWHERE (((Книги.Дисциплина) Like '{subReader[0]}'))\r\nORDER BY Книги.Код;\r\n";
                                OleDbDataAdapter oleDb = new OleDbDataAdapter(subQuery2, Conn);
                                oleDb.Fill(ds, "books");

                                dataGridView1.DataSource = ds;
                                dataGridView1.DataMember = "books";
                            }
                            subReader.Close();
                        }
                        reader.Close();
                        break;
                    case 3:
                        query = $"SELECT [Код] FROM [Специальности] WHERE [Название] LIKE '{spath[1]}' AND [Уровень] LIKE '{spath[0]}'";
                        command = new OleDbCommand(query, Conn);
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string subQuery = $"SELECT [Код] FROM [Дисциплины] WHERE [Специальность] LIKE '{reader[0]}' AND [Название] LIKE '{spath[2]}'";
                            OleDbCommand subCommand = new OleDbCommand(subQuery, Conn);
                            OleDbDataReader subReader = subCommand.ExecuteReader();
                            while (subReader.Read())
                            {
                                string subQuery2 = $"SELECT Книги.Код, Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Дисциплины.Название\r\nFROM Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина\r\nWHERE (((Книги.Дисциплина) Like '{subReader[0]}'))\r\nORDER BY Книги.Код;\r\n";
                                OleDbDataAdapter oleDb = new OleDbDataAdapter(subQuery2, Conn);
                                oleDb.Fill(ds, "books");

                                dataGridView1.DataSource = ds;
                                dataGridView1.DataMember = "books";
                            }
                            subReader.Close();
                        }
                        reader.Close();
                        break;
                }
                ds.Tables[0].Columns[ds.Tables[0].Columns.Count - 1].ColumnName = "Дисциплина";
                ds.Tables[0].Columns[ds.Tables[0].Columns.IndexOf("Книги.Название")].ColumnName = "Название";
            }
            dataGridView1.Sort(this.dataGridView1.Columns["Код"], ListSortDirection.Ascending);

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                удалитьToolStripMenuItem.Enabled = false;
                изменитьToolStripMenuItem.Enabled = false;
                выдатьToolStripMenuItem.Enabled = false;
                показатьИстриюToolStripMenuItem.Enabled = false;
            }
            else
            {
                удалитьToolStripMenuItem.Enabled = true;
                изменитьToolStripMenuItem.Enabled = true;
                выдатьToolStripMenuItem.Enabled = true;
                показатьИстриюToolStripMenuItem.Enabled = true;
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeStatus(3, 11);

        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sCount = dataGridView1.SelectedRows.Count;
            if (MessageBox.Show($"Вы уверены что хотите удалить {sCount} записей?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string query = $"DELETE FROM [Книги] WHERE [Код] LIKE '{row.Cells[0].Value}'";
                    OleDbCommand command = new OleDbCommand(query, Conn);
                    command.ExecuteNonQuery();
                }
                var b = treeView1.SelectedNode;
                treeView1.SelectedNode = null;
                treeView1.SelectedNode = b;
                MessageBox.Show($"Было удаленно {sCount} записей.", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                changeStatus(3, 1);
            }
        }

        private void поискToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Conn.Close();
        }

        private void ученикиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (History h = new History())
            {
                changeStatus(3, 8);
                h.isSearch = false;
                if (h.ShowDialog() == DialogResult.OK)
                {
                    DataUpdate();
                }
            }
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeUpdate();
            DataUpdate();
        }

        private void выдатьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            using (GiveBooks GB = new GiveBooks())
            {
                changeStatus(3, 5);
                if (GB.ShowDialog() == DialogResult.OK)
                {
                    changeStatus(3, 4);
                    DataUpdate();
                }
            }
        }

        private void показатьИстриюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (History h = new History())
            {
                changeStatus(3, 8);
                h.isSearch = true;
                h.sId = new List<string>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    h.sId.Add(row.Cells[0].Value.ToString());
                }
                if (h.ShowDialog() == DialogResult.OK)
                {
                    DataUpdate();
                }
            }
        }

        private void выдатьToolStripMenuItem_Click(object sender, EventArgs e)
        {


            using (GiveBooks GB = new GiveBooks())
            {
                changeStatus(3, 5);
                GB.ids = new List<int>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    GB.ids.Add(Convert.ToInt32(row.Cells[0].Value));
                }
                if (GB.ShowDialog() == DialogResult.OK)
                {
                    changeStatus(3, 4);
                    DataUpdate();
                }
            }
        }

        private void добавитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            using (AddEdBook AEB = new AddEdBook())
            {
                changeStatus(3, 6);
                AEB.isEdit = false;
                if (AEB.ShowDialog() == DialogResult.OK)
                {
                    changeStatus(3, 2);
                    DataUpdate();
                }
            }
        }

        private void изменитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            using (AddEdBook AEB = new AddEdBook())
            {
                changeStatus(3, 6);
                AEB.isEdit = true;
                AEB.rowsID = new List<string>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {

                    AEB.rowsID.Add(row.Cells[0].Value.ToString());
                }

                if (AEB.ShowDialog() == DialogResult.OK)
                {
                    changeStatus(3, 3);
                    DataUpdate();
                }
            }
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeStatus(3, 9);
            options op = new options();
            op.ShowDialog();
        }

        private void поискToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SearchBooks SB = new SearchBooks())
            {
                changeStatus(3, 7);
                if (SB.ShowDialog() == DialogResult.OK)
                {
                    DataUpdate();
                }
            }
        }

        private async void changeStatus(byte p, byte a)
        {
            switch (p)
            {
                case 1:
                    {
                        if (LoadOrAction) { await Task.Delay(5000); }
                        if (dataGridView1.SelectedRows.Count == 1)
                        {
                            label1.Text = $"Выбрана запись №{dataGridView1.SelectedRows[0].Cells[0].Value} из {dataGridView1.Rows.Count}";
                        }
                        else if (dataGridView1.SelectedRows.Count > 1)
                        {
                            label1.Text = $"Выбрано {dataGridView1.SelectedRows[0].Cells[0].Value} записей из {dataGridView1.Rows.Count}";
                        }
                        else { label1.Text = $"Не выбрана ни одна запись"; }
                        LoadOrAction = false;
                        break;
                    }
                case 2:
                    {
                        label1.Text = $"Было загружено {dataGridView1.Rows.Count} записей";
                        LoadOrAction = true;
                        break;
                    }
                case 3:
                    {
                        switch (a)
                        {
                            case 1:
                                {
                                    label1.Text = $"Дейсвие \"удаление\" было выполнено";
                                    break;
                                }
                            case 2:
                                {
                                    label1.Text = $"Дейсвие \"добавление\" было выполнено";
                                    break;
                                }
                            case 3:
                                {
                                    label1.Text = $"Дейсвие \"изменение\" было выполнено";
                                    break;
                                }
                            case 4:
                                {
                                    label1.Text = $"Дейсвие \"выдача\" было выполнено";
                                    break;

                                }
                            case 5:
                                {
                                    label1.Text = $"Форма выдачи открывается...";
                                    break;
                                }
                            case 6:
                                {
                                    label1.Text = $"Форма изменения/добавления открывается...";
                                    break;
                                }
                            case 7:
                                {
                                    label1.Text = $"Форма поиска открывается...";
                                    break;
                                }
                            case 8:
                                {
                                    label1.Text = $"Форма истории открывается...";
                                    break;
                                }
                            case 9:
                                {
                                    label1.Text = $"Форма настроек открывается...";
                                    break;
                                }
                            case 10:
                                {
                                    label1.Text = $"Форма справки открывается...";
                                    break;
                                }
                            case 11:
                                {
                                    label1.Text = $"Приложение закрывается...";
                                    await Task.Delay(1000);
                                    Application.Exit();
                                    break;
                                }
                        }
                        LoadOrAction = true;
                        break;
                    }
                default:
                    {
                        label1.Text = $"{DateTime.Now}";
                        break;
                    }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

            changeStatus(1, 0);
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!File.Exists($"{Environment.CurrentDirectory}\\spravka.pdf"))
            {
                File.WriteAllBytes($"{Environment.CurrentDirectory}\\spravka.pdf", Properties.Resources.spravka);
            }
            Process.Start($"{Environment.CurrentDirectory}\\spravka.pdf");
        }


   
    }
}