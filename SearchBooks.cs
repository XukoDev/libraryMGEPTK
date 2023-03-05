using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Runtime.InteropServices;

namespace Библиотека_МГЭПТК
{
    public partial class SearchBooks : Form
    {
        MainForm MF = new MainForm();
        public SearchBooks()
        {
            InitializeComponent();
        }

        private void SearchBooks_Load(object sender, EventArgs e)
        {
            TreeUpdate();
        }
        private void TreeUpdate()
        {
            string query = "SELECT * FROM [Специальности]";
            OleDbCommand command = new OleDbCommand(query, MF.Conn);
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
                OleDbCommand subCommand = new OleDbCommand(subQuery, MF.Conn);
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

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Checked == true)
            {
                string[] spath = e.Node.FullPath.Split(@"\".ToCharArray());
                if (spath[0] == "Все книги")
                {
                    foreach (TreeNode node in treeView1.Nodes)
                    {
                        if (node.Text != "Все книги")
                        {
                            foreach (TreeNode childNode in node.Nodes)
                            {
                                foreach (TreeNode cchildNode in childNode.Nodes) { cchildNode.Checked = false; }
                                childNode.Checked = false;
                            }
                            node.Checked = false;
                        }
                    }
                }
                else
                {
                    switch (spath.Length)
                    {
                        case 1:
                            int act = 0;
                            foreach (TreeNode node in treeView1.Nodes)
                            {
                                if (node.Text == "Все книги")
                                { node.Checked = false; }
                                foreach (TreeNode childs in node.Nodes)
                                {
                                    childs.Checked = false;
                                    foreach (TreeNode cchilds in childs.Nodes)
                                    {
                                        cchilds.Checked = false;
                                    }
                                }
                                if (node.Checked) act++;
                            }
                            if (act >= 2)
                            {
                                foreach (TreeNode node in treeView1.Nodes)
                                {
                                    if (node.Text == "Все книги")
                                    { node.Checked = true; }
                                    else { node.Checked = false; }
                                }
                            }
                            break;
                        case 2:
                            foreach (TreeNode node in treeView1.Nodes)
                            {
                                if (node.Text == "Все книги" || node.Text == spath[0])
                                { node.Checked = false; }
                                if (node.Text == spath[0])
                                {
                                    act = 0;
                                    foreach (TreeNode childs in node.Nodes)
                                    {
                                        if (childs.Text == spath[1])
                                        {
                                            foreach (TreeNode cchilds in childs.Nodes)
                                            {
                                                cchilds.Checked = false;
                                            }
                                        }

                                        if (childs.Checked == true)
                                        {
                                            act++;
                                        }
                                    }
                                    if (act == node.Nodes.Count)
                                    {
                                        node.Checked = true;
                                        foreach (TreeNode childs in node.Nodes)
                                        {
                                            childs.Checked = false;
                                        }
                                    }
                                }
                            }
                            break;
                        case 3:
                            foreach (TreeNode node in treeView1.Nodes)
                            {
                                if (node.Text == "Все книги" || node.Text == spath[0])
                                { node.Checked = false; }
                                foreach (TreeNode childs in node.Nodes)
                                {
                                    if (childs.Text == spath[1])
                                    {
                                        act = 0;
                                        foreach (TreeNode cchilds in childs.Nodes)
                                        {
                                            if (cchilds.Checked == true)
                                            {
                                                act++;
                                            }
                                        }
                                        childs.Checked = false;
                                        if (act == childs.Nodes.Count)
                                        {
                                            childs.Checked = true;
                                        }
                                    }

                                }
                            }
                            break;
                    }
                }
            }
            int chk = 0;
            foreach (TreeNode node in treeView1.Nodes)
            {
                if (node.Checked == true) chk++;
                foreach (TreeNode cnode in node.Nodes)
                {
                    if (cnode.Checked == true) chk++;
                    foreach (TreeNode ccnode in cnode.Nodes)
                    {
                        if (ccnode.Checked == true) chk++;
                    }
                }
            }
            if (chk > 0) button1.Enabled = true;
            else button1.Enabled = false;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sq = new StringBuilder();
            StringBuilder query = new StringBuilder();
            StringBuilder subQuery = new StringBuilder();
            DataSet ds = new DataSet();
            foreach (string it in checkedListBox1.Items)
            {
                bool isChecked = checkedListBox1.GetItemChecked(checkedListBox1.FindStringExact(it));
                if (isChecked && it == "Код")
                {
                    subQuery.Append($"OR ((Книги.Код) Like \"%%{textBox1.Text}%%\")");
                }
                else if (isChecked && it == "Название")
                {
                    subQuery.Append($"OR ((Книги.Название) Like \"%%{textBox1.Text}%%\")");
                }
                else if (isChecked && it == "Год")
                {
                    subQuery.Append($"OR ((Книги.Год) Like \"%%{textBox1.Text}%%\")");
                }
                else if (isChecked && it == "Автор")
                {
                    subQuery.Append($"OR ((Книги.Автор) Like \"%%{textBox1.Text}%%\")");
                }
                else if (isChecked && it == "Кол-во (на складе)")
                {
                    subQuery.Append($"OR ((Книги.[Кол-во (на складе)]) Like \"%%{textBox1.Text}%%\")");
                }
                else if (isChecked && it == "Кол-во (всего)")
                {
                    subQuery.Append($"OR ((Книги.[Кол-во (всего)]) Like \"%%{textBox1.Text}%%\")");
                }
                else if (isChecked && it == "Дисциплина")
                {
                    subQuery.Append($"OR ((Дисциплины.Название) Like \"%%{textBox1.Text}%%\")");
                }
            }
            if (subQuery.Length > 0)
            {
                subQuery.Remove(0, 3);
            }
            foreach (TreeNode node in treeView1.Nodes)
            {
                if (node.Checked == true && node.Text != "Все книги")
                {
                    sq.Append(node.Text + "|");
                }
                foreach (TreeNode cnode in node.Nodes)
                {
                    if (cnode.Checked == true)
                    {
                        sq.Append(cnode.FullPath + "|");
                    }
                    foreach (TreeNode ccnode in cnode.Nodes)
                    {
                        if (ccnode.Checked == true)
                        {
                            sq.Append(ccnode.FullPath + "|");
                        }
                    }
                }
            }
            if (sq.Length > 0)
            {

                sq.Remove(sq.Length - 1, 1);
                string[] fps = sq.ToString().Split('|');
                foreach (string fp in fps)
                {
                    query.Clear();
                    string[] p = fp.ToString().Split(@"\".ToCharArray());
                    query.AppendLine("SELECT Книги.Код, Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Дисциплины.Название\r\nFROM Специальности INNER JOIN (Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина) ON Специальности.Код = Дисциплины.Специальность");
                    switch (p.Length)
                    {
                        case 1:
                            query.AppendLine($"WHERE (((Специальности.Уровень)=\"{p[0]}\"))"); break;
                        case 2:
                            query.AppendLine($"WHERE (((Специальности.Уровень)=\"{p[0]}\") AND ((Специальности.Название)=\"{p[1]}\"))"); break;
                        case 3:
                            query.AppendLine($"WHERE (((Дисциплины.Название)=\"{p[2]}\") AND ((Специальности.Уровень)=\"{p[0]}\") AND ((Специальности.Название)=\"{p[1]}\"))"); break;
                    }
                    if (subQuery.Length > 0)
                    {
                        query.Append($" AND ({subQuery})");
                    }
                    query.AppendLine("ORDER BY Книги.Код;");

                    OleDbDataAdapter oleDb = new OleDbDataAdapter(query.ToString(), MF.Conn);
                    oleDb.Fill(ds, "books");
                }
            }
            else
            {

                query.Append("SELECT Книги.Код, Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Дисциплины.Название\r\nFROM Специальности INNER JOIN (Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина) ON Специальности.Код = Дисциплины.Специальность \r\n");
                if (subQuery.Length > 0)
                {
                    query.Append($"WHERE ({subQuery}) ");
                }
                query.AppendLine("ORDER BY Книги.Код;");
                OleDbDataAdapter oleDb = new OleDbDataAdapter(query.ToString(), MF.Conn);
                oleDb.Fill(ds, "books");
            }
            ds.Tables[0].Columns[ds.Tables[0].Columns.Count - 1].ColumnName = "Дисциплина";
            ds.Tables[0].Columns[ds.Tables[0].Columns.IndexOf("Книги.Название")].ColumnName = "Название";
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "books";
            tabControl1.SelectedIndex = 1;
        }

        private void SearchBooks_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void выдатьToolStripMenuItem_Click(object sender, EventArgs e)
        {

            using (GiveBooks GB = new GiveBooks())
            {
                GB.ids = new List<int>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    GB.ids.Add(Convert.ToInt32(row.Cells[0].Value));
                }
                if (GB.ShowDialog() == DialogResult.OK)
                {
                    button1_Click(sender, e);
                }
            }
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (AddEdBook AEB = new AddEdBook())
            {
                AEB.isEdit = true;
                AEB.rowsID = new List<string>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {

                    AEB.rowsID.Add(row.Cells[0].Value.ToString());
                }

                if (AEB.ShowDialog() == DialogResult.OK)
                {
                    button1_Click(sender, e);
                }
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var sCount = dataGridView1.SelectedRows.Count;
            if (MessageBox.Show($"Вы уверены что хотите удалить {sCount} записей?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string query = $"DELETE FROM [Книги] WHERE [Код] LIKE '{row.Cells[0].Value}'";
                    OleDbCommand command = new OleDbCommand(query, MF.Conn);
                    command.ExecuteNonQuery();
                }
                var b = treeView1.SelectedNode;
                treeView1.SelectedNode = null;
                treeView1.SelectedNode = b;
                MessageBox.Show($"Было удаленно {sCount} записей.", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            button1_Click(sender, e);
        }

        private void показатьИстриюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (History h = new History())
            {
                h.isSearch = true;
                h.sId = new List<string>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    h.sId.Add(row.Cells[0].Value.ToString());
                }
                if (h.ShowDialog() == DialogResult.OK)
                {
                    button1_Click(sender, e);

                }
            }
        }
    }
}