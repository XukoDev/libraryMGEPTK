using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Resources;
using System.Collections;

namespace Библиотека_МГЭПТК
{
    public partial class GiveBooks : Form
    {
        MainForm MF = new MainForm();
        public List<int> ids { get; set; }
        public GiveBooks()
        {
            InitializeComponent();
        }

        private void GiveBooks_Load(object sender, EventArgs e)
        {
            string query = "SELECT Группы.Номер\r\nFROM Группы\r\nORDER BY Группы.Номер;";
            OleDbCommand con = new OleDbCommand(query, MF.Conn);
            OleDbDataReader reader = con.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add(reader[0]);
            }
            reader.Close();
            query = "SELECT Книги.Автор\r\nFROM Книги\r\nORDER BY Книги.Автор;\r\n";
            con = new OleDbCommand(query, MF.Conn);
            reader = con.ExecuteReader();
            while (reader.Read())
            {
                comboBox3.Items.Add(reader[0]);
            }
            iconUpdate();
            addBooks();
        }
        private void addBooks()
        {
            if (ids != null)
            {
                foreach (int i in ids)
                {
                    string query = $"SELECT Книги.Код, Книги.Название, Книги.Автор, Книги.Год, Книги.[Кол-во (на складе)]\r\nFROM Книги\r\nWHERE (((Книги.Код)={i}));\r\n";
                    OleDbCommand con = new OleDbCommand(query, MF.Conn);
                    OleDbDataReader reader = con.ExecuteReader();
                    while (reader.Read())
                    {
                        bool cont = false;
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value.ToString() == reader[0].ToString()) { cont = true; break; }
                        }
                        if (!cont)
                        {
                            dataGridView1.Rows.Add(reader[0], reader[1], reader[2], reader[3], reader[4]);
                        }
                        else { MessageBox.Show($"Книга с кодом \"{reader[0]}\" уже есть в списке!","Невозможно добавить книгу",MessageBoxButtons.OK,MessageBoxIcon.Error); }
                    }
                    reader.Close();
                }
            }
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox2.Text = "";

            iconUpdate();
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            string query = $"SELECT Ученики.ФИО\r\nFROM Группы INNER JOIN Ученики ON Группы.Код = Ученики.[Код группы]\r\nWHERE (((Группы.Номер)=\"{comboBox1.Text.Trim()}\"));\r\n";
            OleDbCommand con = new OleDbCommand(query, MF.Conn);
            OleDbDataReader reader = con.ExecuteReader();
            while (reader.Read())
            {
                comboBox2.Items.Add(reader[0]);
            }
            reader.Close();
        }
        private void iconUpdate()
        {
            button3.Enabled = false;

            if (comboBox2.Text.Trim() == "")
            {
                pictureBox1.Image = Properties.Resources.error;
                toolTip1.SetToolTip(pictureBox1, "Ученик не выбран");
            }
            else if (!comboBox2.Items.Contains(comboBox2.Text.Trim()))
            {
                pictureBox1.Image = Properties.Resources.caution;
                toolTip1.SetToolTip(pictureBox1, "Ученик не найден");
            }
            else
            {
                pictureBox1.Image = Properties.Resources.success;
                toolTip1.SetToolTip(pictureBox1, "Ошибок не обнаружено");
                button3.Enabled = true;
            }

            button1.Enabled = false;
            if (comboBox4.Text.Trim() == "")
            {
                pictureBox2.Image = Properties.Resources.error;
                toolTip1.SetToolTip(pictureBox2, "Книга не выбрана");
            }
            else if (!comboBox4.Items.Contains(comboBox4.Text.Trim()))
            {
                pictureBox2.Image = Properties.Resources.caution;
                toolTip1.SetToolTip(pictureBox2, "Книга не найдена");
            }
            else
            {
                pictureBox2.Image = Properties.Resources.success;
                toolTip1.SetToolTip(pictureBox2, "Ошибок не обнаружено");
                button1.Enabled = true;
            }
        }

        private void comboBox2_TextUpdate(object sender, EventArgs e)
        {
            iconUpdate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                if (MessageBox.Show($"Удалить {dataGridView1.SelectedRows.Count} записей из списка?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        dataGridView1.Rows.Remove(row);
                    }
                    MessageBox.Show("Выбранные записи были удалены из списка.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else { MessageBox.Show("Список книг пуст!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> arow = new List<DataGridViewRow>();
            if (dataGridView1.Rows.Count > 0)
            {
                if (MessageBox.Show($"Выдать {dataGridView1.Rows.Count} книг?", "Выдать?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    bool suc = false;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        suc = false;
                        if (Convert.ToInt32(row.Cells[4].Value) < Convert.ToInt32(row.Cells[5].Value))
                        {
                            MessageBox.Show($"Книг с кодом \"{row.Cells[0].Value}\" недостаточно!", "Ошибка выдачи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }
                        else if (Convert.ToInt32(row.Cells[4].Value) == 0)
                        {
                            MessageBox.Show($"Книг с кодом \"{row.Cells[0].Value}\" нет в наличии!", "Ошибка выдачи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }
                        else if (row.Cells[5].Value == null)
                        {
                            MessageBox.Show($"Не указано количество для выдачи в книге с кодом \"{row.Cells[0].Value}\"!", "Ошибка выдачи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }
                        string query = $"INSERT INTO История ( [Код ученика], [Код книги], [Кол-во книг], Выдано )\r\nSELECT Ученики.Код, {row.Cells[0].Value}, {row.Cells[5].Value}, '{DateTime.Now}'\r\nFROM Книги, Группы INNER JOIN Ученики ON Группы.Код = Ученики.[Код группы]\r\nGROUP BY Ученики.Код, {row.Cells[0].Value}, {row.Cells[5].Value}, '{DateTime.Now}', Ученики.ФИО, Группы.Номер\r\nHAVING (((Ученики.ФИО)=\"{comboBox2.Text}\") AND ((Группы.Номер)=\"{comboBox1.Text}\"));\r\n";
                        OleDbCommand con = new OleDbCommand(query, MF.Conn);
                        con.ExecuteNonQuery();
                        query = $"UPDATE Книги SET Книги.[Кол-во (на складе)] = [Книги].[Кол-во (на складе)]-{row.Cells[5].Value}\r\nWHERE (((Книги.Код)={row.Cells[0].Value}));\r\n";
                        con = new OleDbCommand(query, MF.Conn);
                        con.ExecuteNonQuery();
                        arow.Add(row);
                        suc = true;
                    }
                    if (suc)
                    {
                        MessageBox.Show("Книги были выданы.", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dataGridView1.Rows.Clear();
                    }
                    else
                    {
                        if (arow.Count > 0)
                        {
                            foreach (var row in arow)
                            {
                                dataGridView1.Rows.Remove(row);
                            }
                            arow.Clear();
                        }
                        MessageBox.Show("Выдача книг была прервана.", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else { MessageBox.Show("Список книг пуст!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void comboBox3_TextUpdate(object sender, EventArgs e)
        {
            comboBox4.Items.Clear();
            comboBox4.Text = "";
            iconUpdate();
            comboBox4.Items.Clear();
            comboBox4.Text = "";
            string query = $"SELECT Книги.Название, Книги.Год\r\nFROM Книги\r\nWHERE (((Книги.Автор)=\"{comboBox3.Text.Trim()}\"))\r\nORDER BY Книги.Название;\r\n";
            OleDbCommand con = new OleDbCommand(query, MF.Conn);
            OleDbDataReader reader = con.ExecuteReader();
            while (reader.Read())
            {
                comboBox4.Items.Add($"{reader[0]} ({reader[1]} г.)");
            }
            reader.Close();
        }


        private void comboBox4_TextUpdate(object sender, EventArgs e)
        {
            iconUpdate();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string year = comboBox4.Text.Substring(comboBox4.Text.Length - 8);
            string name = comboBox4.Text.Remove(comboBox4.Text.Length - 9);
            string query = $"SELECT Книги.Код, Книги.Год\r\nFROM Книги\r\nWHERE (((Книги.Название)=\"{name.Trim()}\") AND ((Книги.Автор)=\"{comboBox3.Text.Trim()}\") AND ((Книги.Год)={year.Replace("г.)", "").Trim()}));\r\n";
            OleDbCommand con = new OleDbCommand(query, MF.Conn);
            var id = con.ExecuteScalar();
            ids = new List<int>();
            ids.Add(Convert.ToInt32(id));
            addBooks();
            comboBox3.Text = "";
            comboBox4.Text = "";
            iconUpdate();
        }
    }
}