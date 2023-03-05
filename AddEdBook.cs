using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Net.WebRequestMethods;

namespace Библиотека_МГЭПТК
{
    public partial class AddEdBook : Form
    {
        MainForm MF = new MainForm();
        ComboBox cblvl = new ComboBox();
        ComboBox cbSpec = new ComboBox();
        ComboBox cbDisc = new ComboBox();
        public List<String> rowsID { get; set; }
        public bool isEdit { get; set; }
        public AddEdBook()
        {
            InitializeComponent();
        }

        private void AddEdBook_Load(object sender, EventArgs e)
        {

            if (!isEdit)
            {
                comboAdd();
                dataGridView1.AllowUserToAddRows = true;
                dataGridView1.Columns[3].ReadOnly = false;
                Column9.Visible = false;
                button1.Text = "Добавить";
                this.Text = "Добавить книги";
            }
            else
            {
                Column9.Visible = true;
                dataGridView1.AllowUserToAddRows = false;
                dataGridView1.Columns[4].ReadOnly = true;
                foreach (string s in rowsID)
                {
                    string query = $"SELECT Книги.Название, Книги.Год, Книги.Автор, Книги.[Кол-во (на складе)], Книги.[Кол-во (всего)], Специальности.Уровень, Специальности.Название, Дисциплины.Название, Книги.Код\r\nFROM Специальности INNER JOIN (Дисциплины INNER JOIN Книги ON Дисциплины.Код = Книги.Дисциплина) ON Специальности.Код = Дисциплины.Специальность\r\nWHERE (((Книги.Код)={s}));\r\n";
                    OleDbCommand con = new OleDbCommand(query, MF.Conn);
                    OleDbDataReader reader = con.ExecuteReader();
                    while (reader.Read())
                    {
                        dataGridView1.Rows.Add(reader[8], reader[0], reader[1], reader[2], reader[3], reader[4], reader[5], reader[6], reader[7]);
                    }
                    reader.Close();
                }
                comboAdd();
                button1.Text = "Изменить";
                this.Text = "Изменить книги";
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(ColumnNum_KeyPress);
            if (dataGridView1.CurrentCell.ColumnIndex == 2 || dataGridView1.CurrentCell.ColumnIndex == 4 || dataGridView1.CurrentCell.ColumnIndex == 5)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(ColumnNum_KeyPress);
                }
            }
        }

        private void ColumnNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        private void comboAdd()
        {
            string query;
            OleDbCommand con;
            OleDbDataReader reader;
            query = "SELECT Специальности.Уровень\r\nFROM Специальности;\r\n";
            con = new OleDbCommand(query, MF.Conn);
            reader = con.ExecuteReader();
            while (reader.Read())
            {
                if (!cblvl.Items.Contains(reader[0])) cblvl.Items.Add(reader[0]);
            }
            reader.Close();

            if (dataGridView1.SelectedRows.Count <= 0) return;
            else
            {
                if (dataGridView1.SelectedRows[0].Cells[Column6.Index].Value == null)
                {
                    return;
                }
                else
                {
                    query = $"SELECT Специальности.Название\r\nFROM Специальности\r\nWHERE (((Специальности.Уровень)=\"{dataGridView1.SelectedRows[0].Cells[Column6.Index].Value}\"));\r\n";
                    con = new OleDbCommand(query, MF.Conn);
                    reader = con.ExecuteReader();
                    while (reader.Read())
                    {
                        if (!cbSpec.Items.Contains(reader[0])) cbSpec.Items.Add(reader[0]);
                    }
                    reader.Close();

                    if (dataGridView1.SelectedRows[0].Cells[Column7.Index].Value == null) return;
                    else
                    {

                        query = $"SELECT Дисциплины.Название\r\nFROM Специальности INNER JOIN Дисциплины ON Специальности.Код = Дисциплины.Специальность\r\nWHERE (((Специальности.Уровень)=\"{dataGridView1.SelectedRows[0].Cells[Column6.Index].Value}\") AND ((Специальности.Название)=\"{dataGridView1.SelectedRows[0].Cells[Column7.Index].Value}\"));\r\n";
                        con = new OleDbCommand(query, MF.Conn);
                        reader = con.ExecuteReader();
                        while (reader.Read())
                        {
                            if (!cbDisc.Items.Contains(reader[0])) cbDisc.Items.Add(reader[0]);
                        }

                    }
                }
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cblvl.Visible = false;
            cblvl.Text = "";
            cblvl.Items.Clear();
            cbSpec.Visible = false;
            cbSpec.Text = "";
            cbSpec.Items.Clear();
            cbDisc.Visible = false;
            cbDisc.Text = "";
            cbDisc.Items.Clear();
            comboAdd();
            try
            {
                if (e.ColumnIndex == Column6.Index)
                {
                    if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                    {
                        cblvl.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    }
                    cblvl.Visible = true;
                    cblvl.DropDownStyle = ComboBoxStyle.DropDownList;
                    var rect = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    cblvl.Size = new Size(rect.Width, rect.Height);
                    cblvl.Location = new Point(rect.X, rect.Y);
                    cblvl.DropDownClosed += new EventHandler(cblvl_ddc);
                    cblvl.SelectedIndexChanged += new EventHandler(cblvl_sic);
                    dataGridView1.Controls.Add(cblvl);
                }
                else if (e.ColumnIndex == Column7.Index)
                {
                    if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                    {
                        cbSpec.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    }
                    cbSpec.Visible = true;
                    cbSpec.DropDownStyle = ComboBoxStyle.DropDownList;
                    var rect = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    cbSpec.Size = new Size(rect.Width, rect.Height);
                    cbSpec.Location = new Point(rect.X, rect.Y);
                    cbSpec.DropDownClosed += new EventHandler(cbSpec_ddc);
                    cbSpec.SelectedIndexChanged += new EventHandler(cbSpec_sic);
                    dataGridView1.Controls.Add(cbSpec);

                }
                else if (e.ColumnIndex == Column8.Index)
                {

                    if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                    {
                        cbDisc.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    }
                    cbDisc.Visible = true;
                    cbDisc.DropDownStyle = ComboBoxStyle.DropDownList;
                    var rect = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    cbDisc.Size = new Size(rect.Width, rect.Height);
                    cbDisc.Location = new Point(rect.X, rect.Y);
                    cbDisc.DropDownClosed += new EventHandler(cbDisc_ddc);
                    cbDisc.SelectedIndexChanged += new EventHandler(cbDisc_sic);
                    dataGridView1.Controls.Add(cbDisc);
                }
            }
            catch { }
        }
        private void cblvl_ddc(object sender, EventArgs e)
        {
            if (cblvl.SelectedItem != null)
                dataGridView1.CurrentCell.Value = cblvl.SelectedItem.ToString();
        }
        void cblvl_sic(object sender, EventArgs e)
        {
            cblvl.Visible = false;
        }
        private void cbSpec_ddc(object sender, EventArgs e)
        {
            if (cbSpec.SelectedItem != null)
                dataGridView1.CurrentCell.Value = cbSpec.SelectedItem.ToString();
        }
        void cbSpec_sic(object sender, EventArgs e)
        {
            cbSpec.Visible = false;
        }
        private void cbDisc_ddc(object sender, EventArgs e)
        {
            if (cbDisc.SelectedItem != null)
                dataGridView1.CurrentCell.Value = cbDisc.SelectedItem.ToString();
        }
        void cbDisc_sic(object sender, EventArgs e)
        {
            cbDisc.Visible = false;
        }
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> delRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string query = $"SELECT Дисциплины.Код\r\nFROM Специальности INNER JOIN Дисциплины ON Специальности.Код = Дисциплины.Специальность\r\nWHERE (((Специальности.Уровень)=\"{row.Cells[6].Value}\") AND ((Специальности.Название)=\"{row.Cells[7].Value}\") AND ((Дисциплины.Название)=\"{row.Cells[8].Value}\"));\r\n";
                    OleDbCommand con = new OleDbCommand(query, MF.Conn);
                    var id = con.ExecuteScalar();
                    if (!isEdit)
                    {
                        query = $"INSERT INTO Книги ( Название, Год, Автор, [Кол-во (на складе)], [Кол-во (всего)], Дисциплина )\r\n" +
                            $"VALUES(\"{row.Cells[1].Value}\",{row.Cells[2].Value},\"{row.Cells[3].Value}\",{row.Cells[4].Value},{row.Cells[5].Value},{id});";
                    }
                    else
                    {
                        query = $"SELECT История.[Кол-во книг]\r\nFROM История\r\nWHERE (((История.[Код книги])={row.Cells[0].Value}) AND ((История.Возвращено) Is Null));\r\n";
                        con = new OleDbCommand(query, MF.Conn);
                        OleDbDataReader reader = con.ExecuteReader();
                        int sklad = 0;
                        while (reader.Read())
                        {
                            sklad += Convert.ToInt32(reader[0]);
                        }
                        reader.Close();
                        query = $"UPDATE Книги SET Книги.Название = \"{row.Cells[1].Value}\", " +
                            $"Книги.Год = {row.Cells[2].Value}, " +
                            $"Книги.Автор = \"{row.Cells[3].Value}\"," +
                            $" Книги.[Кол-во (на складе)] = {Convert.ToInt32(row.Cells[5].Value)-sklad}," +
                            $" Книги.[Кол-во (всего)] = {row.Cells[5].Value}, " +
                            $"Книги.Дисциплина = {id}\r\n" +
                            $"WHERE (((Книги.Код)={row.Cells[0].Value}));\r\n";
                    }
                    con = new OleDbCommand(query, MF.Conn);
                    con.ExecuteNonQuery();
                    delRows.Add(row);
                }
                catch {
                    MessageBox.Show("Данные введены не коректно!\n" +
                        "Пожалуйста проверьте правильность введенных данных.","Ошибка",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    break; }
            }
            foreach (DataGridViewRow row in delRows)
            {
                try
                {
                    dataGridView1.Rows.Remove(row);
                }
                catch { }
            }
            if (dataGridView1.Rows.Count <= 1)
            {
                this.DialogResult = DialogResult.OK;
            }
        }
    }
}