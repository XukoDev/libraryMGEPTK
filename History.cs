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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace Библиотека_МГЭПТК
{
    public partial class History : Form
    {
        private MainForm MF = new MainForm();
        public bool isSearch { get; set; }
        public List<string> sId { get; set; }
        public History()
        {
            InitializeComponent();
        }

        private void History_Load(object sender, EventArgs e)
        {
            TabsUpdate();
        }
        private void TabsUpdate()
        {
            DataSet ds = new DataSet();
            if (!isSearch)
            {
                string query = "SELECT История.Код, Ученики.ФИО, Книги.Название, История.[Кол-во книг], История.Выдано\r\nFROM Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]\r\nWHERE (((История.Возвращено) Is Null))\r\nORDER BY История.Код;\r\n";
                OleDbDataAdapter oleDb = new OleDbDataAdapter(query, MF.Conn);
                oleDb.Fill(ds, "activeB");
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "activeB";
                query = "SELECT История.Код, Ученики.ФИО, Книги.Название, История.[Кол-во книг], История.Выдано, История.Возвращено\r\nFROM Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]\r\nWHERE (((История.Возвращено) Is Not Null))\r\nORDER BY История.Код;\r\n";
                oleDb = new OleDbDataAdapter(query, MF.Conn);
                oleDb.Fill(ds, "closeB");
                dataGridView2.DataSource = ds;
                dataGridView2.DataMember = "closeB";
            }
            else
            {
                foreach (string s in sId)
                {
                    string query = $"SELECT История.Код, Ученики.ФИО, Книги.Название, История.[Кол-во книг], История.Выдано\r\nFROM Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]\r\nWHERE (((История.Возвращено) Is Null) AND ((История.[Код книги])={s}))\r\nORDER BY История.Код;\r\n";
                    OleDbDataAdapter oleDb = new OleDbDataAdapter(query, MF.Conn);
                    oleDb.Fill(ds, "activeB");
                    dataGridView1.DataSource = ds;
                    dataGridView1.DataMember = "activeB";
                    query = $"SELECT История.Код, Ученики.ФИО, Книги.Название, История.[Кол-во книг], История.Выдано\r\nFROM Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]\r\nWHERE (((История.Возвращено) Is Not Null) AND ((История.[Код книги])={s}))\r\nORDER BY История.Код;\r\n";
                    oleDb = new OleDbDataAdapter(query, MF.Conn);
                    oleDb.Fill(ds, "closeB");
                    dataGridView2.DataSource = ds;
                    dataGridView2.DataMember = "closeB";
                }
            }
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0) { button1.Enabled = false; button2.Enabled = false; }
            else if (dataGridView1.SelectedRows.Count == 1) { button1.Enabled = true; button2.Enabled = true; }
            else { button1.Enabled = true; button2.Enabled = false; }

            richTextBox1.Text = "";
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string query = $"SELECT История.Код, Ученики.Код, Ученики.ФИО, Ученики.[Код группы], Группы.Куратор, Книги.Код, Книги.Название, Книги.Автор, Книги.Год, Дисциплины.Название, История.[Кол-во книг], История.Выдано\r\nFROM Дисциплины INNER JOIN (Группы INNER JOIN (Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]) ON Группы.Код = Ученики.[Код группы]) ON Дисциплины.Код = Книги.Дисциплина\r\nWHERE (((История.Код)={row.Cells[0].Value}));\r\n";
                OleDbCommand con = new OleDbCommand(query, MF.Conn);
                OleDbDataReader reader = con.ExecuteReader();
                while (reader.Read())
                {
                    richTextBox1.AppendText($"Код: {reader[0]}\r\n" +
                        $"Ученик: " +
                        $"\r\n\tКод: {reader[1]}" +
                        $"\r\n\tФИО: {reader[2]}" +
                        $"\r\n\tГруппа: {reader[3]}" +
                        $"\r\n\tКуратор: {reader[4]}" +
                        $"\nКнига: " +
                        $"\r\n\tКод: {reader[5]} " +
                        $"\r\n\tНазвание: {reader[6]} " +
                        $"\r\n\tАвтор: {reader[7]}" +
                        $"\r\n\tГод: {reader[8]} " +
                        $"\r\n\tДисциплина: {reader[9]}" +
                        $"\r\nВзято: {reader[10]} шт." +
                        $"\r\nВыдано: {reader[11]}" +
                        $"\r\nВозвращено: ----");
                }
                reader.Close();
                richTextBox1.AppendText("\n--------------------------------------------\n");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены что хотите полностью закрыть выбранные долги?", "Закрыть долг(и)", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    string query = $"UPDATE Книги INNER JOIN История ON Книги.Код = История.[Код книги] SET История.Возвращено = \"{DateTime.Now}\", Книги.[Кол-во (на складе)] = [Книги].[Кол-во (на складе)]+[История].[Кол-во книг]\r\nWHERE (((История.Код)={row.Cells[0].Value}));\r\n";
                    OleDbCommand con = new OleDbCommand(query, MF.Conn);
                    con.ExecuteNonQuery();
                }
                TabsUpdate();
            }
        }

        private void History_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            richTextBox2.Text = "";
            foreach (DataGridViewRow row in dataGridView2.SelectedRows)
            {
                string query = $"SELECT История.Код, Ученики.Код, Ученики.ФИО, Ученики.[Код группы], Группы.Куратор, Книги.Код, Книги.Название, Книги.Автор, Книги.Год, Дисциплины.Название, История.[Кол-во книг], История.Выдано, История.Возвращено\r\nFROM Дисциплины INNER JOIN (Группы INNER JOIN (Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]) ON Группы.Код = Ученики.[Код группы]) ON Дисциплины.Код = Книги.Дисциплина\r\nWHERE (((История.Код)={row.Cells[0].Value}));\r\n";
                OleDbCommand con = new OleDbCommand(query, MF.Conn);
                OleDbDataReader reader = con.ExecuteReader();
                while (reader.Read())
                {
                    richTextBox2.AppendText($"Код: {reader[0]}\r\n" +
                        $"Ученик: " +
                        $"\r\n\tКод: {reader[1]}" +
                        $"\r\n\tФИО: {reader[2]}" +
                        $"\r\n\tГруппа: {reader[3]}" +
                        $"\r\n\tКуратор: {reader[4]}" +
                        $"\nКнига: " +
                        $"\r\n\tКод: {reader[5]} " +
                        $"\r\n\tНазвание: {reader[6]} " +
                        $"\r\n\tАвтор: {reader[7]}" +
                        $"\r\n\tГод: {reader[8]} " +
                        $"\r\n\tДисциплина: {reader[9]}" +
                        $"\r\nВзято: {reader[10]} шт." +
                        $"\r\nВыдано: {reader[11]}" +
                        $"\r\nВозвращено: {reader[12]}");
                }
                reader.Close();
                richTextBox2.AppendText("\n--------------------------------------------\n");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (HAsk Ha = new HAsk())
            {
                if (Ha.ShowDialog() == DialogResult.Yes)
                {
                    int co = Convert.ToInt32(Ha.co);
                    string query;
                    if (co >= Convert.ToInt32(dataGridView1.SelectedCells[3].Value))
                    {
                        query = $"UPDATE Книги INNER JOIN История ON Книги.Код = История.[Код книги] SET История.Возвращено = \"{DateTime.Now}\", Книги.[Кол-во (на складе)] = [Книги].[Кол-во (на складе)]+[История].[Кол-во книг]\r\nWHERE (((История.Код)={dataGridView1.SelectedCells[0].Value}));\r\n";
                        OleDbCommand con = new OleDbCommand(query, MF.Conn);
                        con.ExecuteNonQuery();

                    }
                    else
                    {
                        query = $"UPDATE Книги INNER JOIN История ON Книги.Код = История.[Код книги] SET История.[Кол-во книг] = [История].[Кол-во книг]-{co}, Книги.[Кол-во (на складе)] = [Книги].[Кол-во (на складе)]+{co}\r\nWHERE (((История.Код)={dataGridView1.SelectedCells[0].Value}));\r\n";
                        OleDbCommand con = new OleDbCommand(query, MF.Conn);
                        con.ExecuteNonQuery();
                        query = $"INSERT INTO История ( [Код ученика], [Код книги], [Кол-во книг], Выдано, Возвращено )\r\nSELECT История.[Код ученика], История.[Код книги], \"{co}\", \"{dataGridView1.SelectedCells[4].Value}\",\"{DateTime.Now}\"\r\nFROM История\r\nWHERE (((История.Код)={dataGridView1.SelectedCells[0].Value}));\r\n";
                        con = new OleDbCommand(query, MF.Conn);
                        con.ExecuteNonQuery();
                    }
                    TabsUpdate();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel не установлен!");
                return;
            }
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlApp.DisplayAlerts = false;


            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ФИО";
            xlWorkSheet.Cells[1, 2] = "Книга";
            xlWorkSheet.Cells[1, 3] = "Автор";
            xlWorkSheet.Cells[1, 4] = "Кол-во";
            xlWorkSheet.Cells[1, 5] = "Выдано";
            Excel.Range r = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 5]];
            r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);



            string query = "SELECT Группы.Номер, Ученики.ФИО, Книги.Название, Книги.Год, Книги.Автор, История.[Кол-во книг], История.Выдано, Группы.Куратор\r\nFROM Группы INNER JOIN (Ученики INNER JOIN (Книги INNER JOIN История ON Книги.Код = История.[Код книги]) ON Ученики.Код = История.[Код ученика]) ON Группы.Код = Ученики.[Код группы]\r\nWHERE (((История.Возвращено) Is Null))\r\nORDER BY Группы.Номер, Ученики.ФИО;\r\n";
            OleDbCommand con = new OleDbCommand(query,MF.Conn);
            OleDbDataReader reader = con.ExecuteReader();
            int stroka = 2;
            string group = null;
            string fio = null;
            while(reader.Read())
            {
                if (group != reader[0].ToString()) 
                {
                    group = reader[0].ToString();
                    xlWorkSheet.Cells[stroka, 1] = $"{reader[0]} {reader[7]}";
                    r = xlWorkSheet.Range[xlWorkSheet.Cells[stroka, 1], xlWorkSheet.Cells[stroka, 5]];
                    r.Merge(Type.Missing);
                    r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    r.Interior.Color= System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                    stroka++;
                }
                xlWorkSheet.Cells[stroka, 2] = $"{reader[2]} ({reader[3]})";
                xlWorkSheet.Cells[stroka, 3] = reader[4];
                xlWorkSheet.Cells[stroka, 4] = reader[5];
                xlWorkSheet.Cells[stroka, 5] = reader[6];
                if (fio != reader[1].ToString())
                {
                    fio = reader[1].ToString();
                    xlWorkSheet.Cells[stroka, 1] = reader[1];
                    r = xlWorkSheet.Cells[stroka, 1];
                }
                stroka++;
                xlWorkSheet.Columns.AutoFit();
            }
            reader.Close();
            r = xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[stroka-1, 5]];
            r.Borders.LineStyle = XlLineStyle.xlContinuous;

            if (File.Exists($"{Environment.CurrentDirectory}\\Отчет {DateTime.Today.ToString("d")}.xls"))
            {
                try
                {
                    File.Delete($"{Environment.CurrentDirectory}\\Отчет {DateTime.Today.ToString("d")}.xls");
                }
                catch { }
            }
            try
            {
                xlWorkBook.SaveAs($"{Environment.CurrentDirectory}\\Отчет {DateTime.Today.ToString("d")}.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            }
            catch { }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Отчет расположен в папке программы", "Отчет успешно создан", MessageBoxButtons.OK);
            System.Diagnostics.Process.Start("explorer", Environment.CurrentDirectory);
           
        }
    }
}