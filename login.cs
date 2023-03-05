using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Библиотека_МГЭПТК
{
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Показать")
            {
                textBox1.PasswordChar = '\0';
                button1.Text = "Скрыть";
            }
            else
            {
                textBox1.PasswordChar = '*';
                button1.Text = "Показать";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == settings.Default.pass)
            {
                DialogResult= DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Повторите попытку","Неверный пароль",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            settings.Default.pass = "0000";
            settings.Default.Save();
            MessageBox.Show("Пароль был сброшен на пароль по умолчанию","Сброс пароля",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }
    }
}
