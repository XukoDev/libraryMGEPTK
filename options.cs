using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Библиотека_МГЭПТК.Properties;

namespace Библиотека_МГЭПТК
{
    public partial class options : Form
    {
        private bool val1, val2;
        public options()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (button1.Text == "Показать")
            {
                if (settings.Default.isPass)
                {
                    using (login log = new login())
                    {
                        if (log.ShowDialog() == DialogResult.OK)
                        {
                            textBox1.PasswordChar = '\0';
                            button1.Text = "Скрыть";
                        }
                    }
                }
                else
                {
                    textBox1.PasswordChar = '\0';
                    button1.Text = "Скрыть";
                }
            }
            else
            {
                textBox1.PasswordChar = '*';
                button1.Text = "Показать";
            }
        }

        private void options_Load(object sender, EventArgs e)
        {
            textBox1.Text = settings.Default.pass;
            checkBox1.Checked = settings.Default.isPass;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (settings.Default.isPass)
            {
                using (login log = new login())
                {
                    if (log.ShowDialog() == DialogResult.OK)
                    {
                        settings.Default.pass = textBox1.Text;
                        settings.Default.isPass = checkBox1.Checked;
                        settings.Default.Save();
                        MessageBox.Show("Настройки были сохранены","Успех",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        button2.Enabled = false;
                    }
                }
            }
            else
            {
                settings.Default.pass = textBox1.Text;
                settings.Default.isPass = checkBox1.Checked;
                settings.Default.Save();
                MessageBox.Show("Настройки были сохранены", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button2.Enabled = false;
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != settings.Default.pass)
            {
                val1 = false;
            }
            else { val1 = true; }
                validCheck();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked != settings.Default.isPass)
            {
                val2 = false;
            }
            else { val2 = true; }
            validCheck();
        }
        private void validCheck()
        {
            if (val1 && val2)
            {
                button2.Enabled = false;
            }
            else { button2.Enabled = true; }
        }
    }
}