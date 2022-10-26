using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CsharpReport.Form1;

namespace CsharpReport
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        private string Show_DB(string _account,string _password)
        {
            var command = DBConfig.sqlite_connect.CreateCommand();
            command.CommandText = @"SELECT member_name
                            FROM member
                            WHERE account = @account AND password = @password";
            var name = "";
            command.Parameters.AddWithValue("@account", _account);
            command.Parameters.AddWithValue("@password", _password);
            using (var result = command.ExecuteReader())
            {
                while (result.Read())
                {
                    name = result.GetString(0);
                }
            }

            return name;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Load_DB();
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                Load_DB();
                string result = Show_DB(textBox1.Text, textBox2.Text);
                if (result != "")
                {
                    MessageBox.Show("歡迎"+ result + "登入");
                    this.Close();
                }
                else
                {
                    MessageBox.Show("登入失敗");
                }
            }
            else
            {
                MessageBox.Show("輸入不完全");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
    }
}
