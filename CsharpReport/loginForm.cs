using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CsharpReport.Form1;

namespace CsharpReport
{
    public partial class loginForm : Form
    {
        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        /// <summary>
        /// 驗證使用者
        /// </summary>
        /// <param name="account">帳號</param>
        /// <param name="password">密碼</param>
        /// <returns>使用者姓名</returns>
        private string Check_User(string account, string password)
        {
            var command = DBConfig.sqlite_connect.CreateCommand();
            command.CommandText = @"SELECT member_name
                            FROM member
                            WHERE account = @account AND password = @password";
            var name = "";

            command.Parameters.AddWithValue("@account", account);
            command.Parameters.AddWithValue("@password", password);
            using (var result = command.ExecuteReader())
            {
                while (result.Read())
                {
                    name = result.GetString(0);
                }
            }

            return name;
        }

        public loginForm()
        {
            InitializeComponent();
        }

        //紀錄是否登入
        bool loginCheck = false;
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                Load_DB();
                string result = Check_User(textBox1.Text, textBox2.Text);
                if (result != "")
                {
                    MessageBox.Show("歡迎 " + result + " 登入");
                    loginCheck = true;
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
            registerForm registerForm;
            registerForm = new registerForm();
            registerForm.ShowDialog();
        }

        /// <summary>
        /// 關閉登入視窗時，若未登入則直接關閉整個程式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void loginForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!loginCheck)
            {
                System.Environment.Exit(0);
            }
        }

        /// <summary>
        /// 打完密碼按enter可直接登入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                button1_Click(sender, e);
            }
        }
    }
}
