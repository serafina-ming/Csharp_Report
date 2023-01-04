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
using System.Windows.Input;
using System.Xml.Linq;
using static CsharpReport.Form1;

namespace CsharpReport
{
    public partial class registerForm : Form
    {
        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        public registerForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 新增會員資料
        /// </summary>
        /// <param name="name"></param>
        /// <param name="account"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        private bool CreateMember(string name, string account, string password)
        {
            if (!CheckAccountExist(account))
            {
                var command = DBConfig.sqlite_connect.CreateCommand();
                command.CommandText = @"INSERT INTO member (member_name, account, password)
                                            VALUES (@name, @account, @password);";
                command.Parameters.AddWithValue("@name", name);
                command.Parameters.AddWithValue("@account", account);
                command.Parameters.AddWithValue("@password", password);
                var result = command.ExecuteNonQuery();

                if (result != 0)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 確認帳號是否存在，存在回傳true
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        private bool CheckAccountExist(string account)
        {
            var command = DBConfig.sqlite_connect.CreateCommand();
            command.CommandText = @"SELECT member_id
                            FROM member
                            WHERE account = @account";
            command.Parameters.AddWithValue("@account", account);
            using (var result = command.ExecuteReader())
            {
                while (result.Read())
                {
                    int memberID = result.GetInt32(0);
                    if (memberID != 0)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text !="")
            {
                Load_DB();
                if (CreateMember(textBox1.Text, textBox2.Text, textBox3.Text))
                {
                    MessageBox.Show("註冊成功，請登入");
                    this.Close();
                }
                else
                {
                    MessageBox.Show("帳號已存在");
                }
            }
            else
            {
                MessageBox.Show("輸入不完全");
            }
        }

        /// <summary>
        /// 輸入框啟用enter輸入功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                button1_Click(sender, e);
            }
        }
    }
}
