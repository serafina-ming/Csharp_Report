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
        /// <summary>
        /// 查詢是否存在該帳號，若存在則回傳使用者帳號名稱
        /// </summary>
        /// <param name="_account"></param>
        /// <param name="_password"></param>
        /// <returns></returns>
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

        //判斷有沒有登入成功
        string login = "N";
        /// <summary>
        /// 登入按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //判斷欄位有沒有輸入完全
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                Load_DB();
                string name = Show_DB(textBox1.Text, textBox2.Text);
                //如果有回傳的使用者名稱，就可以登入
                if (name != "")
                {
                    MessageBox.Show("歡迎"+ name + "登入");
                    login = "Y";
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
            //關閉此視窗則關閉系統
            //login變數判斷有沒有登入
            if (login == "N") 
            { 
                System.Environment.Exit(0);
            }
            
        }

        /// <summary>
        /// 密碼欄按下enter可以傳送登入驗證結果
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
