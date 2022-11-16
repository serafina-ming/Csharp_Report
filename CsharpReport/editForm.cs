using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CsharpReport.Form1;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using ComboBox = System.Windows.Forms.ComboBox;

namespace CsharpReport
{
    public partial class editForm : Form
    {
        public editForm()
        {
            InitializeComponent();
            GetBookData();
        }

        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open
        }

        /// <summary>
        /// 下拉式選單的內容
        /// </summary>
        private void GetBookData()
        {
            
            var command = DBConfig.sqlite_connect.CreateCommand();
            string sql = @"SELECT member_id, member_name
                            FROM member";
            string status = "";

            command.CommandText = sql;
            DBConfig.sqlite_datareader = command.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    string member_id = Convert.ToString(DBConfig.sqlite_datareader["member_id"]);
                    string member_name = Convert.ToString(DBConfig.sqlite_datareader["member_name"]);

                    comboBox1.Items.Add(new ComboBoxItem(member_id, member_name));
                }
                //DBConfig.sqlite_datareader.Close();
            }
        }
        /// <summary>
        /// 設定下拉式選單的Value與Text
        /// </summary>
        public class ComboBoxItem
        {
            public string Value { get; set; }
            public string Text { get; set; }
            public ComboBoxItem(string value, string text)
            {
                Value = value;
                Text = text;
            }
            public override string ToString()
            {
                return Text;
            }
        }
        /// <summary>
        /// 取得下拉式選單的Value
        /// </summary>
        public class ComboUtil
        {
            /// <summary>
            /// 取得下拉項目的值
            /// </summary>
            /// <param name="cbo">物件</param>
            /// <returns></returns>
            public static ComboBoxItem GetItem(ComboBox cbo)
            {
                ComboBoxItem item = new ComboBoxItem("", "");
                if (cbo.SelectedIndex > -1)
                {
                    item = cbo.Items[cbo.SelectedIndex] as ComboBoxItem;
                }
                return item;
            }

        }

        /// <summary>
        /// 設定目前的書籍資料
        /// </summary>
        public object[] setValue
        {
            set
            {
                label7.Text = value[0].ToString();
                textBox1.Text = value[1].ToString();
                textBox2.Text = value[2].ToString();
                textBox3.Text = value[3].ToString();
                comboBox1.SelectedIndex = (int)value[4]-1;
                textBox4.Text = value[6].ToString();

                if (value[5].ToString() == "可借出")
                {
                    radioButton1.Checked = true;
                }
                else
                {
                    radioButton2.Checked = true;
                }
            }
        }

        /// <summary>
        /// 更新按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //取得下拉式選單的值
            //ComboUtil.GetItem(comboBox1).Value

            var command = DBConfig.sqlite_connect.CreateCommand();
            command.CommandText = @"UPDATE book_data
                        SET book_name = @book_name,writer = @writer,publish = @publish,
                            category = @category, status = @status,book_keeper = @book_keeper
                        WHERE book_id = @book_id";

            command.Parameters.AddWithValue("@book_id", label7.Text);
            command.Parameters.AddWithValue("@book_name", textBox1.Text);
            command.Parameters.AddWithValue("@writer", textBox2.Text);
            command.Parameters.AddWithValue("@publish", textBox3.Text);
            command.Parameters.AddWithValue("@category", comboBox1.SelectedIndex + 1);
            command.Parameters.AddWithValue("@status", radioButton1.Text);

            //判斷借閱狀態，有切換到且以借出有借閱人才可以執行更新動作
            var status = false;
            if (radioButton1.Checked == true)
            {
                command.Parameters.AddWithValue("@status", radioButton1.Text);
                command.Parameters.AddWithValue("@book_keeper", "");
                status = true;
            }
            else if (textBox4.Text != "")
            {
                command.Parameters.AddWithValue("@status", radioButton2.Text);
                command.Parameters.AddWithValue("@book_keeper", textBox4.Text);
                status = true;
            }
            else
            {
                MessageBox.Show("請填寫借閱人");
            }
            if (status == true)
            {
                try
                {
                    command.ExecuteNonQuery();
                    MessageBox.Show("更新成功");
                }
                catch (SqlException ex)
                {
                    //MessageBox.Show(ex.Message);
                    MessageBox.Show("系統出現沒辦法解決的錯誤，請放棄並離開");
                }
            }
        }

        /// <summary>
        /// 可借出按鈕偵測
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //如果切換到可借出就將借閱人內容清空
            if (radioButton1.Checked == true)
            {
                textBox4.Text = "";
            }
            
        }

    }
}
