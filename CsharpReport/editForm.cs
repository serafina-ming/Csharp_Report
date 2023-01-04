using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static CsharpReport.Form1;

namespace CsharpReport
{
    public partial class editForm : Form
    {
        int _bookId = 0;
        string _bookName = "";
        string _writer = "";
        string _publish = "";
        int _categoryId = 0;
        string _status = "";
        int _memberId = 0;
        public editForm()
        {
            InitializeComponent();
            GetMemberData();
        }

        /// <summary>
        /// 設定初始值(書籍編號)，從Form1傳入
        /// </summary>
        public object[] setValue
        {
            set
            {
                label7.Text = value[0].ToString();
            }
        }

        /// <summary>
        /// 依據書籍編號取得書籍資料
        /// </summary>
        public void GetThisBookData()
        {
            var command = DBConfig.sqlite_connect.CreateCommand();
            string sql = @"SELECT book_id, book_name, writer, publish,
                            category_id, status, member_id
                            FROM book_data
                            LEFT JOIN category_data
                            ON category = category_id
                            LEFT JOIN member
                            ON book_keeper = member_id
                            WHERE book_id = @book_id";

            command.CommandText = sql;
            command.Parameters.AddWithValue("@book_id", label7.Text);
            DBConfig.sqlite_datareader = command.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    _bookId = Convert.ToInt32(DBConfig.sqlite_datareader["book_id"]);
                    _bookName = Convert.ToString(DBConfig.sqlite_datareader["book_name"]);
                    _writer = Convert.ToString(DBConfig.sqlite_datareader["writer"]);
                    _publish = Convert.ToString(DBConfig.sqlite_datareader["publish"]);
                    _categoryId = Convert.ToInt32(DBConfig.sqlite_datareader["category_id"]);
                    _status = Convert.ToString(DBConfig.sqlite_datareader["status"]);
                    _memberId = 0;
                    if (Convert.ToString(DBConfig.sqlite_datareader["member_id"]) != "")
                    {
                        _memberId = Convert.ToInt32(DBConfig.sqlite_datareader["member_id"]);
                    }
                    textBox1.Text = _bookName;
                    textBox2.Text = _writer;
                    textBox3.Text = _publish;
                    comboBox1.SelectedIndex = _categoryId - 1;
                    comboBox2.SelectedIndex = _memberId - 1;

                    if (_status == "可借出")
                    {
                        radioButton1.Checked = true;
                        comboBox2.Enabled = false;
                    }
                    else
                    {
                        radioButton2.Checked = true;
                        textBox2.Enabled = true;
                    }
                }
                DBConfig.sqlite_datareader.Close();
            }
        }

        /// <summary>
        /// 下拉式選單的內容
        /// </summary>
        private void GetMemberData()
        {

            var command = DBConfig.sqlite_connect.CreateCommand();
            string sql = @"SELECT member_id, member_name
                            FROM member";

            command.CommandText = sql;
            DBConfig.sqlite_datareader = command.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    string member_id = Convert.ToString(DBConfig.sqlite_datareader["member_id"]);
                    string member_name = Convert.ToString(DBConfig.sqlite_datareader["member_name"]);

                    comboBox2.Items.Add(new ComboBoxItem(member_id, member_id + " " + member_name));
                }
            }
            DBConfig.sqlite_datareader.Close();
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
        /// 更新按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
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
            else if (comboBox2.SelectedIndex > -1)
            {
                command.Parameters.AddWithValue("@status", radioButton2.Text);
                command.Parameters.AddWithValue("@book_keeper", ComboUtil.GetItem(comboBox2).Value);
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
                    var confirmResult = MessageBox.Show("更新成功，是否回到主畫面？", "更新成功！！", MessageBoxButtons.YesNo);
                    if (confirmResult == DialogResult.Yes)
                    {
                        this.Close();
                    }
                    else
                    {
                        GetThisBookData();
                    }
                }
                catch (SqlException ex)
                {
                    //直接輸出錯誤訊息 MessageBox.Show(ex.Message);
                    MessageBox.Show("系統錯誤");
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
                comboBox2.SelectedIndex = -1;
                comboBox2.Enabled = false;
            }
            else if (radioButton2.Checked == true)
            {
                comboBox2.Enabled = true;
            }
        }

        /// <summary>
        /// 刪除按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            var bookId = label7.Text;
            //判斷借閱狀態，已借出不可刪除
            if (_status == "可借出")
            {
                var confirmResult = MessageBox.Show("確定刪除此書？", "刪除書籍！！", MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {
                    var command = DBConfig.sqlite_connect.CreateCommand();
                    string sql = @"DELETE FROM book_data WHERE book_id = @book_id";
                    command.CommandText = sql;
                    command.Parameters.AddWithValue("@book_id", bookId);
                    try
                    {
                        command.ExecuteNonQuery();
                        MessageBox.Show("刪除成功");
                        this.Close();
                    }
                    catch (SQLiteException ex)
                    {
                        MessageBox.Show("系統錯誤");
                    }
                }
            }
            else
            {
                MessageBox.Show("外借中不可刪除");
            }

        }
    }
}
