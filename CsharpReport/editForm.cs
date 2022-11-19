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
        public editForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 設定初始值，從Form1傳入
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
                    var confirmResult = MessageBox.Show("更新成功，是否回到主畫面？",
                                     "更新成功！！",
                                     MessageBoxButtons.YesNo);
                    if (confirmResult == DialogResult.Yes)
                    {
                        this.Close();
                    }
                    else
                    {
                        // If 'No', do something here.
                    }
                }
                catch (SqlException ex)
                {
                    //MessageBox.Show(ex.Message);
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
                textBox4.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var bookId = label7.Text;
            var confirmResult = MessageBox.Show("確定刪除此書？",
                                     "刪除書籍！！",
                                     MessageBoxButtons.YesNo);
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
            else
            {
                // If 'No', do something here.
            }
        }
    }
}
