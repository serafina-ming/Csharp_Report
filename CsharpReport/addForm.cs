﻿using System;
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
    public partial class addForm : Form
    {
        public addForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 新增資料按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Load_DB();
            //檢查輸入值是否都不為空值
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && comboBox1.SelectedIndex > -1)
            {
                if (AddBook(textBox1.Text, textBox2.Text, textBox3.Text, comboBox1.SelectedIndex + 1))
                {
                    MessageBox.Show("新增成功");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    comboBox1.SelectedIndex = -1;
                    //this.Close();
                }
                else
                {
                    MessageBox.Show("資料有誤");
                }
            }
            else
            {
                MessageBox.Show("輸入不完全");
            }
        }

        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        /// <summary>
        /// 新增書籍資料至資料庫
        /// </summary>
        /// <param name="book_name">書名</param>
        /// <param name="writer">作者</param>
        /// <param name="publish">出版社</param>
        /// <param name="category">類型</param>
        /// <returns></returns>
        public bool AddBook(string book_name, string writer, string publish, int category)
        {
            var command = DBConfig.sqlite_connect.CreateCommand();
            command.CommandText = @"INSERT INTO book_data (book_name, writer, publish, category, status)
                                        VALUES (@book_name, @writer, @publish, @category, @status);";
            command.Parameters.AddWithValue("@book_name", book_name);
            command.Parameters.AddWithValue("@writer", writer);
            command.Parameters.AddWithValue("@publish", publish);
            command.Parameters.AddWithValue("@category", (category));
            command.Parameters.AddWithValue("@status", "可借出");
            var result = command.ExecuteNonQuery();

            if (result != 0)
            {
                return true;
            }

            return false;
        }
    }
}
