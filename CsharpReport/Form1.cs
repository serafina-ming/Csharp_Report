﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;

namespace CsharpReport
{
    public partial class Form1 : Form
    {
        int index = 1;

        public class DBConfig
        {
            public static string dbFile = Application.StartupPath + @"\book.db";

            public static string dbPath = "Data source=" + dbFile;

            public static SQLiteConnection sqlite_connect;
            public static SQLiteCommand sqlite_cmd;
            public static SQLiteDataReader sqlite_datareader;
        }

        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        private void Show_DB()
        {
            this.dataGridView1.Rows.Clear();

            string sql = @"SELECT book_id, book_name, writer, publish,
                            category_name, status, member_name
                            FROM book_data
                            LEFT JOIN category_data
                            ON category=category_id
                            LEFT JOIN member
                            ON book_keeper=member_id;";
            DBConfig.sqlite_cmd = new SQLiteCommand(sql, DBConfig.sqlite_connect);
            DBConfig.sqlite_datareader = DBConfig.sqlite_cmd.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    int _bookId = Convert.ToInt32(DBConfig.sqlite_datareader["book_id"]);
                    string _bookName = Convert.ToString(DBConfig.sqlite_datareader["book_name"]);
                    string _writer = Convert.ToString(DBConfig.sqlite_datareader["writer"]);
                    string _publish = Convert.ToString(DBConfig.sqlite_datareader["publish"]);
                    string _categoryName = Convert.ToString(DBConfig.sqlite_datareader["category_name"]);
                    string _status = Convert.ToString(DBConfig.sqlite_datareader["status"]);
                    string _memberName = Convert.ToString(DBConfig.sqlite_datareader["member_name"]);

                    index = _bookId;
                    DataGridViewRowCollection rows = dataGridView1.Rows;
                    rows.Add(new Object[] { index, _bookName, _writer, _publish, _categoryName, _status, _memberName });
                }
                DBConfig.sqlite_datareader.Close();
            }

        }


        public Form1()
        {
            InitializeComponent();

            //登入時顯示登入視窗
            loginForm loginForm;
            loginForm = new loginForm();
            loginForm.ShowDialog();

            //讀取資料庫
            Load_DB();
            Show_DB();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("關於");
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 登出後，顯示登入視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 登出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            loginForm loginForm;
            loginForm = new loginForm();
            loginForm.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex > -1)
            {
                
                //獲取當前被點擊的單元格
                DataGridViewButtonCell vCell = (DataGridViewButtonCell)dataGridView1.CurrentCell;

                //現在第幾欄
                //MessageBox.Show(e.ColumnIndex.ToString());
                //目前第幾筆
                //MessageBox.Show(e.RowIndex.ToString());
                //MessageBox.Show(dataGridView1.Columns[e.ColumnIndex].ToString());
                if (dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value != null)
                {
                    //點下去顯示書籍編號
                    MessageBox.Show(dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value.ToString());
                }
                else
                {
                    MessageBox.Show("查無此書");
                }
                
                //參考
                //https://www.twblogs.net/a/5e5578e8bd9eee2117c61e72
                //https://learn.microsoft.com/zh-tw/dotnet/api/system.windows.forms.datagridviewbuttoncolumn?view=windowsdesktop-7.0
            }
        }
    }
}
