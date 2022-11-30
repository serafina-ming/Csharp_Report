using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Policy;

namespace CsharpReport
{
    public partial class Form1 : Form
    {
        int index = 1;

        /// <summary>
        /// 資料庫設定
        /// </summary>
        public class DBConfig
        {
            public static string dbFile = Application.StartupPath + @"\book.db";

            public static string dbPath = "Data source=" + dbFile;

            public static SQLiteConnection sqlite_connect;
            public static SQLiteCommand sqlite_cmd;
            public static SQLiteDataReader sqlite_datareader;
        }

        /// <summary>
        /// 連接資料庫
        /// </summary>
        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        /// <summary>
        /// 取得書籍資料並顯示於dataGrid
        /// </summary>
        private void GetBookData()
        {
            this.dataGridView1.Rows.Clear();

            var command = DBConfig.sqlite_connect.CreateCommand();
            string sql = @"SELECT book_id, book_name, writer, publish,
                            category_name, status, member_name
                            FROM book_data
                            LEFT JOIN category_data
                            ON category = category_id
                            LEFT JOIN member
                            ON book_keeper = member_id
                            WHERE book_name LIKE @book_name
                            AND writer LIKE @writer
                            AND publish LIKE @publish ";
            string status = "";

            if (comboBox1.SelectedIndex > 0)
            {
                sql += "AND category = @category ";
            }
            if((checkBox1.Checked || checkBox2.Checked) && !(checkBox1.Checked && checkBox2.Checked))
            {
                sql += "AND status = @status ";
                if(checkBox1.Checked)
                {
                    status = checkBox1.Text;
                }
                else
                {
                    status = checkBox2.Text;
                }
            }
            
            command.CommandText = sql;
            command.Parameters.AddWithValue("@book_name", "%"+textBox1.Text+"%");
            command.Parameters.AddWithValue("@writer", "%"+textBox2.Text+"%");
            command.Parameters.AddWithValue("@publish", "%"+textBox3.Text+"%");
            command.Parameters.AddWithValue("@category", comboBox1.SelectedIndex);
            command.Parameters.AddWithValue("@status", status);
            DBConfig.sqlite_datareader = command.ExecuteReader();

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
                    rows.Add(new Object[] { index, _bookName, _writer, _publish, _categoryName, _status, _memberName, "編輯", "刪除" });
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
            //設定dataGrid資料
            GetBookData();
        }

        /// <summary>
        /// menuStrip說明按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("關於");
        }

        /// <summary>
        /// 查詢功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            GetBookData();
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

        /// <summary>
        /// 顯示新增資料視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            addForm addForm;
            addForm = new addForm();
            addForm.ShowDialog();
            //新增後，重新載入dataGrid資料
            GetBookData();
        }

        /// <summary>
        /// 設定dataGrid編輯刪除按鈕功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                if (e.ColumnIndex == 7 && dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value != null)
                {
                    //點下去顯示書籍編號
                    //MessageBox.Show(dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value.ToString());
                    int bookId = (int)dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value;
                    var a = new Object[] { bookId };
                    editForm editForm;
                    editForm = new editForm();
                    editForm.setValue = a;
                    editForm.GetThisBookData();
                    editForm.ShowDialog();
                    GetBookData();
                }
                else if (e.ColumnIndex == 8 && dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value != null)
                {
                    int bookId = (int)dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value;
                    string bookName = (string)dataGridView1.Rows[e.RowIndex].Cells["Column2"].Value;
                    if (dataGridView1.Rows[e.RowIndex].Cells["Column6"].Value.ToString() == "可借出") 
                    {
                        var confirmResult = MessageBox.Show("確定刪除《" + bookName + "》？",
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
                                //刪除後，重新載入dataGrid資料
                                GetBookData();
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
                    else
                    {
                        MessageBox.Show("外借中不可刪除");
                    }
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

        /// <summary>
        /// 匯出excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            // 設定儲存excel檔
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "Export_Chart_Data7" ;
            save.Filter = "*.xlsx|*.xlsx" ;
            if (save.ShowDialog() != DialogResult.OK) return;
            // Excel 物件
            Excel.Application xls = null;
            try
            {
                // 打開excel
                xls = new Excel.Application();
                // 新增第一個sheet
                // Excel WorkBook
                Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Excel.Worksheet Sheet = xls.ActiveSheet;
                // 把資料塞進 Excel 內
                // 標題
                Sheet.Cells[1, 1] = "書籍編號";
                Sheet.Cells[1, 2] = "書名";
                Sheet.Cells[1, 3] = "作者";
                Sheet.Cells[1, 4] = "出版社";
                Sheet.Cells[1, 5] = "類型";
                Sheet.Cells[1, 6] = "借閱狀態";
                Sheet.Cells[1, 7] = "借閱人";
                // 內容
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    //MessageBox.Show("書籍編號：" + Convert.ToString(dataGridView1.Rows[i].Cells[0].Value) +
                    //    "\n書名：" + Convert.ToString(dataGridView1.Rows[i].Cells[1].Value) +
                    //    "\n作者：" + Convert.ToString(dataGridView1.Rows[i].Cells[2].Value) +
                    //    "\n出版社：" + Convert.ToString(dataGridView1.Rows[i].Cells[3].Value) +
                    //    "\n類型：" + Convert.ToString(dataGridView1.Rows[i].Cells[4].Value) +
                    //    "\n借閱狀態：" + Convert.ToString(dataGridView1.Rows[i].Cells[5].Value) +
                    //    "\n借閱人 " + Convert.ToString(dataGridView1.Rows[i].Cells[6].Value) +
                    //    "\ni=" + i +
                    //    "\nRows.Count=" + dataGridView1.Rows.Count
                    //    );
                    Sheet.Cells[i + 2, 1] = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                    Sheet.Cells[i + 2, 2] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                    Sheet.Cells[i + 2, 3] = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                    Sheet.Cells[i + 2, 4] = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                    Sheet.Cells[i + 2, 5] = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                    Sheet.Cells[i + 2, 6] = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                    Sheet.Cells[i + 2, 7] = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);

                }

                // 儲存檔案
                book.SaveAs(save.FileName);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                xls.Quit();
            }
        }

        //將excel的資料讀取到資料庫
        private void button4_Click(object sender, EventArgs e)
        {
            // 讀取excel檔的檔案選擇視窗
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            open.Filter = "*.xlsx|*.xlsx" ;
            if (open.ShowDialog() != DialogResult.OK) return;
            MessageBox.Show(open.FileName);
            // Excel 物件
            Excel.Application xls = null;
            try
            {
                // 打開excel
                xls = new Excel.Application();
                // 打開第一個sheet
                // Excel WorkBook
                Excel.Workbook book = xls.Workbooks.Open(open.FileName);
                Excel.Worksheet Sheet = xls.ActiveSheet;

                //新增書籍
                addForm addForm = new addForm();
                // 讀取cell
                for (var i = 2; i <Sheet.UsedRange.Rows.Count; i++) 
                {
                    int category = 1;
                    if (Sheet.Cells[i, 5].Value.ToString() == "文學")
                    {
                        category = 1;
                    }
                    else if (Sheet.Cells[i, 5].Value.ToString() == "飲食料理")
                    {
                        category = 2;
                    }
                    else if (Sheet.Cells[i, 5].Value.ToString() == "心靈勵志")
                    {
                        category = 3;
                    }
                    else if (Sheet.Cells[i, 5].Value.ToString() == "漫畫")
                    {
                        category = 4;
                    }
                    else if (Sheet.Cells[i, 5].Value.ToString() == "輕小說")
                    {
                        category = 5;
                    }
                    else if (Sheet.Cells[i, 5].Value.ToString() == "電腦資訊")
                    {
                        category = 5;
                    }
                    else if (Sheet.Cells[i, 5].Value.ToString() == "藝術設計")
                    {
                        category = 5;
                    }

                    addForm.AddBook(Sheet.Cells[i, 2].Value.ToString(), Sheet.Cells[i, 3].Value.ToString(), Sheet.Cells[i, 4].Value.ToString(), category);
                }
            }
            catch (Exception a)
            {
                MessageBox.Show(a.ToString());
                throw;
            }
            finally
            {
                xls.Quit();
            }
        }

        /// <summary>
        /// 匯出csv檔
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            // 設定儲存excel檔
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName =" Export_Chart_Data.csv" ;
            if (save.ShowDialog() != DialogResult.OK) return;
            string strFilePath = save.FileName;
            StringBuilder sbOutput = new StringBuilder();

            string tmp = String.Format("書籍編號,書名,作者,出版社,類型,借閱狀態,借閱人");
            sbOutput.AppendLine(tmp);
            for (int i = 0 ; i < dataGridView1.Rows.Count; i++)
            {
                tmp = String.Format("{0}", Convert.ToString(dataGridView1.Rows[i].Cells[0].Value));
                for (int j = 1 ; j < 7 ; j++)
                {
                    tmp = String.Format("{0},{1}", tmp , Convert.ToString(dataGridView1.Rows[i].Cells[j].Value));
                }
                sbOutput.AppendLine(tmp);
            }
            // Create and write the csv file
            System.IO.File.WriteAllText(strFilePath, sbOutput.ToString(), Encoding.UTF8);
            // To append more lines to the csv file
            //System.IO.File.AppendAllText(strFilePath, sbOutput.ToString(), Encoding.UTF8);
        }
    }
}
