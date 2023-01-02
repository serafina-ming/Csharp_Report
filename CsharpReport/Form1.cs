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
using System.Net;

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
            //設定統計圖表
            UpdateChart();
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
            //新增後，重新載入統計圖表資料
            UpdateChart();
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
                    //取得書籍編號
                    int bookId = (int)dataGridView1.Rows[e.RowIndex].Cells["Column1"].Value;
                    var a = new Object[] { bookId };
                    editForm editForm;
                    editForm = new editForm();
                    editForm.setValue = a;
                    editForm.GetThisBookData();
                    editForm.ShowDialog();
                    GetBookData();
                    UpdateChart();
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
                                UpdateChart();
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
        /// 匯出資料按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            List<bookData> result = new List<bookData>();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                result.Add(new bookData()
                {
                    bookId = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value),
                    bookName = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value),
                    writer = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value),
                    publish = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value),
                    categoryName = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value),
                    status = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value),
                    memberName = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value)
                });
            }

            var dataSend = new Object[] { "export", result };
            exportOrImportForm exportOrImportForm;
            exportOrImportForm = new exportOrImportForm();
            exportOrImportForm.setValue = dataSend;
            exportOrImportForm.ShowDialog();
        }

        /// <summary>
        /// 匯入資料按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            var dataSend = new Object[] { "import" };
            exportOrImportForm exportOrImportForm;
            exportOrImportForm = new exportOrImportForm();
            exportOrImportForm.setValue = dataSend;
            exportOrImportForm.ShowDialog();
            GetBookData();
            UpdateChart();
        }

        /// <summary>
        /// 設定統計圖表值(長條圖)
        /// </summary>
        /// <param name="i_sort_bar"></param>
        private void SetBar(Dictionary<string, int> i_sort_bar)
        {
            this.chart1.Series["類型"].Points.Clear();
            foreach (var OneItem in i_sort_bar)
            {
                this.chart1.Series["類型"].Points.AddXY(OneItem.Key, OneItem.Value);
            }
        }

        /// <summary>
        /// 設定統計圖表值(圓餅圖)
        /// </summary>
        /// <param name="i_sort_pie"></param>
        private void SetPie(Dictionary<string, int> i_sort_pie)
        {
            foreach (var OneItem in i_sort_pie)
            {
                this.chart2.Series["類型"].Points.AddXY(OneItem.Key, OneItem.Value);
            }
        }

        /// <summary>
        /// 設定統計圖表值(折線圖)
        /// </summary>
        /// <param name="i_sort_line"></param>
        private void SetLine(Dictionary<string, int> i_sort_line)
        {
            int index = 0;
            foreach (var OneItem in i_sort_line)
            {
                index = this.chart3.Series["類型"].Points.AddXY(OneItem.Key, OneItem.Value);
            }
        }

        /// <summary>
        /// 取得統計圖表資料並更新
        /// </summary>
        public void UpdateChart()
        {
            Dictionary<string, int> _sort_data = new Dictionary<string, int>();
            var command = DBConfig.sqlite_connect.CreateCommand();
            string sql = @"SELECT category_name, count(*) AS count
                            FROM book_data
                            LEFT JOIN category_data
                            ON category = category_id
                            GROUP BY category;";

            command.CommandText = sql;
            DBConfig.sqlite_datareader = command.ExecuteReader();

            if (DBConfig.sqlite_datareader.HasRows)
            {
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    _sort_data.Add(Convert.ToString(DBConfig.sqlite_datareader["category_name"]), Convert.ToInt32(DBConfig.sqlite_datareader["count"]));
                }
                DBConfig.sqlite_datareader.Close();
            }

            SetBar(_sort_data);
            SetPie(_sort_data);
            SetLine(_sort_data);
        }

        /// <summary>
        /// 匯出統計圖表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            var a = new Object[] { chart1, chart2, chart3 };
            exportChartForm exportChartForm;
            exportChartForm = new exportChartForm();
            exportChartForm.setValue = a;
            exportChartForm.ShowDialog();
        }

        /// <summary>
        /// 匯出統計資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "書籍類型統計資料";
            save.Filter = "*.xlsx|*.xlsx";
            if (save.ShowDialog() != DialogResult.OK) return;

            // Excel 物件

            Microsoft.Office.Interop.Excel.Application xls = null;
            try
            {
                xls = new Microsoft.Office.Interop.Excel.Application();
                // Excel WorkBook
                Microsoft.Office.Interop.Excel.Workbook book = xls.Workbooks.Add();
                //Excel.Worksheet Sheet = (Excel.Worksheet)book.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet Sheet = xls.ActiveSheet;

                // 把資料塞進 Excel 內

                // 標題
                Sheet.Cells[1, 1] = "類型";
                Sheet.Cells[1, 2] = "數量";

                // 內容
                for (int k = 0; k < this.chart1.Series["類型"].Points.Count; k++)
                {
                    Sheet.Cells[k + 2, 1] = this.chart1.Series["類型"].Points[k].AxisLabel.ToString();
                    Sheet.Cells[k + 2, 2] = this.chart1.Series["類型"].Points[k].YValues[0].ToString();
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
    }
}
