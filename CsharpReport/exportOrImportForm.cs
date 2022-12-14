using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Policy;
using System.Reflection.Emit;
using static CsharpReport.Form1;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Data.SQLite;

namespace CsharpReport
{
    public partial class exportOrImportForm : Form
    {
        string mode = "export";
        object[][] dataGridView1 = { new object[] { "a", "1", "a", "1", "a", "1", "2" }, new object[] { "c", "b", "s", "5", "r", "g", "b" } };
        public exportOrImportForm()
        {
            InitializeComponent();
            if (mode == "import")
            {
                button2.Visible = false;
            }
        }

        /// <summary>
        /// 設定初始值，從Form1傳入
        /// </summary>
        public object[] setValue
        {
            set
            {
                label1.Text = "請選擇匯入資料格式：";
                button1.Text = "匯入書籍資料";
                mode = "import";
            }
        }

        /// <summary>
        /// 連接資料庫
        /// </summary>
        private void Load_DB()
        {
            DBConfig.sqlite_connect = new SQLiteConnection(DBConfig.dbPath);
            DBConfig.sqlite_connect.Open();// Open

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (mode == "export")
            {
                if (GetDataType() == "excel")
                {
                    // 設定儲存excel檔
                    SaveFileDialog save = new SaveFileDialog();
                    save.InitialDirectory =
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    save.FileName = "Export_Chart_Data7";
                    save.Filter = "*.xlsx|*.xlsx";
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
                        for (int i = 0; i < dataGridView1.Length; i++)
                        {
                            Sheet.Cells[i + 2, 1] = Convert.ToString(dataGridView1[i][0]);
                            Sheet.Cells[i + 2, 2] = Convert.ToString(dataGridView1[i][1]);
                            Sheet.Cells[i + 2, 3] = Convert.ToString(dataGridView1[i][2]);
                            Sheet.Cells[i + 2, 4] = Convert.ToString(dataGridView1[i][3]);
                            Sheet.Cells[i + 2, 5] = Convert.ToString(dataGridView1[i][4]);
                            Sheet.Cells[i + 2, 6] = Convert.ToString(dataGridView1[i][5]);
                            Sheet.Cells[i + 2, 7] = Convert.ToString(dataGridView1[i][6]);

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

        /// <summary>
        /// 匯出所有書籍資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            //判斷欲匯出的資料型態
            if (GetDataType() == "excel")
            {
                // 設定儲存excel檔
                SaveFileDialog save = new SaveFileDialog();
                save.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                save.FileName = "Export_Chart_Data7";
                save.Filter = "*.xlsx|*.xlsx";
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
                    var command = DBConfig.sqlite_connect.CreateCommand();
                    string sql = @"SELECT book_id, book_name, writer, publish,
                            category_name, status, member_name
                            FROM book_data
                            LEFT JOIN category_data
                            ON category = category_id
                            LEFT JOIN member
                            ON book_keeper = member_id";

                    command.CommandText = sql;
                    DBConfig.sqlite_datareader = command.ExecuteReader();

                    if (DBConfig.sqlite_datareader.HasRows)
                    {
                        int i = 0;
                        while (DBConfig.sqlite_datareader.Read()) //read every data
                        {
                            string _bookId = Convert.ToString(DBConfig.sqlite_datareader["book_id"]);
                            string _bookName = Convert.ToString(DBConfig.sqlite_datareader["book_name"]);
                            string _writer = Convert.ToString(DBConfig.sqlite_datareader["writer"]);
                            string _publish = Convert.ToString(DBConfig.sqlite_datareader["publish"]);
                            string _categoryName = Convert.ToString(DBConfig.sqlite_datareader["category_name"]);
                            string _status = Convert.ToString(DBConfig.sqlite_datareader["status"]);
                            string _memberName = Convert.ToString(DBConfig.sqlite_datareader["member_name"]);

                            Sheet.Cells[i + 2, 1] = _bookId;
                            Sheet.Cells[i + 2, 2] = _bookName;
                            Sheet.Cells[i + 2, 3] = _writer;
                            Sheet.Cells[i + 2, 4] = _publish;
                            Sheet.Cells[i + 2, 5] = _categoryName;
                            Sheet.Cells[i + 2, 6] = _status;
                            Sheet.Cells[i + 2, 7] = _memberName;
                            i = i+ 1;
                        }
                        DBConfig.sqlite_datareader.Close();
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

        //取得欲匯出的資料型態
        private string GetDataType()
        {
            if (radioButton1.Checked == true)
            {
                return "excel";
            }
            else if (radioButton2.Checked == true)
            {
                return "csv";
            }
            else if (radioButton3.Checked == true)
            {
                return "json";
            }
            else
            {
                return "wrong";
            }
        }
    }
}
