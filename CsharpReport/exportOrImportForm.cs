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
using System.Net.NetworkInformation;
using System.Net;
using Newtonsoft.Json;
using System.Runtime.Remoting.Messaging;

namespace CsharpReport
{
    public partial class exportOrImportForm : Form
    {
        string mode = "export";
        List<bookData> dataGridViewdata;
        public exportOrImportForm()
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
                mode = value[0].ToString();
                if (value[0].ToString() == "export")
                {
                    dataGridViewdata = (List<bookData>)value[1];
                }
                else
                {
                    label1.Text = "請選擇匯入資料格式：";
                    button1.Text = "匯入書籍資料";
                    button2.Visible = false;
                }
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

        /// <summary>
        /// 匯出查詢結果或匯入書籍資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (mode == "export")
            {
                if (GetDataType() == "excel")
                {
                    ExportExcel("search");
                }
                else if (GetDataType() == "csv")
                {
                    ExportCsv("search");
                }
                else if (GetDataType() == "json")
                {
                    ExportJson("search");
                }
                else
                {
                    MessageBox.Show(GetDataType());
                }
            }
            else if (mode == "import")
            {
                if (GetDataType() == "excel")
                {
                    // 讀取excel檔的檔案選擇視窗
                    OpenFileDialog open = new OpenFileDialog();
                    open.InitialDirectory =
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    open.Filter = "*.xlsx|*.xlsx";
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
                        for (var i = 2; i < Sheet.UsedRange.Rows.Count; i++)
                        {
                            int category = GetCategoryId(Sheet.Cells[i, 5].Value.ToString());

                            addForm.AddBook(Sheet.Cells[i, 2].Value.ToString(), Sheet.Cells[i, 3].Value.ToString(), Sheet.Cells[i, 4].Value.ToString(), category);
                        }
                        MessageBox.Show("成功匯入資料");
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
                else if (GetDataType() == "csv") 
                {
                    OpenFileDialog open = new OpenFileDialog();
                    open.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    if (open.ShowDialog() != DialogResult.OK) return;

                    System.IO.StreamReader sr = new System.IO.StreamReader(open.FileName);

                    string firstLine = sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split(',');
                        int category = GetCategoryId(words[4]);
                        addForm addForm = new addForm();
                        addForm.AddBook(words[1], words[2], words[3], category);
                    }
                    MessageBox.Show("成功匯入資料");
                }
                else if (GetDataType() == "json")
                {
                    // 打開json檔
                    OpenFileDialog open = new OpenFileDialog();
                    open.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    if (open.ShowDialog() != DialogResult.OK) return;

                    System.IO.StreamReader sr = new System.IO.StreamReader(open.FileName);

                    string lines = sr.ReadToEnd();
                    //Newtonsoft.Json反序列化: collection
                    List<bookData> json_collection = JsonConvert.DeserializeObject<List<bookData>>(lines);

                    
                }
                else
                {
                    MessageBox.Show(GetDataType());
                }
            }
        }

        /// <summary>
        /// 依據類型中文抓取類型id
        /// </summary>
        /// <param name="categoryName"></param>
        /// <returns></returns>
        private int GetCategoryId(string categoryName)
        {
            if (categoryName == "文學")
            {
                return 1;
            }
            else if (categoryName == "飲食料理")
            {
                return 2;
            }
            else if (categoryName == "心靈勵志")
            {
                return 3;
            }
            else if (categoryName == "漫畫")
            {
                return 4;
            }
            else if (categoryName == "輕小說")
            {
                return 5;
            }
            else if (categoryName == "電腦資訊")
            {
                return 6;
            }
            else if (categoryName == "藝術設計")
            {
                return 7;
            }
            return -1;
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
                ExportExcel("all");
            }
            else if (GetDataType() == "csv")
            {
                ExportCsv("all");
            }
            else if (GetDataType() == "json")
            {
                ExportJson("all");
            }
            else
            {
                MessageBox.Show(GetDataType());
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

        /// <summary>
        /// 取得所有書籍資料
        /// </summary>
        /// <returns></returns>
        public List<bookData> GetBookData() 
        {
            List<bookData> result = new List<bookData>();
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
                while (DBConfig.sqlite_datareader.Read()) //read every data
                {
                    result.Add(new bookData() 
                    {
                        bookId = Convert.ToString(DBConfig.sqlite_datareader["book_id"]),
                        bookName = Convert.ToString(DBConfig.sqlite_datareader["book_name"]),
                        writer = Convert.ToString(DBConfig.sqlite_datareader["writer"]),
                        publish = Convert.ToString(DBConfig.sqlite_datareader["publish"]),
                        categoryName = Convert.ToString(DBConfig.sqlite_datareader["category_name"]),
                        status = Convert.ToString(DBConfig.sqlite_datareader["status"]),
                        memberName = Convert.ToString(DBConfig.sqlite_datareader["member_name"])
                    });
                }
                DBConfig.sqlite_datareader.Close();
            }
            return result;
        }

        /// <summary>
        /// 匯出excel
        /// </summary>
        /// <param name="dataMode">判斷匯出資料範圍</param>
        private void ExportExcel(string dataMode)
        {
            List<bookData> bookDataModel = new List<bookData>();
            if (dataMode == "all")
            {
                bookDataModel = GetBookData();
            }
            else if (dataMode == "search")
            {
                bookDataModel = dataGridViewdata;
            }

            // 設定儲存excel檔
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "書籍資料";
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
                int i = 0;
                foreach (bookData bookData in bookDataModel)
                {
                    Sheet.Cells[i + 2, 1] = bookData.bookId;
                    Sheet.Cells[i + 2, 2] = bookData.bookName;
                    Sheet.Cells[i + 2, 3] = bookData.writer;
                    Sheet.Cells[i + 2, 4] = bookData.publish;
                    Sheet.Cells[i + 2, 5] = bookData.categoryName;
                    Sheet.Cells[i + 2, 6] = bookData.status;
                    Sheet.Cells[i + 2, 7] = bookData.memberName;
                    i++;
                }

                // 儲存檔案
                book.SaveAs(save.FileName);
                MessageBox.Show("成功匯出資料");
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

        /// <summary>
        /// 匯出csv
        /// </summary>
        /// <param name="dataMode">判斷匯出資料範圍</param>
        private void ExportCsv(string dataMode)
        {
            List<bookData> bookDataModel = new List<bookData>();
            if (dataMode == "all")
            {
                bookDataModel = GetBookData();
            }
            else if (dataMode == "search")
            {
                bookDataModel = dataGridViewdata;
            }

            // 設定儲存excel檔
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "書籍資料.csv";
            if (save.ShowDialog() != DialogResult.OK) return;
            string strFilePath = save.FileName;
            StringBuilder sbOutput = new StringBuilder();

            string tmp = String.Format("書籍編號,書名,作者,出版社,類型,借閱狀態,借閱人");
            sbOutput.AppendLine(tmp);
            int i = 0;
            foreach (bookData bookData in bookDataModel)
            {
                tmp = String.Format("{0}", bookData.bookId);
                tmp = String.Format("{0},{1}", tmp, bookData.bookName);
                tmp = String.Format("{0},{1}", tmp, bookData.writer);
                tmp = String.Format("{0},{1}", tmp, bookData.publish);
                tmp = String.Format("{0},{1}", tmp, bookData.categoryName);
                tmp = String.Format("{0},{1}", tmp, bookData.status);
                tmp = String.Format("{0},{1}", tmp, bookData.memberName);

                sbOutput.AppendLine(tmp);
                i++;
            }
            // Create and write the csv file
            System.IO.File.WriteAllText(strFilePath, sbOutput.ToString(), Encoding.UTF8);

            MessageBox.Show("成功匯出資料");
        }

        /// <summary>
        /// 匯出json
        /// </summary>
        /// <param name="dataMode">判斷匯出資料範圍</param>
        private void ExportJson(string dataMode)
        {
            List<bookData> bookDataModel = new List<bookData>();
            if (dataMode == "all")
            {
                bookDataModel = GetBookData();
            }
            else if (dataMode == "search")
            {
                bookDataModel = dataGridViewdata;
            }

            // 設定儲存json檔
            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "書籍資料.json";
            if (save.ShowDialog() != DialogResult.OK) return;

            string strFilePath = save.FileName;

            //Newtonsoft.Json序列化
            string jsonData = JsonConvert.SerializeObject(bookDataModel);

            System.IO.File.WriteAllText(strFilePath, jsonData);

            MessageBox.Show("成功匯出資料");
        }
    }

    public class bookData
    {
        public string bookId { get; set; }
        public string bookName { get; set; }
        public string writer { get; set; }
        public string publish { get; set; }
        public string categoryName { get; set; }
        public string status { get; set; }
        public string memberName { get; set; }
    }
}
