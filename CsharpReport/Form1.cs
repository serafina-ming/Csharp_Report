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
using iText.Forms;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Extgstate;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Pdf.Canvas.Draw;
using TabAlignment = iText.Layout.Properties.TabAlignment;
using iText.Layout.Borders;

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

        private void button3_Click(object sender, EventArgs e)
        {
            exportOrImportForm exportOrImportForm;
            exportOrImportForm = new exportOrImportForm();
            exportOrImportForm.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PrintPDF();
        }

        void PrintPDF()
        {
            // Set the output dir and file name
            // string directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            string src = "./test.pdf";
            string dst = @"./new_test.pdf";

            //manipulatePdf(src, dst);

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); ;
            saveFileDialog.FileName = "書籍資料";
            saveFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                dst = saveFileDialog.FileName;
                manipulatePdf(src, dst);
            }

        }

        void manipulatePdf(String src, String dst)
        {
            // 1. create pdf
            PdfWriter writer = new PdfWriter(dst);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            // 標楷體
            PdfFont font_tr = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);
            // 正黑體
            //PdfFont font_msjh = PdfFontFactory.CreateFont(@"./msjh.ttf", PdfEncodings.IDENTITY_H);
            document.SetFont(font_tr); // 預設中文字型
            
            // 2. 加文字

            // 2.1. add header
            Paragraph header_1 = new Paragraph("書籍資料")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(36)
               .SetFont(font_tr);
            document.Add(header_1);

            //Paragraph header_2 = new Paragraph("文件 標題2")
            //   .SetTextAlignment(TextAlignment.CENTER)
            //   .SetFontSize(24);
            //document.Add(header_2);

            // Line separator
            LineSeparator ls = new LineSeparator(new SolidLine());
            document.Add(ls);



            float[] colWidths = { 1, 6, 3, 2, 2, 2, 2};
            // Creating a table       
            Table table = new Table(UnitValue.CreatePercentArray(colWidths));

            // Add paragraph 1
            //增加段落一
            Paragraph paragraph = new Paragraph();

            //TabStop TabStop1 = new TabStop(50f);
            //paragraph.AddTabStops(TabStop1);
            //Tab Tab1 = new Tab();
            //TabStop TabStop2 = new TabStop(250f);
            //paragraph.AddTabStops(TabStop2);
            //Tab Tab2 = new Tab();
            //讀取grid資料
            String Gridparagraph = "";
            //Gridparagraph = Gridparagraph + "書籍編號\t書名\t作者\t出版社\t類型\t借閱狀態\t借閱人\n";

            //paragraph.Add("編號").Add(Tab1).Add("書名").Add(Tab2).Add("作者").Add(Tab1).Add("出版社").Add(Tab1).Add("類型").Add(Tab1).Add("借閱狀態").Add(Tab1).Add("借閱人\n");

 
            table.AddCell("編號");
            table.AddCell("書名");
            table.AddCell("作者");
            table.AddCell("出版社");
            table.AddCell("類型");
            table.AddCell("借閱狀態");
            table.AddCell("借閱人");
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j <= 6; j++)
                {
                    //Gridparagraph = Gridparagraph + Convert.ToString(dataGridView1.Rows[i].Cells[j].Value)+"\t";
                    Gridparagraph = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    // Adding cells to the table       
                    Cell cell = new Cell().Add(new Paragraph(Gridparagraph));
                    cell.SetBorder(new SolidBorder(ColorConstants.WHITE, 0));
                    cell.SetBorderBottom(new SolidBorder(ColorConstants.GRAY, 0));
                    //table.addCell(cell);
                    table.AddCell(cell);
                    //if (j == 1)
                    //{
                    //    paragraph.Add(Gridparagraph).Add(Tab1);
                    //}
                    //else 
                    //{ 
                    //    paragraph.Add(Gridparagraph).Add(Tab2);
                    //}
                }
                //Gridparagraph = Gridparagraph + "\n";
                //paragraph.Add("\n");
            }
            //paragraph.Add("111");
            //paragraph.AddTabStops(new TabStop(100f, TabAlignment.LEFT));
            //paragraph.Add(new Tab());
            //paragraph.Add(Gridparagraph);

            

            // Adding Table to document        
            document.Add(table);

            //document.Add(paragraph);

            // Add paragraph 2
            //設定在右下角
            paragraph = new Paragraph("根據《今日美國報》報導，美職28日賽後確定了16" +
                "支季後賽球隊的對戰席次，戰況膠著的國聯因巨人、費城人都輸球，釀酒人搶" +
                "下最後一張晉級門票，他們將作客對上頭號種子道奇，國聯第2種子勇士將迎戰" +
                "紅人，第3種子小熊將迎戰今年鹹魚大翻身的馬林魚，4、5種子教士與紅雀正面" +
                "對決。");
            paragraph.SetMarginLeft(123);
            paragraph.SetMarginTop(555);
            document.Add(paragraph);

            //// 隨便插一段話
            //PdfPage page1 = pdf.GetPage(4); // ##設定文字方塊插在哪一頁
            //PdfCanvas pdfCanvas1 = new PdfCanvas(page1);
            //iText.Kernel.Geom.Rectangle rectangle = new iText.Kernel.Geom.Rectangle(100, 700, 100, 100); // ##設定文字方塊的位置與大小
            //iText.Kernel.Colors.Color bgColour = new DeviceRgb(255, 504, 204);  // ##設定文字方塊的顏色
            //pdfCanvas1.SaveState()
            //        .SetFillColor(bgColour)
            //        .Rectangle(rectangle)
            //        .Fill()
            //        .RestoreState();
            ////iText.Layout.Canvas canvas = new iText.Layout.Canvas(pdfCanvas1, pdf, rectangle);
            //iText.Layout.Canvas canvas = new iText.Layout.Canvas(pdfCanvas1, rectangle);
            //canvas.Add(new Paragraph("隨便插一段話").SetFont(font_tr));  // ##設定文字方塊的文字

            // 4. close content
            //canvas.Close();
            document.Close();
        }

    }
}
