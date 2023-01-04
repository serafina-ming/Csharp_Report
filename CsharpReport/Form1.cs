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
using iText.IO.Font;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Kernel.Pdf;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Forms;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Extgstate;
using iText.Layout;
using System.IO;
using System.Drawing.Imaging;
using Org.BouncyCastle.Utilities;

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
            if ((checkBox1.Checked || checkBox2.Checked) && !(checkBox1.Checked && checkBox2.Checked))
            {
                sql += "AND status = @status ";
                if (checkBox1.Checked)
                {
                    status = checkBox1.Text;
                }
                else
                {
                    status = checkBox2.Text;
                }
            }

            command.CommandText = sql;
            command.Parameters.AddWithValue("@book_name", "%" + textBox1.Text + "%");
            command.Parameters.AddWithValue("@writer", "%" + textBox2.Text + "%");
            command.Parameters.AddWithValue("@publish", "%" + textBox3.Text + "%");
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
        /// menuStrip說明按鈕(關於)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("簡易圖書系統\n可紀錄書籍資料，有新增、修改、刪除功能，還能匯出資料及圖表，以及匯出pdf檔跟QRcode。");
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
                //判斷點選的是編輯還是刪除
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
            this.chart2.Series["類型"].Points.Clear();
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
            this.chart3.Series["類型"].Points.Clear();
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
                MessageBox.Show("成功匯出統計資料");
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
        /// 匯出pdf按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            PrintPDF();
        }

        /// <summary>
        /// 設定pdf檔案
        /// </summary>
        void PrintPDF()
        {
            // Set the output dir and file name
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); ;
            saveFileDialog.FileName = "書籍資料";
            saveFileDialog.Filter = "*.pdf|*.pdf";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                string src = saveFileDialog.FileName;
                string dst = src + "_tmp";
                ManipulatePdf(src, dst);
            }

        }

        /// <summary>
        /// 設定pdf內容
        /// </summary>
        /// <param name="src"></param>
        /// <param name="dst"></param>
        void ManipulatePdf(String src, String dst)
        {
            // 1. create pdf
            PdfWriter writer = new PdfWriter(dst);
            PdfDocument pdf = new PdfDocument(writer);
            Document document = new Document(pdf);

            // 標楷體
            PdfFont font_tr = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);
            // 正黑體 PdfFont font_msjh = PdfFontFactory.CreateFont(@"./msjh.ttf", PdfEncodings.IDENTITY_H);
            document.SetFont(font_tr); // 預設中文字型

            // 2. 加文字

            // 2.1. add header
            Paragraph header_1 = new Paragraph("書籍資料")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(36)
               .SetFont(font_tr);
            document.Add(header_1);

            // Line separator
            LineSeparator ls = new LineSeparator(new SolidLine());
            document.Add(ls);

            float[] colWidths = { 1, 6, 3, 2, 2, 2, 2 };
            // Creating a table       
            Table table_bookData = new Table(UnitValue.CreatePercentArray(colWidths));
            int tableNum = dataGridView1.ColumnCount - 2;

            //表格標題列
            for (int k = 0; k < tableNum; k++)
            {
                table_bookData.AddHeaderCell(new Paragraph(dataGridView1.Columns[k].HeaderText));
            }

            //表格內容
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < tableNum; j++)
                {
                    string Gridparagraph = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    // Adding cells to the table       
                    Cell cell = new Cell().Add(new Paragraph(Gridparagraph));
                    cell.SetBorder(new SolidBorder(ColorConstants.WHITE, 0));
                    cell.SetBorderBottom(new SolidBorder(ColorConstants.GRAY, 0));
                    table_bookData.AddCell(cell);
                }
            }

            // Adding Table to document
            document.Add(table_bookData);
            //table_data.Complete(); 放這行資料table會重複輸出一次

            // move to next page
            // Creating an Area Break          
            AreaBreak a_ch_1 = new AreaBreak();

            // Adding area break to the PDF
            document.Add(a_ch_1);

            /////////////////////////////////
            // table chart 1

            Paragraph header_2 = new Paragraph("書籍類型統計圖表\n")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(24)
               .SetFont(font_tr);
            document.Add(header_2);

            Table table_ch_1 = new Table(1, true);
            table_ch_1.SetFont(font_tr);

            table_ch_1.AddHeaderCell("書籍類型數量");

            // 出貨各類別數量
            Bitmap bitmap_ch1_1 = new Bitmap(chart1.Width, chart1.Height, PixelFormat.Format24bppRgb);
            chart1.DrawToBitmap(bitmap_ch1_1, new System.Drawing.Rectangle(0, 0, chart1.Width, chart1.Height));
            ImageData imageData_ch1_1 = ImageDataFactory.Create(BmpToBytes(bitmap_ch1_1));
            iText.Layout.Element.Image image_ch1_1 = new iText.Layout.Element.Image(imageData_ch1_1);
            image_ch1_1.SetAutoScale(true);
            table_ch_1.AddCell(image_ch1_1);


            table_ch_1.AddCell("書籍類型數量比例");

            // 出貨各類別數量
            Bitmap bitmap_ch1_2 = new Bitmap(chart2.Width, chart2.Height, PixelFormat.Format24bppRgb);
            chart2.DrawToBitmap(bitmap_ch1_2, new System.Drawing.Rectangle(0, 0, chart2.Width, chart2.Height));
            ImageData imageData_ch1_2 = ImageDataFactory.Create(BmpToBytes(bitmap_ch1_2));
            iText.Layout.Element.Image image_ch1_2 = new iText.Layout.Element.Image(imageData_ch1_2);
            image_ch1_2.SetAutoScale(true);
            table_ch_1.AddCell(image_ch1_2);

            document.Add(table_ch_1);
            table_ch_1.Complete();

            // move to next page
            // Creating an Area Break          
            AreaBreak a_ch_2 = new AreaBreak();

            // Adding area break to the PDF       
            document.Add(a_ch_2);

            /////////////////////////////////
            // table chart 2

            Paragraph header_3 = new Paragraph("統計資料QRcode\n")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(24)
               .SetFont(font_tr);
            document.Add(header_3);

            Table table_ch_2 = new Table(2, true);
            table_ch_2.SetFont(font_tr);

            table_ch_2.AddHeaderCell("統計資料");
            table_ch_2.AddHeaderCell("統計資料 QR Code");

            string qrcodeData = GetQrcodeData();

            // 出貨統計資料
            table_ch_2.AddCell(new Paragraph(qrcodeData));

            // 出貨統計資料 QR Code
            System.Drawing.Bitmap bitmap_2 = GetQrcode(qrcodeData, pictureBox2.Width, pictureBox2.Height);
            ImageData imageData_2 = ImageDataFactory.Create(BmpToBytes(bitmap_2));
            iText.Layout.Element.Image image_ch2 = new iText.Layout.Element.Image(imageData_2);
            image_ch2.SetAutoScale(true);
            table_ch_2.AddCell(image_ch2);

            document.Add(table_ch_2);
            table_ch_2.Complete();

            document.Close();

            // 4. edit existed pdf
            PdfReader reader2 = new PdfReader(dst);
            PdfWriter writer2 = new PdfWriter(src);
            PdfDocument pdfDoc2 = new PdfDocument(reader2, writer2);
            Document document2 = new Document(pdfDoc2);

            // 5. add Page numbers
            DrawHeader(pdfDoc2, document2);
            document2.Close();
            File.Delete(dst);
            MessageBox.Show("成功匯出pdf");
        }

        /// <summary>
        /// 設定pdf頁首頁尾、浮水印
        /// </summary>
        /// <param name="pdfDoc"></param>
        /// <param name="document"></param>
        void DrawHeader(PdfDocument pdfDoc, Document document)
        {
            PdfFont font = PdfFontFactory.CreateFont(@"c:/Windows/fonts/kaiu.ttf", PdfEncodings.IDENTITY_H);
            iText.Kernel.Geom.Rectangle pageSize;
            PdfCanvas canvas;
            int n = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= n; i++)
            {
                PdfPage page = pdfDoc.GetPage(i);
                pageSize = page.GetPageSize();
                canvas = new PdfCanvas(page);

                //Draw header text
                canvas.BeginText()
                    .SetFontAndSize(font, 15)
                    .MoveText(pageSize.GetWidth() / 2 - 30, pageSize.GetHeight() - 20)
                    .ShowText("圖書系統")
                    .EndText();

                //Draw footer line
                iText.Kernel.Colors.Color bgColour = new DeviceRgb(0, 0, 0);
                canvas.SetStrokeColor(bgColour)
                    .SetLineWidth(2.2f)
                    .MoveTo(pageSize.GetWidth() / 2 - 30, 20)
                    .LineTo(pageSize.GetWidth() / 2 + 30, 20)
                    .Stroke();

                //Draw page number
                canvas.BeginText()
                    .SetFontAndSize(font, 7)
                    .MoveText(pageSize.GetWidth() / 2 - 7, 10)
                    .ShowText(i.ToString())
                    .ShowText(" of ")
                    .ShowText(n.ToString())
                    .EndText();

                //Draw watermark
                Paragraph p = new Paragraph("極  機  密 \n Confidential").SetFont(font).SetFontSize(60);
                canvas.SaveState();
                PdfExtGState gs1 = new PdfExtGState().SetFillOpacity(0.2f);
                canvas.SetExtGState(gs1);
                document.ShowTextAligned(p, pageSize.GetWidth() / 2, pageSize.GetHeight() / 2, pdfDoc.GetPageNumber(page), TextAlignment.CENTER, VerticalAlignment.MIDDLE, 45);
                canvas.RestoreState();
            }
        }

        //Bitmap to Byte array
        public byte[] BmpToBytes(Bitmap bmp)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            byte[] b = ms.GetBuffer();
            return b;
        }

        public Bitmap GetQrcode(string log, int i_width, int i_height)
        {
            System.Drawing.Bitmap bitmap = null;
            //let string to qr-code
            string strQrCodeContent = log;

            ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter
            {
                Format = ZXing.BarcodeFormat.QR_CODE,
                Options = new ZXing.QrCode.QrCodeEncodingOptions
                {
                    //Create Photo 
                    Height = i_width,
                    Width = i_height,
                    CharacterSet = "UTF-8",

                    //錯誤修正容量
                    //L水平    7%的字碼可被修正
                    //M水平    15%的字碼可被修正
                    //Q水平    25%的字碼可被修正
                    //H水平    30%的字碼可被修正
                    ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H
                }

            };
            //Create Qr-code , use input string
            bitmap = writer.Write(strQrCodeContent);
            return bitmap;
        }

        /// <summary>
        /// 取得pdfQRcode資料
        /// </summary>
        /// <returns></returns>
        private string GetQrcodeData()
        {
            string strQrcodeData = "";
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
                    strQrcodeData = strQrcodeData + Convert.ToString(DBConfig.sqlite_datareader["category_name"]) + "：" + Convert.ToInt32(DBConfig.sqlite_datareader["count"]) + "\n";
                }
                DBConfig.sqlite_datareader.Close();
            }
            return strQrcodeData;
        }

        /// <summary>
        /// 產生QRcode功能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button12_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text == "")
            {
                MessageBox.Show("請輸入欲轉換文字");
            }
            else
            {
                //Use bitmap to storage qr-code
                System.Drawing.Bitmap bitmap = null;
                //let string to qr-code
                string strQrCodeContent = richTextBox1.Text;

                // QR Code產生器
                ZXing.BarcodeWriter writer = new ZXing.BarcodeWriter
                {
                    Format = ZXing.BarcodeFormat.QR_CODE,
                    Options = new ZXing.QrCode.QrCodeEncodingOptions
                    {
                        //Create Photo
                        Height = 200,
                        Width = 200,
                        CharacterSet = "UTF-8",

                        //錯誤修正容量
                        //L水平    7%的字碼可被修正
                        //M水平    15%的字碼可被修正
                        //Q水平    25%的字碼可被修正
                        //H水平    30%的字碼可被修正
                        ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H
                    }

                };
                //Create Qr-code , use input string
                bitmap = writer.Write(strQrCodeContent);
                //Storage bitmpa

                string strDir;
                strDir = Directory.GetCurrentDirectory();
                strDir += "\\temp.jpg";
                bitmap.Save(strDir, System.Drawing.Imaging.ImageFormat.Jpeg);
                //Display to picturebox
                pictureBox2.Image = bitmap;
            }
        }
    }
}
