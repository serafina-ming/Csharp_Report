using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;

namespace CsharpReport
{
    public partial class exportChartForm : Form
    {
        Chart chart1;
        Chart chart2;
        Chart chart3;

        public exportChartForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 取得圖表值
        /// </summary>
        public object[] setValue
        {
            set
            {
                chart1 = (Chart)value[0];
                chart2 = (Chart)value[1];
                chart3 = (Chart)value[2];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //取得欲匯出圖表樣式
            string chartType = GetChartType();

            SaveFileDialog save = new SaveFileDialog();
            save.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            save.FileName = "書籍類型統計圖";
            save.Filter = "*.jpg|*.jpg";
            if (save.ShowDialog() != DialogResult.OK) return;

            //匯出對印樣式圖表
            if (chartType == "bar")
            {
                chart1.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            else if (chartType == "pie")
            {
                chart2.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            else if (chartType == "line")
            {
                chart3.SaveImage(save.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            else
            {
                MessageBox.Show("系統錯誤");
            }
        }

        /// <summary>
        /// 判斷欲匯出圖表樣式
        /// </summary>
        /// <returns>圖表樣式名稱</returns>
        public string GetChartType()
        {
            if (radioButton1.Checked == true)
            {
                return "bar";
            }
            else if (radioButton2.Checked == true)
            {
                return "pie";
            }
            else if (radioButton3.Checked == true)
            {
                return "line";
            }
            else
            {
                return "wrong";
            }
        }
    }
}
