using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CsharpReport
{
    public partial class editForm : Form
    {
        public editForm()
        {
            InitializeComponent();
        }

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
    }
}
