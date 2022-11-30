# Csharp_Report

資料庫名稱：book.db <br>
資料表名稱：book_data、category_data、member

for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                  this.label8.Text = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                  this.textBox1.Text = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                  this.textBox2.Text = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                  this.textBox3.Text = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                  this.comboBox1.Text = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                  this.textBox5.Text = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                  this.textBox6.Text = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);
                  MessageBox.Show("書名：" + this.textBox1.Text +"\n作者：" + this.textBox2.Text + "\n出版社：" + this.textBox3.Text + "\n類型：" + this.comboBox1.Text + "\n類別：" + this.textBox5.Text + "\nISBN：" + this.textBox6.Text, "序號 " + this.label8.Text);
                  break;
            }
