using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class FatKon : Form
    {
        public FatKon()
        {
            InitializeComponent();
        }
        OleDbConnection bag = new OleDbConnection("Provider=Microsoft.Ace.Oledb.12.0;Data Source=FatKon.accdb");
        OleDbDataAdapter da;
        System.Data.DataTable tablo = new System.Data.DataTable();
        OleDbCommandBuilder cb;
        TimeSpan fark;
        double gunfark;

        void Listele()
        {
            tablo.Clear();  
            da = new OleDbDataAdapter("SELECT * From alaybey ", bag);
            da.Fill(tablo);
            dataGridView1.DataSource = tablo;
        }
        void renklendir()
        {

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {

                fark = Convert.ToDateTime(dataGridView1.Rows[i].Cells["sontarih"].Value.ToString()) - Convert.ToDateTime(DateTime.Now.ToShortDateString());
                gunfark = fark.TotalDays;
                bool odeme = Convert.ToBoolean(dataGridView1.Rows[i].Cells["odendi"].Value);
                if (gunfark <= 3 && odeme == false)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
                else if (gunfark > 3 && gunfark < 7 && odeme == false)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                }
                else
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                }
            }
        }
        private void FatKon_Load(object sender, EventArgs e)
        {
 label1.Text = DateTime.Now.ToShortDateString();
            Listele();
            dataGridView1.Columns[0].Visible = false;
            renklendir();
            dataGridView1.Columns[1].HeaderText = "Fatura Cinsi";
            dataGridView1.Columns[2].HeaderText = "Fatura Tarihi";
            dataGridView1.Columns[3].HeaderText = "Son Ödeme Tarihi";
            dataGridView1.Columns[4].HeaderText = "Tutar";
            dataGridView1.Columns[5].HeaderText = "Ödeme Yapıldı";
 
        }

        private void button1_Click(object sender, EventArgs e)
        {
                try
                {
                    cb = new OleDbCommandBuilder(da);
                    da.Update(tablo);
                    MessageBox.Show("Kayıt güncellendi");
                    renklendir();
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }

        private void button2_Click(object sender, EventArgs e)
        {
         
                Excel.Application excel = new Excel.Application();
                excel.Visible = true;
                object Missing = Type.Missing;
                Workbook workbook = excel.Workbooks.Add(Missing);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }
                StartRow++;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {

                        Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                        myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        myRange.Select();


                    }
                }
            }
        }
    }
    

