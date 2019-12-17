using MetroFramework.Controls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ProjectHouse
{
    public partial class frmQuotationNew : MetroFramework.Controls.MetroUserControl
    {

        SqlConnection conn = Connection.getConnection();
        float A = 0;

        public frmQuotationNew()
        {
            InitializeComponent();
        }

        private void MetroButton5_Click(object sender, EventArgs e)
        {
            if (metroTextBox1.Text == "")
            {
                MessageBox.Show("Please Enter ID Form before add new Product");
            }
            else
            {
                int n = metroGrid1.Rows.Add();
                metroGrid1.Rows[n].Cells[0].Value = metroTextBox1.Text;
                metroGrid1.Rows[n].Cells[1].Value = n + 1;
            }
        }

        int bill = 0;

        private void MetroGrid1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 2 && metroGrid1.Rows[e.RowIndex].Cells[2].Value != null)
                {

                    conn.Open();

                    String str = metroGrid1.Rows[e.RowIndex].Cells[2].Value.ToString();

                    SqlDataAdapter sda = new SqlDataAdapter("select * from tblProduct where Code like '" + str + "%'", conn);

                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    int n = dt.Rows.Count;

                    if (n == 1)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            metroGrid1.Rows[e.RowIndex].Cells[2].Value = dr[1].ToString();
                            metroGrid1.Rows[e.RowIndex].Cells[4].Value = dr[2].ToString();
                            metroGrid1.Rows[e.RowIndex].Cells[5].Value = dr[3].ToString();
                        }
                    }

                    conn.Close();

                }

                //Amount Calculation
                if (metroGrid1.Rows[e.RowIndex].Cells[0].Value != null
                    && metroGrid1.Rows[e.RowIndex].Cells[1].Value != null
                    && metroGrid1.Rows[e.RowIndex].Cells[2].Value != null
                    && metroGrid1.Rows[e.RowIndex].Cells[3].Value != null
                    && metroGrid1.Rows[e.RowIndex].Cells[4].Value != null
                    && metroGrid1.Rows[e.RowIndex].Cells[5].Value != null
                    && metroGrid1.Rows[e.RowIndex].Cells[6].Value != null)
                {
                    //3 * 5 * 6 * .01
                    float a;
                    float quantity = float.Parse(metroGrid1.Rows[e.RowIndex].Cells[3].Value.ToString());
                    float price = float.Parse(metroGrid1.Rows[e.RowIndex].Cells[5].Value.ToString());
                    float discount = float.Parse(metroGrid1.Rows[e.RowIndex].Cells[6].Value.ToString());
                    a = quantity * price * (100 - discount) / 100;
                    metroGrid1.Rows[e.RowIndex].Cells[7].Value = String.Format("{0:0.00}", a);

                    A = 0;
                    for (int i = 0; i < metroGrid1.Rows.Count; i++)
                    {
                        if (metroGrid1.Rows[i].Cells[7].Value != null)
                            A = A + float.Parse(metroGrid1.Rows[i].Cells[7].Value.ToString());
                        else
                        {
                            break;
                        }
                    }

                    metroTextBox11.Text = String.Format("{0:0.00}", A);
                    metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) * 7 / 100);
                    metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) + float.Parse(metroTextBox12.Text));

                }
            }
            catch (Exception)
            {
                conn.Close();
            }


            

        }

        private void FrmQuotationNew_Load(object sender, EventArgs e)
        {
            reset();

            metroTextBox11.Text = "0.00";
            metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text.ToString()) * 7 / 100);
            metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) + float.Parse(metroTextBox12.Text));
            metroGrid1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            metroDateTime1.Format = DateTimePickerFormat.Custom;
            metroDateTime1.CustomFormat = "dd-MM-yyyy";

        } 

        private void MetroButton2_Click(object sender, EventArgs e)
        {
            reset();
        }

        private void MetroButton3_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                bool chk = true;
                String ID = metroTextBox1.Text;
                String date = metroDateTime1.Text;
                String customerName = metroTextBox3.Text;
                String customerTaxID = metroTextBox4.Text;
                String customerAddress = metroTextBox5.Text;
                String customerTel = metroTextBox6.Text;
                String customerFax = metroTextBox7.Text;
                String customerEmail = metroTextBox8.Text;
                String ATTN = metroTextBox9.Text;
                int CR = int.Parse(metroTextBox10.Text);
                float valueBeforeTax = float.Parse(metroTextBox11.Text);
                float valueTax = float.Parse(metroTextBox12.Text);
                float valueAfterTax = float.Parse(metroTextBox13.Text);

                String str = "insert into tblQuotationReport values(N'" + ID + "',N'" + date + "',N'" + customerName + "',N'" + customerTaxID + "',N'" + customerAddress + "',N'" + customerTel + "'" +
                    ",N'" + customerFax + "',N'" + customerEmail + "',N'" + ATTN + "'," + CR + "," + valueBeforeTax + "," + valueTax + "," + valueAfterTax + ")";

                SqlCommand sql = new SqlCommand(str, conn);
                int k = sql.ExecuteNonQuery();
                if (k<1)
                {
                    chk = false;
                }
                for (int i = 0;i < metroGrid1.Rows.Count;i++)
                {
                    int no = int.Parse(metroGrid1.Rows[i].Cells[1].Value.ToString());
                    String desc = metroGrid1.Rows[i].Cells[2].Value.ToString();
                    int quan = int.Parse(metroGrid1.Rows[i].Cells[3].Value.ToString());
                    String unit = metroGrid1.Rows[i].Cells[4].Value.ToString();
                    float price = float.Parse(metroGrid1.Rows[i].Cells[5].Value.ToString());
                    float discount = float.Parse(metroGrid1.Rows[i].Cells[6].Value.ToString());
                    float amount = float.Parse(metroGrid1.Rows[i].Cells[7].Value.ToString());

                    str = "insert into tblQuotation values(N'" + ID + "'," + no + ",N'" + desc + "'," + quan + ",N'" + unit + "'," + price + "," + discount + "," + amount + ")";
                    sql = new SqlCommand(str, conn);
                    k = sql.ExecuteNonQuery();
                    if (k < 1)
                    {
                        chk = false;
                    }

                }

                //GenerateID Save
                String qr = "update tblGenerateID set NumYear = " + metroDateTime1.Value.Year + 
                    ", NumMonth = " + metroDateTime1.Value.Month + ", NumBill = " + bill + " where Name = 'QY'";

                SqlCommand sr = new SqlCommand(qr, conn);
                int j = sr.ExecuteNonQuery();

                if (chk)
                {
                    MessageBox.Show("Quotation " + ID + " has been created!");
                }
                else
                {
                    MessageBox.Show("Failed..");
                }

                conn.Close();

                printPreviewDialog1.ShowDialog();
                DialogResult dialogResult = printDialog1.ShowDialog();
                if (dialogResult == DialogResult.Cancel) { }
                else printDocument1.Print();

                reset();


            }
            catch (Exception)
            {
                conn.Close();
            }
            
        }

        private void reset()
        {
            metroTextBox1.Text = "";
            metroDateTime1.Value = DateTime.Now;
            metroTextBox3.Text = "";
            metroTextBox4.Text = "";
            metroTextBox5.Text = "";
            metroTextBox6.Text = "";
            metroTextBox7.Text = "";
            metroTextBox8.Text = "";
            metroTextBox9.Text = "";
            metroTextBox10.Text = "";

            metroTextBox11.Text = "0.00";
            metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text.ToString()) * 7 / 100);
            metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) + float.Parse(metroTextBox12.Text));
            metroGrid1.Rows.Clear();

            metroTextBox1.Text = generateId(metroDateTime1.Value);
            
        }

        private string generateId(DateTime time)
        {
            //GenerateID
            try
            {
                conn.Open();
                string gid = "";
                int k = 1;

                while (true)
                {
                    
                    gid = String.Format("QY{0:yy}0{0:MM}{1:000}", time, k);
                    String str = "select * from tblQuotationReport where Id like N'"+gid+"'";
                    SqlDataAdapter sda = new SqlDataAdapter(str, conn);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    if (dt.Rows.Count == 0)
                    {
                        break;
                    }
                    k = k + 1;
                }

                conn.Close();

                return gid;

            }
            catch (Exception)
            {
                conn.Close();

                return "";
            }
        }

        private void MetroTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintDoc printDoc = new PrintDoc();
            printDoc.printQuotation(metroTextBox1.Text, metroDateTime1.Text, metroTextBox3.Text, metroTextBox4.Text, metroTextBox5.Text,
                metroTextBox6.Text, metroTextBox7.Text, metroTextBox8.Text, metroTextBox9.Text, metroTextBox10.Text, metroTextBox11.Text,
                metroTextBox12.Text, metroTextBox13.Text, sender, e);
        }

        private void MetroTextBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                String str = metroTextBox2.Text;
                SqlDataAdapter sda = new SqlDataAdapter("select * from tblCustomer where Code like N'" + str + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                conn.Close();
                if (dt.Rows.Count == 1)
                {
                    metroTextBox3.Text = dt.Rows[0][0].ToString();
                    metroTextBox4.Text = dt.Rows[0][1].ToString();
                    metroTextBox5.Text = dt.Rows[0][2].ToString();
                    metroTextBox6.Text = dt.Rows[0][3].ToString();
                    metroTextBox7.Text = dt.Rows[0][4].ToString();
                    metroTextBox8.Text = dt.Rows[0][5].ToString();
                    metroTextBox2.Text = dt.Rows[0][6].ToString();
                }

                
            }
            catch (Exception)
            {
                conn.Close();
            }

        }

        private void MetroButton1_Click(object sender, EventArgs e)
        {

            foreach (DataGridViewRow item in metroGrid1.SelectedRows)
            {
                metroGrid1.Rows.RemoveAt(item.Index);
            }

            int i = 1;
            foreach (DataGridViewRow item in metroGrid1.Rows)
            {
                item.Cells[1].Value = i;
                i++;
            }

        }

        private void MetroDateTime1_ValueChanged(object sender, EventArgs e)
        {
            metroTextBox1.Text = generateId(metroDateTime1.Value);
        }
    }
}
