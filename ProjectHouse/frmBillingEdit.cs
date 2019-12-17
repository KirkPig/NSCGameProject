using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectHouse
{
    public partial class frmBillingEdit : MetroFramework.Controls.MetroUserControl
    {

        SqlConnection conn = Connection.getConnection();
        float A = 0;
        String temp = "";

        public frmBillingEdit()
        {
            InitializeComponent();
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

            metroTextBox13.Text = "0.00";
            metroGrid1.Rows.Clear();
        }

        private void FrmBillingEdit_Load(object sender, EventArgs e)
        {
            reset();

            metroTextBox13.Text = "0.00";
            metroDateTime1.Format = DateTimePickerFormat.Custom;
            metroDateTime2.Format = DateTimePickerFormat.Custom;
            metroDateTime1.CustomFormat = "dd-MM-yyyy";
            metroDateTime2.CustomFormat = "dd-MM-yyyy";
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
                metroGrid1.Rows[n].Cells[6].Value = "";
                metroGrid1.Rows[n].Cells[1].Value = n + 1;
            }
        }

        private void MetroGrid1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                conn.Open();

                if (e.ColumnIndex == 2 && metroGrid1.Rows[e.RowIndex].Cells[2].Value != null)
                {
                    String str = metroGrid1.Rows[e.RowIndex].Cells[2].Value.ToString();

                    SqlDataAdapter sda = new SqlDataAdapter("select * from tblInvoiceReport where Id like '" + str + "%'", conn);

                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    int n = dt.Rows.Count;

                    if (n == 1)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            metroGrid1.Rows[e.RowIndex].Cells[2].Value = dr[0].ToString();
                            metroGrid1.Rows[e.RowIndex].Cells[3].Value = dr[1].ToString();
                            metroGrid1.Rows[e.RowIndex].Cells[4].Value = dr[11].ToString();
                            metroGrid1.Rows[e.RowIndex].Cells[5].Value = String.Format("{0:0.00}", float.Parse(dr[15].ToString()));
                        }
                    }
                    A = 0;
                    for (int i = 0; i < metroGrid1.Rows.Count; i++)
                    {
                        if (metroGrid1.Rows[i].Cells[5].Value != null)
                            A = A + float.Parse(metroGrid1.Rows[i].Cells[5].Value.ToString());
                        else
                        {
                            break;
                        }
                    }

                    metroTextBox13.Text = String.Format("{0:0.00}", A);

                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintDoc printDoc = new PrintDoc();
            printDoc.printBilling(metroTextBox1.Text, metroDateTime1.Text, metroTextBox2.Text, metroTextBox3.Text, metroTextBox4.Text, metroTextBox5.Text,
                metroTextBox6.Text, metroTextBox7.Text, metroTextBox8.Text, metroTextBox9.Text, metroDateTime2.Text,
                metroTextBox10.Text, metroTextBox13.Text, sender, e);
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

                //Delete

                String qry = "delete from tblBillingReport where Id = N'" + temp + "' ";
                SqlCommand sql = new SqlCommand(qry, conn);
                sql.ExecuteNonQuery();

                qry = "delete from tblBilling where Id = N'" + temp + "' ";
                sql = new SqlCommand(qry, conn);
                sql.ExecuteNonQuery();

                //Upload


                String date = metroDateTime1.Text;
                String customerName = metroTextBox3.Text;
                String customerTaxID = metroTextBox4.Text;
                String customerAddress = metroTextBox5.Text;
                String customerTel = metroTextBox6.Text;
                String customerFax = metroTextBox7.Text;
                String customerEmail = metroTextBox8.Text;
                String dateBilling = metroDateTime2.Text;
                String SavedBy = metroTextBox9.Text;
                String CustomerCode = metroTextBox2.Text;
                String PS = metroTextBox10.Text;
                float valueAfterTax = float.Parse(metroTextBox13.Text);

                String str = "insert into tblBillingReport values(N'" + ID + "',N'" + date + "',N'" + customerName + "',N'" + customerTaxID + "',N'" + customerAddress + "',N'" + customerTel + "'" +
                    ",N'" + customerFax + "',N'" + customerEmail + "',N'" + dateBilling + "',N'" + SavedBy + "',N'" + CustomerCode + "',N'" + PS + "'," + valueAfterTax + ")";

                sql = new SqlCommand(str, conn);
                int k = sql.ExecuteNonQuery();
                if (k < 1)
                {
                    chk = false;
                }
                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    int no = int.Parse(metroGrid1.Rows[i].Cells[1].Value.ToString());
                    String InvoiceID = metroGrid1.Rows[i].Cells[2].Value.ToString();
                    String Date = metroGrid1.Rows[i].Cells[3].Value.ToString();
                    String DateDue = metroGrid1.Rows[i].Cells[4].Value.ToString();
                    float amount = float.Parse(metroGrid1.Rows[i].Cells[5].Value.ToString());
                    String ps = metroGrid1.Rows[i].Cells[6].Value.ToString();

                    str = "insert into tblBilling values(N'" + ID + "'," + no + ",N'" + InvoiceID + "',N'" + Date + "',N'" + DateDue + "'," + amount + ",N'" + ps + "')";
                    sql = new SqlCommand(str, conn);
                    k = sql.ExecuteNonQuery();
                    if (k < 1)
                    {
                        chk = false;
                    }

                }

                if (chk)
                {
                    MessageBox.Show(ID + " has been created!");
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

        private void MetroTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                String str = metroTextBox1.Text;


                SqlDataAdapter sda = new SqlDataAdapter("select * from tblBillingReport where Id like '" + str + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                if (dt.Rows.Count == 1)
                {
                    metroTextBox1.Text = dt.Rows[0][0].ToString();
                    temp = dt.Rows[0][0].ToString();
                    metroDateTime1.Value = DateTime.ParseExact(dt.Rows[0][1].ToString(), "dd-MM-yyyy", null);
                    metroTextBox3.Text = dt.Rows[0][2].ToString();
                    metroTextBox4.Text = dt.Rows[0][3].ToString();
                    metroTextBox5.Text = dt.Rows[0][4].ToString();
                    metroTextBox6.Text = dt.Rows[0][5].ToString();
                    metroTextBox7.Text = dt.Rows[0][6].ToString();
                    metroTextBox8.Text = dt.Rows[0][7].ToString();
                    
                    metroDateTime2.Value = DateTime.ParseExact(dt.Rows[0][8].ToString(), "dd-MM-yyyy", null);
                    metroTextBox9.Text = dt.Rows[0][9].ToString();
                    metroTextBox2.Text = dt.Rows[0][10].ToString();
                    metroTextBox10.Text = dt.Rows[0][11].ToString();
                    metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][12].ToString()));

                    sda = new SqlDataAdapter("select * from tblBilling where Id like '" + str + "%'", conn);

                    dt = new DataTable();
                    sda.Fill(dt);

                    metroGrid1.Rows.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        int n = metroGrid1.Rows.Add();
                        metroGrid1.Rows[n].Cells[0].Value = dr[0].ToString();
                        metroGrid1.Rows[n].Cells[1].Value = dr[1].ToString();
                        metroGrid1.Rows[n].Cells[2].Value = dr[2].ToString();
                        metroGrid1.Rows[n].Cells[3].Value = dr[3].ToString();
                        metroGrid1.Rows[n].Cells[4].Value = dr[4].ToString();
                        metroGrid1.Rows[n].Cells[5].Value = dr[5].ToString();
                        metroGrid1.Rows[n].Cells[6].Value = dr[6].ToString();
                    }

                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }
        }
    }
}
