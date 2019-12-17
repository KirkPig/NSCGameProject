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
    public partial class frmCreditNoteEdit : MetroFramework.Controls.MetroUserControl
    {

        SqlConnection conn = Connection.getConnection();
        String temp = "";

        public frmCreditNoteEdit()
        {
            InitializeComponent();
        }

        private void reset()
        {
            metroGrid1.Rows.Clear();
            metroTextBox1.Text = "";
            metroDateTime1.Value = DateTime.Now;
            metroDateTime2.Value = DateTime.Now;
            metroTextBox3.Text = "";
            metroTextBox4.Text = "";
            metroTextBox5.Text = "";
            metroTextBox6.Text = "";
            metroTextBox7.Text = "";
            metroTextBox8.Text = "";
            metroTextBox9.Text = "";

            metroTextBox10.Text = "0.00";
            metroTextBox2.Text = "0.00";
            metroTextBox11.Text = "0.00";
            metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text.ToString()) * 7 / 100);
            metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) + float.Parse(metroTextBox12.Text));

        }

        private void FrmCreditNoteEdit_Load(object sender, EventArgs e)
        {
            reset();

            metroTextBox10.Text = "0.00";
            metroTextBox2.Text = "0.00";
            metroTextBox11.Text = "0.00";
            metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text.ToString()) * 7 / 100);
            metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) + float.Parse(metroTextBox12.Text));
            metroGrid1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            metroGrid1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            metroDateTime1.Format = DateTimePickerFormat.Custom;
            metroDateTime1.CustomFormat = "dd-MM-yyyy";
            metroDateTime2.Format = DateTimePickerFormat.Custom;
            metroDateTime2.CustomFormat = "dd-MM-yyyy";
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

                String qry = "delete from tblCreditNoteReport where Id = N'" + temp + "' ";
                SqlCommand sql = new SqlCommand(qry, conn);
                sql.ExecuteNonQuery();

                qry = "delete from tblCreditNote where Id = N'" + temp + "' ";
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
                String invoiceID = (metroTextBox9.Text);
                String invoiceDate = (metroDateTime2.Text);
                float valueOld = float.Parse(metroTextBox10.Text);
                float valueReal = float.Parse(metroTextBox2.Text);
                float valueBeforeTax = float.Parse(metroTextBox11.Text);
                float valueTax = float.Parse(metroTextBox12.Text);
                float valueAfterTax = float.Parse(metroTextBox13.Text);

                String str = "insert into tblCreditNoteReport values(N'" + ID + "',N'" + date + "',N'" + customerName + "',N'" + customerTaxID + "',N'" +
                    customerAddress + "',N'" + customerTel + "'" +
                    ",N'" + customerFax + "',N'" + customerEmail + "',N'" + invoiceID + "',N'" + invoiceDate + "'," + valueOld + "," + valueReal + "," +
                    valueBeforeTax + "," + valueTax + "," + valueAfterTax + ")";

                sql = new SqlCommand(str, conn);
                int k = sql.ExecuteNonQuery();
                if (k < 1)
                {
                    chk = false;
                }
                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    int no = int.Parse(metroGrid1.Rows[i].Cells[1].Value.ToString());
                    String desc = metroGrid1.Rows[i].Cells[2].Value.ToString();
                    float amount = float.Parse(metroGrid1.Rows[i].Cells[3].Value.ToString());

                    str = "insert into tblCreditNote values(N'" + ID + "'," + no + ",N'" + desc + "'," + amount + ")";
                    sql = new SqlCommand(str, conn);
                    k = sql.ExecuteNonQuery();
                    if (k < 1)
                    {
                        chk = false;
                    }

                }

                if (chk)
                {
                    MessageBox.Show("Order " + ID + " has been created!");
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

        private void MetroTextBox14_TextChanged(object sender, EventArgs e)
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
                    metroTextBox14.Text = dt.Rows[0][6].ToString();
                }


            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        private void MetroTextBox9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                String str = metroTextBox9.Text;


                SqlDataAdapter sda = new SqlDataAdapter("select * from tblInvoiceReport where Id like '" + str + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                if (dt.Rows.Count == 1)
                {
                    metroTextBox9.Text = dt.Rows[0][0].ToString();
                    metroDateTime2.Value = DateTime.ParseExact(dt.Rows[0][1].ToString(), "dd-MM-yyyy", null);
                    metroTextBox3.Text = dt.Rows[0][2].ToString();
                    metroTextBox4.Text = dt.Rows[0][3].ToString();
                    metroTextBox5.Text = dt.Rows[0][4].ToString();
                    metroTextBox6.Text = dt.Rows[0][5].ToString();
                    metroTextBox7.Text = dt.Rows[0][6].ToString();
                    metroTextBox8.Text = dt.Rows[0][7].ToString();
                    metroTextBox10.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][15].ToString()));

                    sda = new SqlDataAdapter("select * from tblInvoice where Id like '" + str + "%'", conn);

                    dt = new DataTable();
                    sda.Fill(dt);

                    metroGrid1.Rows.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        int n = metroGrid1.Rows.Add();
                        metroGrid1.Rows[n].Cells[0].Value = dr[0].ToString();
                        metroGrid1.Rows[n].Cells[1].Value = dr[1].ToString();
                        metroGrid1.Rows[n].Cells[2].Value = dr[2].ToString();
                        metroGrid1.Rows[n].Cells[3].Value = dr[6].ToString();
                    }

                }

                conn.Close();


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


                SqlDataAdapter sda = new SqlDataAdapter("select * from tblCreditNoteReport where Id like '" + str + "%'", conn);

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
                    metroTextBox9.Text = dt.Rows[0][8].ToString();
                    metroDateTime2.Value = DateTime.ParseExact(dt.Rows[0][9].ToString(), "dd-MM-yyyy", null);

                    metroTextBox10.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][10].ToString()));
                    metroTextBox2.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][11].ToString()));
                    metroTextBox11.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][12].ToString()));
                    metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][13].ToString()));
                    metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][14].ToString()));

                    sda = new SqlDataAdapter("select * from tblCreditNote where Id like '" + str + "%'", conn);

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
                    }

                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        private void MetroTextBox10_Leave(object sender, EventArgs e)
        {
            if (metroTextBox10.Text == "")
            {

            }
            else
            {
                metroTextBox10.Text = String.Format("{0:0.00}", float.Parse(metroTextBox10.Text));
                float A = float.Parse(metroTextBox10.Text) - float.Parse(metroTextBox2.Text);
                metroTextBox11.Text = String.Format("{0:0.00}", A);
                metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) * 7 / 100);
                metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(metroTextBox11.Text) + float.Parse(metroTextBox12.Text));
            }
        }

        private void MetroGrid1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                float a = 0;
                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    if (metroGrid1.Rows[i].Cells[3].Value != null)
                        a = a + float.Parse(metroGrid1.Rows[i].Cells[3].Value.ToString());
                    else
                    {
                        break;
                    }
                }

                metroTextBox2.Text = String.Format("{0:0.00}", a);
            }
            catch (Exception err)
            {
                
            }
        }
    }
}
