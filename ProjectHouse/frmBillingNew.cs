using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ProjectHouse
{
    public partial class frmBillingNew : MetroFramework.Controls.MetroUserControl
    {

        SqlConnection conn = Connection.getConnection();
        float A = 0;

        public frmBillingNew()
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
                metroGrid1.Rows[n].Cells[6].Value = "";
                metroGrid1.Rows[n].Cells[1].Value = n + 1;
            }
        }

        int bill = 0;

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

        private void FrmBillingNew_Load(object sender, EventArgs e)
        {
            reset();

            metroTextBox13.Text = "0.00";
            metroDateTime1.Format = DateTimePickerFormat.Custom;
            metroDateTime2.Format = DateTimePickerFormat.Custom;
            metroDateTime1.CustomFormat = "dd-MM-yyyy";
            metroDateTime2.CustomFormat = "dd-MM-yyyy";
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

                    gid = String.Format("RB{0:yy}0{0:MM}{1:000}", time, k);
                    String str = "select * from tblBillingReport where Id like N'" + gid + "'";
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
                String dateBilling = metroDateTime2.Text;
                String SavedBy = metroTextBox9.Text;
                String CustomerCode = metroTextBox2.Text;
                String PS = metroTextBox10.Text;
                float valueAfterTax = float.Parse(metroTextBox13.Text);

                String str = "insert into tblBillingReport values(N'" + ID + "',N'" + date + "',N'" + customerName + "',N'" + customerTaxID + "',N'" + customerAddress + "',N'" + customerTel + "'" +
                    ",N'" + customerFax + "',N'" + customerEmail + "',N'" + dateBilling + "',N'" + SavedBy + "',N'" + CustomerCode + "',N'" + PS + "'," + valueAfterTax + ")";

                SqlCommand sql = new SqlCommand(str, conn);
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

                //GenerateID Save
                String qr = "update tblGenerateID set NumYear = " + metroDateTime1.Value.Year +
                    ", NumMonth = " + metroDateTime1.Value.Month + ", NumBill = " + bill + " where Name = 'RB'";

                SqlCommand sr = new SqlCommand(qr, conn);
                int j = sr.ExecuteNonQuery();

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

        int dpi = 96;
        float mmpi = 25.4f;

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintDoc printDoc = new PrintDoc();
            printDoc.printBilling(metroTextBox1.Text, metroDateTime1.Text, metroTextBox2.Text, metroTextBox3.Text, metroTextBox4.Text, metroTextBox5.Text,
                metroTextBox6.Text, metroTextBox7.Text, metroTextBox8.Text, metroTextBox9.Text, metroDateTime2.Text,
                metroTextBox10.Text, metroTextBox13.Text, sender, e);

        }

        private void MetroTextBox3_TextChanged(object sender, EventArgs e)
        {

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

        private void MetroButton6_Click(object sender, EventArgs e)
        {
            metroGrid1.Rows.Clear();
        }

        private void MetroButton4_Click(object sender, EventArgs e)
        {
            try
            {

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblInvoiceReport", conn);
            }
            catch (Exception)
            {

            }
           
        }

        private void MetroDateTime1_ValueChanged(object sender, EventArgs e)
        {
            metroTextBox1.Text = generateId(metroDateTime1.Value);
        }
    }
}
