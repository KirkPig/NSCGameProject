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
    public partial class frmBillingFast : MetroFramework.Controls.MetroUserControl
    {
        public frmBillingFast()
        {
            InitializeComponent();
        }

        SqlConnection conn = Connection.getConnection();

        private void FrmBillingFast_Load(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblCustomer", conn);
                DataTable dt = new DataTable();
                sda.Fill(dt);

                foreach (DataRow dr in dt.Rows)
                {
                    int n = metroComboBox1.Items.Add(dr[0].ToString());

                }

                conn.Close();

                metroComboBox1.SelectedItem = metroComboBox1.Items[0];
                metroDateTime1.Value = DateTime.Now;

                metroDateTime1.Format = DateTimePickerFormat.Custom;
                metroDateTime2.Format = DateTimePickerFormat.Custom;
                metroDateTime1.CustomFormat = "MM-yyyy";
                metroDateTime2.CustomFormat = "dd-MM-yyyy";


            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        private void MetroButton2_Click(object sender, EventArgs e)
        {
            metroComboBox1.SelectedItem = metroComboBox1.Items[0];
            metroDateTime1.Value = DateTime.Now;
            metroDateTime2.Value = DateTime.Now;
            metroTextBox9.Text = "";
            metroTextBox10.Text = "";
            metroGrid1.Rows.Clear();
        }

        private void MetroButton1_Click(object sender, EventArgs e)
        {
            try
            {
                String Month = metroDateTime1.Text;
                String CustomerName = metroComboBox1.SelectedItem.ToString();

                conn.Open();

                String str = "select * from tblInvoiceReport where (Date like '%" + Month + "%' AND Customer like N'%" + CustomerName + "%') order by Id ASC, Date ASC , Customer ASC";
                SqlDataAdapter sda = new SqlDataAdapter(str, conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                metroGrid1.Rows.Clear();
                foreach (DataRow dr in dt.Rows)
                {
                    int n = metroGrid1.Rows.Add();
                    metroGrid1.Rows[n].Cells[0].Value = true;
                    metroGrid1.Rows[n].Cells[1].Value = dr[0].ToString();
                    metroGrid1.Rows[n].Cells[2].Value = dr[1].ToString();
                    metroGrid1.Rows[n].Cells[3].Value = dr[11].ToString();
                    metroGrid1.Rows[n].Cells[4].Value = dr[15].ToString();
                    metroGrid1.Rows[n].Cells[5].Value = " ";
                }

                conn.Close();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString());
                conn.Close();
            }
        }

        int bill;
        String ID;
        DateTime dateTime;
        String date;
        String customerName;

        String customerTaxID;
        String customerAddress;
        String customerTel;
        String customerFax;
        String customerEmail;

        String dateBilling;
        String SavedBy;
        String CustomerCode;
        String Ps;

        float valueAfterTax;


        private void MetroButton3_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                dateTime = DateTime.Now;
                int ku = 1;

                while (true)
                {

                    ID = String.Format("RB{0:yy}0{0:MM}{1:000}", dateTime, ku);
                    String str1 = "select * from tblBillingReport where Id like N'" + ID + "'";
                    SqlDataAdapter sda1 = new SqlDataAdapter(str1, conn);
                    DataTable dt1 = new DataTable();
                    sda1.Fill(dt1);

                    if (dt1.Rows.Count == 0)
                    {
                        break;
                    }
                    ku = ku + 1;
                }

                bool chk = true;
                date = String.Format("{0:dd}-{0:MM}-{0:yyyy}", dateTime);
                customerName = metroComboBox1.SelectedItem.ToString();


                String str = "select * from tblCustomer where Name like N'"+ customerName +"'";
                SqlDataAdapter sda = new SqlDataAdapter(str, conn);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                customerTaxID = dt.Rows[0][1].ToString();
                customerAddress = dt.Rows[0][2].ToString();
                customerTel = dt.Rows[0][3].ToString();
                customerFax = dt.Rows[0][4].ToString();
                customerEmail = dt.Rows[0][5].ToString();

                dateBilling = metroDateTime2.Text;
                SavedBy = metroTextBox9.Text;
                CustomerCode = dt.Rows[0][6].ToString();
                Ps = metroTextBox10.Text;
                valueAfterTax = 0f;
                SqlCommand sql;
                int k;
                int b = 1;

                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    
                    if ((bool)metroGrid1.Rows[i].Cells[0].Value == false)
                    {
                        continue;
                    }
                    int no = b;
                    b++;
                    String InvoiceID = metroGrid1.Rows[i].Cells[1].Value.ToString();
                    String Date = metroGrid1.Rows[i].Cells[2].Value.ToString();
                    String DateDue = metroGrid1.Rows[i].Cells[3].Value.ToString();
                    float amount = float.Parse(metroGrid1.Rows[i].Cells[4].Value.ToString());
                    String ps = metroGrid1.Rows[i].Cells[5].Value.ToString();
                    valueAfterTax = valueAfterTax + amount;

                    str = "insert into tblBilling values(N'" + ID + "'," + no + ",N'" + InvoiceID + "',N'" + Date + "',N'" + DateDue + "'," + amount + ",N'" + ps + "')";
                    sql = new SqlCommand(str, conn);
                    k = sql.ExecuteNonQuery();
                    if (k < 1)
                    {
                        chk = false;
                    }

                }

                str = "insert into tblBillingReport values(N'" + ID + "',N'" + date + "',N'" + customerName + "',N'" + customerTaxID + "',N'" + customerAddress + "',N'" + customerTel + "'" +
                    ",N'" + customerFax + "',N'" + customerEmail + "',N'" + dateBilling + "',N'" + SavedBy + "',N'" + CustomerCode + "',N'" + Ps + "'," + valueAfterTax + ")";

                sql = new SqlCommand(str, conn);
                k = sql.ExecuteNonQuery();
                if (k < 1)
                {
                    chk = false;
                }

                //GenerateID Save
                String qr = "update tblGenerateID set NumYear = " + dateTime.Year +
                    ", NumMonth = " + dateTime.Month + ", NumBill = " + bill + " where Name = 'RB'";

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
            }
            catch (Exception err)
            {

                MessageBox.Show(err.ToString());
                conn.Close();
            }
        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintDoc printDoc = new PrintDoc();
            printDoc.printBilling(ID, date, CustomerCode, customerName, customerTaxID, customerAddress,
                customerTel,customerFax, customerEmail, metroTextBox9.Text, metroDateTime2.Text,
                metroTextBox10.Text, String.Format("{0:0,0.00}",valueAfterTax), sender, e);
        }
    }
}
