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

    public partial class frmCustomer : MetroFramework.Forms.MetroForm
    {

        SqlConnection conn = Connection.getConnection();

        public frmCustomer()
        {
            InitializeComponent();
        }

        private void MetroButton2_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                String Code = metroTextBox1.Text;
                String Name = metroTextBox3.Text;
                String TaxID = metroTextBox4.Text;
                String Address = metroTextBox5.Text;
                String Tel = metroTextBox6.Text;
                String Fax = metroTextBox7.Text;
                String Email = metroTextBox8.Text;

                String qry = "insert into tblCustomer values(N'" + Name + "',N'" + TaxID + "',N'" + Address + "',N'" + Tel + "',N'" + Fax + "',N'" + Email + "','" + Code + "')";

                SqlCommand sql = new SqlCommand(qry, conn);
                int i = sql.ExecuteNonQuery();
                if (i > 0)
                {
                    MessageBox.Show(Name + " Has been Added");
                }
                else
                {
                    MessageBox.Show("Failed");
                }

                
                conn.Close();
                show();


            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        void show()
        {
            SqlDataAdapter sda = new SqlDataAdapter("select * from tblCustomer", conn);

            DataTable dt = new DataTable();
            sda.Fill(dt);

            metroGrid1.Rows.Clear();
            foreach (DataRow dr in dt.Rows)
            {
                int n = metroGrid1.Rows.Add();
                metroGrid1.Rows[n].Cells[0].Value = dr[6].ToString();
                metroGrid1.Rows[n].Cells[1].Value = dr[0].ToString();
                metroGrid1.Rows[n].Cells[2].Value = dr[1].ToString();
                metroGrid1.Rows[n].Cells[3].Value = dr[2].ToString();
                metroGrid1.Rows[n].Cells[4].Value = dr[3].ToString();
                metroGrid1.Rows[n].Cells[5].Value = dr[4].ToString();
                metroGrid1.Rows[n].Cells[6].Value = dr[5].ToString();
            }

        }

        private void FrmCustomer_Load(object sender, EventArgs e)
        {
            show();
        }

        private void MetroButton1_Click(object sender, EventArgs e)
        {
            metroTextBox1.Text = "";
            metroTextBox3.Text = "";
            metroTextBox4.Text = "";
            metroTextBox5.Text = "";
            metroTextBox6.Text = "";
            metroTextBox7.Text = "";
            metroTextBox8.Text = "";
        }

        String temp = "";

        private void MetroGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {

                metroTextBox1.Text = metroGrid1.Rows[e.RowIndex].Cells[0].Value.ToString();
                metroTextBox3.Text = metroGrid1.Rows[e.RowIndex].Cells[1].Value.ToString();
                metroTextBox4.Text = metroGrid1.Rows[e.RowIndex].Cells[2].Value.ToString();
                metroTextBox5.Text = metroGrid1.Rows[e.RowIndex].Cells[3].Value.ToString();
                metroTextBox6.Text = metroGrid1.Rows[e.RowIndex].Cells[4].Value.ToString();
                metroTextBox7.Text = metroGrid1.Rows[e.RowIndex].Cells[5].Value.ToString();
                metroTextBox8.Text = metroGrid1.Rows[e.RowIndex].Cells[6].Value.ToString();
                temp = metroTextBox3.Text;
            }
        }

        private void MetroButton3_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                String Code = metroTextBox1.Text;
                String Name = metroTextBox3.Text;
                String TaxID = metroTextBox4.Text;
                String Address = metroTextBox5.Text;
                String Tel = metroTextBox6.Text;
                String Fax = metroTextBox7.Text;
                String Email = metroTextBox8.Text;

                String qry = "update tblCustomer set Name = N'" + Name + "', TaxID = N'" + TaxID + "', Address = N'" + Address + "', Tel = N'" + Tel + "', " +
                    "Fax = N'" + Fax + "', Email = N'" + Email + "', Code = N'" + Code + "' where Name = N'" + temp + "' ";

                SqlCommand sql = new SqlCommand(qry, conn);
                int i = sql.ExecuteNonQuery();
                if (i > 0)
                    MessageBox.Show("Update Success");
                else
                    MessageBox.Show("Update Failed");

                

                conn.Close();
                show();

            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        private void MetroButton4_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                String Code = metroTextBox1.Text;
                String Name = metroTextBox3.Text;
                String TaxID = metroTextBox4.Text;
                String Address = metroTextBox5.Text;
                String Tel = metroTextBox6.Text;
                String Fax = metroTextBox7.Text;
                String Email = metroTextBox8.Text;

                String qry = "delete from tblProduct where Name = N'" + Name + "' ";

                SqlCommand sql = new SqlCommand(qry, conn);
                int i = sql.ExecuteNonQuery();
                if (i > 0)
                    MessageBox.Show("Delete Success");
                else
                    MessageBox.Show("Delete Failed");

                conn.Close();
                show();

                

            }
            catch (Exception)
            {
                conn.Close();
            }
        }
    }
}
