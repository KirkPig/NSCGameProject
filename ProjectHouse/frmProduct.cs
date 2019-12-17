using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ProjectHouse
{
    public partial class frmProduct : MetroFramework.Forms.MetroForm
    {

        SqlConnection conn = Connection.getConnection();
        public frmProduct()
        {
            InitializeComponent();
            show();
        }

        private void MetroButton1_Click(object sender, EventArgs e)
        {
            metroTextBox1.Text = "";
            metroTextBox2.Text = "";
            metroTextBox3.Text = "";
            metroTextBox4.Text = "";
        }

        private void MetroButton2_Click(object sender, EventArgs e)
        {

            try
            {

                conn.Open();

                String Code = metroTextBox1.Text;
                String Name = metroTextBox1.Text + ": " + metroTextBox2.Text;
                String Unit = metroTextBox3.Text;
                float Price = float.Parse(metroTextBox4.Text);

                bool B = true;

                for (int i = 0;i<Code.Length;i++)
                {
                    if (Code[i]==':')
                    {
                        B = false;
                    }
                }


                if (B)
                {
                    String qry = "insert into tblProduct values(N'" + Code + "',N'" + Name + "',N'" + Unit + "'," + Price + ")";

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
                else
                {
                    MessageBox.Show("Product Code Can't Have ':'");
                    conn.Close();
                }
                


            }catch(Exception)
            {
                conn.Close();
            }

        }

        void show()
        {
            try
            {

                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblProduct", conn);

                DataTable dt = new DataTable();
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

                conn.Close();

            }
            catch (Exception)
            {
                conn.Close();
            }

        }

        private void MetroGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {

                metroTextBox1.Text = metroGrid1.Rows[e.RowIndex].Cells[0].Value.ToString();
                String str = metroGrid1.Rows[e.RowIndex].Cells[1].Value.ToString();
                metroTextBox3.Text = metroGrid1.Rows[e.RowIndex].Cells[2].Value.ToString();
                metroTextBox4.Text = metroGrid1.Rows[e.RowIndex].Cells[3].Value.ToString();

                bool A = false;
                String str2 = "";
                for (int i = 0;i<str.Length;i++)
                {
                    if (A)
                    {
                        str2 = str2 + str[i].ToString();
                    }
                    if (str[i] == ':')  
                    {
                        i++;
                        A = true;
                    }
                }

                metroTextBox2.Text = str2;
                temp = metroTextBox1.Text;

            }
        }

        String temp = "";

        private void MetroButton3_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                String Code = metroTextBox1.Text;
                String Name = metroTextBox1.Text + ": " + metroTextBox2.Text;
                String Unit = metroTextBox3.Text;
                float Price = float.Parse(metroTextBox4.Text);

                String qry = "update tblProduct set Code = N'" + Code + "', Name = N'" + Name + "', Unit = N'" + Unit + "', Price = "+ Price +" where Code = N'" + temp + "' ";

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
                String Name = metroTextBox1.Text + ": " + metroTextBox2.Text;
                String Unit = metroTextBox3.Text;
                float Price = float.Parse(metroTextBox4.Text);

                String qry = "delete from tblProduct where Code = N'" + Code + "' ";

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

        private void MetroButton5_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                String str = metroTextBox9.Text;
                SqlDataAdapter sda = new SqlDataAdapter("select * from tblProduct where Code like N'" + str + "%'", conn);

                DataTable dt = new DataTable();
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

                conn.Close();

            }
            catch (Exception)
            {
                conn.Close();
            }
        }
    }
}
