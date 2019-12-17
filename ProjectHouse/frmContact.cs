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
    public partial class frmContact : Form
    {

        SqlConnection conn = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=\"C:\\\\Users\\Kirk Pig\\source\\repos\\ProjectHouse\\ProjectHouse\\dtbHouse.mdf\";Integrated Security=True");

        public frmContact()
        {
            InitializeComponent();
            show();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            try
            {

                conn.Open();

                int id = Int32.Parse(textBox1.Text);
                String name = textBox2.Text;
                String phone = textBox3.Text;

                String qry = "insert into tblContact values("+id+",'"+name+"','"+phone+"')";

                SqlCommand sql = new SqlCommand(qry,conn);
                int i  =sql.ExecuteNonQuery();
                if (i > 0)
                    MessageBox.Show("Success");
                else
                    MessageBox.Show("Failed");

                show();

                conn.Close();

            }
            catch (Exception)
            {
                conn.Close();
            }

        }

        void show()
        {
            SqlDataAdapter sda = new SqlDataAdapter("select * from tblContact",conn);

            DataTable dt = new DataTable();
            sda.Fill(dt);

            dataGridView1.Rows.Clear();
            foreach (DataRow dr in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = dr[0].ToString();
                dataGridView1.Rows[n].Cells[1].Value = dr[1].ToString();
                dataGridView1.Rows[n].Cells[2].Value = dr[2].ToString();
            }

        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {

                textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();

            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {

            try
            {

                conn.Open();

                int id = Int32.Parse(textBox1.Text);
                String name = textBox2.Text;
                String phone = textBox3.Text;

                String qry = "update tblContact set Name = '" + name + "', Phone = '" + phone + "' where Id = " + id + " ";

                SqlCommand sql = new SqlCommand(qry, conn);
                int i = sql.ExecuteNonQuery();
                if (i > 0)
                    MessageBox.Show("Update Success");
                else
                    MessageBox.Show("Update Failed");

                show();

                conn.Close();

            }
            catch (Exception)
            {
                conn.Close();
            }

        }

        private void Button4_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                int id = Int32.Parse(textBox1.Text);
                String name = textBox2.Text;
                String phone = textBox3.Text;

                String qry = "delete from tblContact where Id = " + id + " ";

                SqlCommand sql = new SqlCommand(qry, conn);
                int i = sql.ExecuteNonQuery();
                if (i > 0)
                    MessageBox.Show("Delete Success");
                else
                    MessageBox.Show("Delete Failed");

                show();

                conn.Close();

            }
            catch (Exception)
            {
                conn.Close();
            }
        }
    }
}
