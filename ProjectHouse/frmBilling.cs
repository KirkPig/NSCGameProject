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
    public partial class frmBilling : MetroFramework.Controls.MetroUserControl
    {
        public frmBilling()
        {
            InitializeComponent();
        }

        private void FrmBilling_Load(object sender, EventArgs e)
        {
            try
            {

                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblBillingReport", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                metroGrid1.Rows.Clear();
                foreach (DataRow dr in dt.Rows)
                {
                    int n = metroGrid1.Rows.Add();
                    metroGrid1.Rows[n].Cells[0].Value = dr[0].ToString();
                    metroGrid1.Rows[n].Cells[1].Value = dr[1].ToString();
                    metroGrid1.Rows[n].Cells[2].Value = dr[2].ToString();
                    metroGrid1.Rows[n].Cells[3].Value = dr[12].ToString();
                }

                sda = new SqlDataAdapter("select * from tblCustomer", conn);
                dt = new DataTable();
                sda.Fill(dt);

                foreach (DataRow dr in dt.Rows)
                {
                    int n = metroComboBox1.Items.Add(dr[0].ToString());

                }

                conn.Close();

                metroComboBox1.SelectedItem = metroComboBox1.Items[0];
                metroDateTime1.Value = DateTime.Now;

                metroDateTime1.Format = DateTimePickerFormat.Custom;
                metroDateTime1.CustomFormat = "MM-yyyy";
            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        void show()
        {
            try
            {
                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblBillingReport", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                metroGrid1.Rows.Clear();
                foreach (DataRow dr in dt.Rows)
                {
                    int n = metroGrid1.Rows.Add();
                    metroGrid1.Rows[n].Cells[0].Value = dr[0].ToString();
                    metroGrid1.Rows[n].Cells[1].Value = dr[1].ToString();
                    metroGrid1.Rows[n].Cells[2].Value = dr[2].ToString();
                    metroGrid1.Rows[n].Cells[3].Value = dr[12].ToString();
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }

        }

        SqlConnection conn = Connection.getConnection();

        private void MetroButton2_Click(object sender, System.EventArgs e)
        {

            show();

            metroComboBox1.SelectedItem = metroComboBox1.Items[0];
            metroDateTime1.Value = DateTime.Now;
        }

        private void MetroButton1_Click(object sender, System.EventArgs e)
        {
            try
            {
                String Month = metroDateTime1.Text;
                String CustomerName = metroComboBox1.SelectedItem.ToString();

                conn.Open();

                String str = "select * from tblBillingReport where (Date like '%" + Month + "%' AND Customer like N'%" + CustomerName + "%')";
                SqlDataAdapter sda = new SqlDataAdapter(str, conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                metroGrid1.Rows.Clear();
                foreach (DataRow dr in dt.Rows)
                {
                    int n = metroGrid1.Rows.Add();
                    metroGrid1.Rows[n].Cells[0].Value = dr[0].ToString();
                    metroGrid1.Rows[n].Cells[1].Value = dr[1].ToString();
                    metroGrid1.Rows[n].Cells[2].Value = dr[2].ToString();
                    metroGrid1.Rows[n].Cells[3].Value = dr[12].ToString();
                }

                conn.Close();

            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        int dpi = 96;
        float mmpi = 25.4f;

        int currentPage = 0;
        int Page = 1;
        float sum1 = 0.0f;
        float sum2 = 0.0f;
        float sum3 = 0.0f;

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                String Month = metroDateTime1.Text;
                String CustomerName = metroComboBox1.SelectedItem.ToString();

                conn.Open();

                String str = "select * from tblBillingReport where (Date like '%" + Month + "%' AND Customer like N'%" + CustomerName + "%') order by Date";
                SqlDataAdapter sda = new SqlDataAdapter(str, conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                if (dt.Rows.Count % 24 == 0 && dt.Rows.Count != 0) Page = (dt.Rows.Count) / 24;
                else Page = ((dt.Rows.Count) / 24) + 1;

                String font = "TH Sarabun New";
                StringFormat strformat = new StringFormat();
                strformat.Alignment = StringAlignment.Center;
                strformat.LineAlignment = StringAlignment.Center;

                //Header
                Rectangle rect = new Rectangle((int)(12.8 / mmpi * dpi), (int)(10 / mmpi * dpi), (int)(197.2 / mmpi * dpi), (int)(15.6 / mmpi * dpi));
                e.Graphics.DrawString("สรุปใบเสนอราคาประจำเดือน " + Month, new Font(font, 14, FontStyle.Bold), Brushes.Black, rect, strformat);

                //List Header
                Rectangle rec1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(35.6 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                Rectangle rec2 = new Rectangle((int)(38.4 / mmpi * dpi), (int)(35.6 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                Rectangle rec3 = new Rectangle((int)(64 / mmpi * dpi), (int)(35.6 / mmpi * dpi), (int)(51.2 / mmpi * dpi), (int)(9 / mmpi * dpi));
                Rectangle rec4 = new Rectangle((int)(115.2 / mmpi * dpi), (int)(35.6 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                Rectangle rec5 = new Rectangle((int)(140.8 / mmpi * dpi), (int)(35.6 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                Rectangle rec6 = new Rectangle((int)(166.4 / mmpi * dpi), (int)(35.6 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                e.Graphics.DrawString("วันที่", new Font(font, 12, FontStyle.Bold), Brushes.Black, rec1, strformat);
                e.Graphics.DrawString("รหัส", new Font(font, 12, FontStyle.Bold), Brushes.Black, rec2, strformat);
                e.Graphics.DrawString("ชื่อลูกค้า", new Font(font, 12, FontStyle.Bold), Brushes.Black, rec3, strformat);
                e.Graphics.DrawString("มูลค่าก่อนภาษี", new Font(font, 12, FontStyle.Bold), Brushes.Black, rec4, strformat);
                e.Graphics.DrawString("ภาษี 7 %", new Font(font, 12, FontStyle.Bold), Brushes.Black, rec5, strformat);
                e.Graphics.DrawString("มูลค่าหลังภาษี", new Font(font, 12, FontStyle.Bold), Brushes.Black, rec6, strformat);

                //List
                String dateTemp = "";
                strformat.Alignment = StringAlignment.Center;
                strformat.LineAlignment = StringAlignment.Center;
                int k = 0;
                for (int i = 0; i < 24; i++)
                {
                    if ((currentPage * 24) + i >= dt.Rows.Count)
                    {
                        k = i;
                        break;
                    }
                    Rectangle re1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)((44.6 + (i * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle re2 = new Rectangle((int)(38.4 / mmpi * dpi), (int)((44.6 + (i * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle re3 = new Rectangle((int)(64 / mmpi * dpi), (int)((44.6 + (i * 9)) / mmpi * dpi), (int)(51.2 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle re4 = new Rectangle((int)(115.2 / mmpi * dpi), (int)((44.6 + (i * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle re5 = new Rectangle((int)(140.8 / mmpi * dpi), (int)((44.6 + (i * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle re6 = new Rectangle((int)(166.4 / mmpi * dpi), (int)((44.6 + (i * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));

                    if (dateTemp != dt.Rows[(currentPage * 24) + i][1].ToString())
                    {
                        dateTemp = dt.Rows[(currentPage * 24) + i][1].ToString();
                        e.Graphics.DrawString(dt.Rows[(currentPage * 24) + i][1].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, re1, strformat);
                    }

                    e.Graphics.DrawString(dt.Rows[(currentPage * 24) + i][0].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, re2, strformat);
                    e.Graphics.DrawString(dt.Rows[(currentPage * 24) + i][2].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, re3, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dt.Rows[(currentPage * 24) + i][10].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, re4, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dt.Rows[(currentPage * 24) + i][11].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, re5, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dt.Rows[(currentPage * 24) + i][12].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, re6, strformat);
                    sum1 = sum1 + float.Parse(dt.Rows[(currentPage * 24) + i][10].ToString());
                    sum2 = sum2 + float.Parse(dt.Rows[(currentPage * 24) + i][11].ToString());
                    sum3 = sum3 + float.Parse(dt.Rows[(currentPage * 24) + i][12].ToString());
                }


                if (currentPage >= Page - 1)
                {
                    currentPage = 0;
                    Rectangle r3 = new Rectangle((int)(64 / mmpi * dpi), (int)((44.6 + (k * 9)) / mmpi * dpi), (int)(51.2 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle r4 = new Rectangle((int)(115.2 / mmpi * dpi), (int)((44.6 + (k * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle r5 = new Rectangle((int)(140.8 / mmpi * dpi), (int)((44.6 + (k * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    Rectangle r6 = new Rectangle((int)(166.4 / mmpi * dpi), (int)((44.6 + (k * 9)) / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(9 / mmpi * dpi));
                    strformat.Alignment = StringAlignment.Far;
                    strformat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("รวมทั้งหมด", new Font(font, 12, FontStyle.Bold), Brushes.Black, r3, strformat);
                    strformat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", sum1), new Font(font, 12, FontStyle.Bold), Brushes.Black, r4, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", sum2), new Font(font, 12, FontStyle.Bold), Brushes.Black, r5, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", sum3), new Font(font, 12, FontStyle.Bold), Brushes.Black, r6, strformat);

                    sum1 = 0.0f;
                    sum2 = 0.0f;
                    sum3 = 0.0f;
                    e.HasMorePages = false;
                }
                else
                {
                    currentPage++;
                    e.HasMorePages = true;
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
