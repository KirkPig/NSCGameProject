﻿using System;
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
    public partial class frmQuotationPrint : MetroFramework.Controls.MetroUserControl
    {

        SqlConnection conn = Connection.getConnection();

        private void reset()
        {
            metroGrid1.Rows.Clear();
            metroTextBox2.Text = "";
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
        }

        public frmQuotationPrint()
        {
            InitializeComponent();
        }

        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            PrintDoc printDoc = new PrintDoc();
            printDoc.printQuotation(metroTextBox2.Text, metroDateTime1.Text, metroTextBox3.Text, metroTextBox4.Text, metroTextBox5.Text, 
                metroTextBox6.Text, metroTextBox7.Text, metroTextBox8.Text, metroTextBox9.Text, metroTextBox10.Text, metroTextBox11.Text, 
                metroTextBox12.Text, metroTextBox13.Text, sender, e);

        }


        private void FrmQuotationPrint_Load(object sender, EventArgs e)
        {
            metroGrid1.RowHeadersVisible = false;
            metroDateTime1.Format = DateTimePickerFormat.Custom;
            metroDateTime1.CustomFormat = "dd-MM-yyyy";
        }

        private void MetroTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                String str = metroTextBox2.Text;


                SqlDataAdapter sda = new SqlDataAdapter("select * from tblQuotationReport where Id like '" + str + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                if (dt.Rows.Count == 1)
                {
                    metroTextBox2.Text = dt.Rows[0][0].ToString();
                    metroDateTime1.Value = DateTime.ParseExact(dt.Rows[0][1].ToString(), "dd-MM-yyyy", null);
                    metroTextBox3.Text = dt.Rows[0][2].ToString();
                    metroTextBox4.Text = dt.Rows[0][3].ToString();
                    metroTextBox5.Text = dt.Rows[0][4].ToString();
                    metroTextBox6.Text = dt.Rows[0][5].ToString();
                    metroTextBox7.Text = dt.Rows[0][6].ToString();
                    metroTextBox8.Text = dt.Rows[0][7].ToString();
                    metroTextBox9.Text = dt.Rows[0][8].ToString();
                    metroTextBox10.Text = dt.Rows[0][9].ToString();

                    metroTextBox11.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][10].ToString()));
                    metroTextBox12.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][11].ToString()));
                    metroTextBox13.Text = String.Format("{0:0.00}", float.Parse(dt.Rows[0][12].ToString()));

                    sda = new SqlDataAdapter("select * from tblQuotation where Id like '" + str + "%'", conn);

                    dt = new DataTable();
                    sda.Fill(dt);

                    metroGrid1.Rows.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        int n = metroGrid1.Rows.Add();
                        metroGrid1.Rows[n].Cells[0].Value = dr[1].ToString();
                        metroGrid1.Rows[n].Cells[1].Value = dr[2].ToString();
                        metroGrid1.Rows[n].Cells[2].Value = dr[3].ToString();
                        metroGrid1.Rows[n].Cells[3].Value = dr[4].ToString();
                        metroGrid1.Rows[n].Cells[4].Value = String.Format("{0:0.00}", float.Parse(dr[5].ToString()));
                        metroGrid1.Rows[n].Cells[5].Value = dr[6].ToString();
                        metroGrid1.Rows[n].Cells[6].Value = String.Format("{0:0.00}", float.Parse(dr[7].ToString()));
                    }

                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }
        }

        private void MetroButton3_Click(object sender, EventArgs e)
        {
            printPreviewDialog1.ShowDialog();
            DialogResult dialogResult = printDialog1.ShowDialog();
            if (dialogResult == DialogResult.Cancel) { }
            else printDocument1.Print();
        }
        private void MetroButton2_Click(object sender, EventArgs e)
        {
            reset();
        }
    }
}
