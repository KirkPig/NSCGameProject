using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectHouse
{
    public class PrintDoc
    {

        private static int dpi = 96;
        private static float mmpi = 25.4f;
        private static SqlConnection conn = Connection.getConnection();
        

        private static string ThaiBaht(string txt)
        {
            string bahtTxt, n, bahtTH = "";
            double amount;
            try { amount = Convert.ToDouble(txt); }
            catch { amount = 0; }
            bahtTxt = amount.ToString("####.00");
            string[] num = { "ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า", "สิบ" };
            string[] rank = { "", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน" };
            string[] temp = bahtTxt.Split('.');
            string intVal = temp[0];
            string decVal = temp[1];
            if (Convert.ToDouble(bahtTxt) == 0)
                bahtTH = "ศูนย์บาทถ้วน";
            else
            {
                for (int i = 0; i < intVal.Length; i++)
                {
                    n = intVal.Substring(i, 1);
                    if (n != "0")
                    {
                        if ((i == (intVal.Length - 1)) && (n == "1"))
                            bahtTH += "เอ็ด";
                        else if ((i == (intVal.Length - 2)) && (n == "2"))
                            bahtTH += "ยี่";
                        else if ((i == (intVal.Length - 2)) && (n == "1"))
                            bahtTH += "";
                        else
                            bahtTH += num[Convert.ToInt32(n)];
                        bahtTH += rank[(intVal.Length - i) - 1];
                    }
                }
                bahtTH += "บาท";
                if (decVal == "00")
                    bahtTH += "ถ้วน";
                else
                {
                    for (int i = 0; i < decVal.Length; i++)
                    {
                        n = decVal.Substring(i, 1);
                        if (n != "0")
                        {
                            if ((i == decVal.Length - 1) && (n == "1"))
                                bahtTH += "เอ็ด";
                            else if ((i == (decVal.Length - 2)) && (n == "2"))
                                bahtTH += "ยี่";
                            else if ((i == (decVal.Length - 2)) && (n == "1"))
                                bahtTH += "";
                            else
                                bahtTH += num[Convert.ToInt32(n)];
                            bahtTH += rank[(decVal.Length - i) - 1];
                        }
                    }
                    bahtTH += "สตางค์";
                }
            }
            return bahtTH;
        }

        private static void docHeader(String font, Brush col, String doc, String ID, String Date, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;

            //Header
            Rectangle rect1 = new Rectangle((int)(122.2 / mmpi * dpi), (int)(10 / mmpi * dpi), (int)(87.8 / mmpi * dpi), (int)(15.6 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rect1);
            e.Graphics.DrawString(doc, new Font(font, 28, FontStyle.Bold), Brushes.Black, rect1, strformat);

            Rectangle rect2 = new Rectangle((int)(122.2 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(12.8 / mmpi * dpi));
            strformat.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("NO." + ID + "  DATE: " + Date, new Font(font, 14, FontStyle.Bold), Brushes.Black, rect2, strformat);

            Rectangle rect3 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(38.4 / mmpi * dpi), (int)(80.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            strformat.Alignment = StringAlignment.Near;
            e.Graphics.DrawString("YONO TOOLS CO.,LTD.", new Font(font, 24, FontStyle.Bold), Brushes.Black, rect3, strformat);

            Rectangle rect4 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(47.1 / mmpi * dpi), (int)(98 / mmpi * dpi), (int)(33.8 / mmpi * dpi));
            strformat.LineAlignment = StringAlignment.Near;
            e.Graphics.DrawString("103/314 M.5 T.Phanthai Norasing, A.Muang Samut Sakhon,\nSamut Sakhon 74000" +
                "\nTEL: 034-116655  Fax: 034-116656  MOBILE: 099-0568889\nE-MAIL: sale.yonotools@gmail.com\n" +
                "TAX-ID: 0125560000590", new Font(font, 13, FontStyle.Bold), Brushes.Black, rect4, strformat);

            Rectangle rect5 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(10 / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(26.9 / mmpi * dpi));
            Image img = Image.FromFile(Connection.getLogo());
            e.Graphics.DrawImage(img, rect5);

        }

        private static void docCustomerHeader(String font, Brush col, String CustomerName, String CustomerTaxID, String CustomerAddress, String CustomerTel, 
            String CustomerFax, String CustomerEmail, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;

            Rectangle r3 = new Rectangle((int)(128.1 / mmpi * dpi), (int)((38.3) / mmpi * dpi), (int)(69.1 / mmpi * dpi), (int)(68 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r3);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Near;
            Rectangle r4 = new Rectangle((int)(134 / mmpi * dpi), (int)((42.4) / mmpi * dpi), (int)(57.5 / mmpi * dpi), (int)(17.5 / mmpi * dpi));
            e.Graphics.DrawString("เลขประจำตัวผู้เสียภาษีอากร\n" + CustomerTaxID, new Font(font, 12, FontStyle.Bold), Brushes.Black, r4, strformat);

            string fax = CustomerFax;
            if (fax == "")
            {
                fax = "-";
            }
            Rectangle r5 = new Rectangle((int)(134 / mmpi * dpi), (int)((59.7) / mmpi * dpi), (int)(63.3 / mmpi * dpi), (int)(46.7 / mmpi * dpi));
            e.Graphics.DrawString("ข้อมูลลูกค้า\n" + CustomerName + "\n" + CustomerAddress + "\nTel : " + CustomerTel + "\nFax : " + fax,
                new Font(font, 12, FontStyle.Bold), Brushes.Black, r5, strformat);
        }

        public void printQuotation(String ID, String Date, String CustomerName, String CustomerTaxID, String CustomerAddress, 
            String CustomerTel, String CustomerFax, String CustomerEmail, String ATTN, String CR, String ValueBeforeTax, 
            String ValueTax, String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            String font = "TH Sarabun New";
            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Brush col = Brushes.LightCyan;

            //Header
            docHeader(font, col, "ใบเสนอราคา (Quotation)", ID, Date, sender, e);

            //Side Header
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle r1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(35 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r1);
            e.Graphics.DrawString("ATTN: " + ATTN, new Font(font, 14, FontStyle.Bold), Brushes.Black, r1, strformat);

            Rectangle r2 = new Rectangle((int)(59.3 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r2);
            if (int.Parse(CR) == 0)
            {
                e.Graphics.DrawString("CR: CASH", new Font(font, 14, FontStyle.Bold), Brushes.Black, r2, strformat);
            }
            else if (int.Parse(CR) == 1)
            {
                e.Graphics.DrawString("CR: " + CR + " DAY", new Font(font, 14, FontStyle.Bold), Brushes.Black, r2, strformat);
            }
            else
            {
                e.Graphics.DrawString("CR: " + CR + " DAYS", new Font(font, 14, FontStyle.Bold), Brushes.Black, r2, strformat);
            }

            //Customer Header
            docCustomerHeader(font, col, CustomerName, CustomerTaxID, CustomerAddress, CustomerTel, CustomerFax, CustomerEmail, sender, e);

            //List Header
            Rectangle re1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re5 = new Rectangle((int)(128.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, re1);
            e.Graphics.FillRectangle(col, re2);
            e.Graphics.FillRectangle(col, re3);
            e.Graphics.FillRectangle(col, re4);
            e.Graphics.FillRectangle(col, re5);
            e.Graphics.FillRectangle(col, re6);
            e.Graphics.FillRectangle(col, re7);
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("NO", new Font(font, 12, FontStyle.Bold), Brushes.Black, re1, strformat);
            e.Graphics.DrawString("DESCRIPTION", new Font(font, 12, FontStyle.Bold), Brushes.Black, re2, strformat);
            e.Graphics.DrawString("QUANTITY", new Font(font, 12, FontStyle.Bold), Brushes.Black, re3, strformat);
            e.Graphics.DrawString("UNIT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re4, strformat);
            e.Graphics.DrawString("PRICE", new Font(font, 12, FontStyle.Bold), Brushes.Black, re5, strformat);
            e.Graphics.DrawString("DISCOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re6, strformat);
            e.Graphics.DrawString("AMOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re7, strformat);

            //List
            try
            {

                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblQuotation where Id like '" + ID + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);
                int i = 0;

                foreach (DataRow dr in dt.Rows)
                {
                    int a = 6;
                    
                    String st = dr[2].ToString();
                    String st2 = "";
                    bool A = false;
                    for (int j = 0; j < st.Length; j++)
                    {
                        if (A)
                        {
                            st2 = st2 + st[j].ToString();
                        }
                        else if (st[j] == ':')
                        {
                            j++;
                            A = true;
                        }
                    }
                    String desc = st2;
                    //if (desc.Length > 40) a = 12;
                    Rectangle rec1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec5 = new Rectangle((int)(128 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((123.1 + (i * a)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(a / mmpi * dpi));
                    
                    strformat.Alignment = StringAlignment.Center;
                    strformat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0}", i + 1), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec1, strformat);
                    strformat.Alignment = StringAlignment.Near;
                    e.Graphics.DrawString(desc, new Font(font, 12, FontStyle.Regular), Brushes.Black, rec2, strformat);
                    strformat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[3].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec3, strformat);
                    e.Graphics.DrawString(dr[4].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec4, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[5].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec5, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[6].ToString())) + "%", new Font(font, 12, FontStyle.Regular), Brushes.Black, rec6, strformat);
                    strformat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[7].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec7, strformat);
                    i++;
                    //if (desc.Length > 40) i++;
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }

            //Footer
            Rectangle rc1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc2 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rc1);
            e.Graphics.FillRectangle(col, rc2);

            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc3 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(103.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc3, strformat);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc5 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc6 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString("มูลค่าก่อนภาษี", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc4, strformat);
            e.Graphics.DrawString("ภาษีมูลค่าเพิ่ม", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc5, strformat);
            e.Graphics.DrawString("รวมสุทธิ", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc6, strformat);

            strformat.Alignment = StringAlignment.Far;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc7 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc8 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc9 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueBeforeTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc7, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc8, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc9, strformat);

            //Signature
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rtg1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(93.1 / mmpi * dpi), (int)(27 / mmpi * dpi));
            Rectangle rtg2 = new Rectangle((int)(105.9 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(91.2 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้ขอซื้อ\nวันที่....../....../......", 
                new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg1, strformat);
            e.Graphics.DrawString("ลงชื่อ.............................................\nพนักงานขาย\nวันที่....../....../......", 
                new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg2, strformat);


            //DrawLine

            /*Vertical*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(18.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(18.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(93.6 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(93.6 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(116.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(128 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(128 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(151.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(151.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(174.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(174.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            /*Horizintal*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(123.1 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(123.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(233 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(241.2 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(241.2 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));

        }

        public void printOrder(String ID, String Date, String CustomerName, String CustomerTaxID, String CustomerAddress,
            String CustomerTel, String CustomerFax, String CustomerEmail, String TermOfPayment, String ValueBeforeTax,
            String ValueTax, String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String font = "TH Sarabun New";
            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Brush col = Brushes.LightGray;

            //Header
            docHeader(font, col, "ใบสั่งซื้อ (Orders)", ID, Date, sender, e);

            //Side Header
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle r1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(69.9 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r1);
            e.Graphics.DrawString("TERM OF PAYMENTS: " + TermOfPayment, new Font(font, 14, FontStyle.Bold), Brushes.Black, r1, strformat);

            //Customer Header
            docCustomerHeader(font, col, CustomerName, CustomerTaxID, CustomerAddress, CustomerTel, CustomerFax, CustomerEmail, sender, e);

            //List Header
            Rectangle re1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re5 = new Rectangle((int)(128.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, re1);
            e.Graphics.FillRectangle(col, re2);
            e.Graphics.FillRectangle(col, re3);
            e.Graphics.FillRectangle(col, re4);
            e.Graphics.FillRectangle(col, re5);
            e.Graphics.FillRectangle(col, re6);
            e.Graphics.FillRectangle(col, re7);
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("NO", new Font(font, 12, FontStyle.Bold), Brushes.Black, re1, strformat);
            e.Graphics.DrawString("DESCRIPTION", new Font(font, 12, FontStyle.Bold), Brushes.Black, re2, strformat);
            e.Graphics.DrawString("QUANTITY", new Font(font, 12, FontStyle.Bold), Brushes.Black, re3, strformat);
            e.Graphics.DrawString("UNIT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re4, strformat);
            e.Graphics.DrawString("PRICE", new Font(font, 12, FontStyle.Bold), Brushes.Black, re5, strformat);
            e.Graphics.DrawString("DISCOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re6, strformat);
            e.Graphics.DrawString("AMOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re7, strformat);

            //List
            try
            {
                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblOrders where Id like '" + ID + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);
                int i = 0;

                foreach (DataRow dr in dt.Rows)
                {
                    String st = dr[2].ToString();
                    String st2 = "";
                    bool A = false;
                    for (int j = 0; j < st.Length; j++)
                    {
                        if (A)
                        {
                            st2 = st2 + st[j].ToString();
                        }
                        else if (st[j] == ':')
                        {
                            j++;
                            A = true;
                        }
                    }
                    String desc = st2;
                    Rectangle rec1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec5 = new Rectangle((int)(128 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    strformat.Alignment = StringAlignment.Center;
                    strformat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0}", i + 1), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec1, strformat);
                    strformat.Alignment = StringAlignment.Near;
                    e.Graphics.DrawString(desc, new Font(font, 12, FontStyle.Regular), Brushes.Black, rec2, strformat);
                    strformat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[3].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec3, strformat);
                    e.Graphics.DrawString(dr[4].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec4, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[5].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec5, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[6].ToString())) + "%", new Font(font, 12, FontStyle.Regular), Brushes.Black, rec6, strformat);
                    strformat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[7].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec7, strformat);
                    i++;
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }

            //Footer
            Rectangle rc1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc2 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rc1);
            e.Graphics.FillRectangle(col, rc2);

            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc3 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(103.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc3, strformat);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc5 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc6 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString("มูลค่าก่อนภาษี", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc4, strformat);
            e.Graphics.DrawString("ภาษีมูลค่าเพิ่ม", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc5, strformat);
            e.Graphics.DrawString("รวมสุทธิ", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc6, strformat);

            strformat.Alignment = StringAlignment.Far;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc7 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc8 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc9 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueBeforeTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc7, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc8, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc9, strformat);

            //Signature
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rtg1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(93.1 / mmpi * dpi), (int)(27 / mmpi * dpi));
            Rectangle rtg2 = new Rectangle((int)(105.9 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(91.2 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้ขอซื้อ\nวันที่....../....../......", 
                new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg1, strformat);
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้อนุมัติ\nวันที่....../....../......", 
                new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg2, strformat);


            //DrawLine

            /*Vertical*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(18.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(18.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(93.6 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(93.6 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(116.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(128 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(128 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(151.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(151.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(174.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(174.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            /*Horizintal*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(123.1 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(123.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(233 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(241.2 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(241.2 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));

        }

        public void printProductLoan(String ID, String Date, String CustomerName, String CustomerTaxID, String CustomerAddress,
            String CustomerTel, String CustomerFax, String CustomerEmail, String Contact, String ValueBeforeTax,
            String ValueTax, String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String font = "TH Sarabun New";
            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Brush col = Brushes.LightPink;

            //Header
            docHeader(font, col, "ใบยืมสินค้า(Product Loan)", ID, Date, sender, e);

            //Side Header
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle r1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(69.9 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r1);
            e.Graphics.DrawString("ติดต่อ: " + Contact, new Font(font, 14, FontStyle.Bold), Brushes.Black, r1, strformat);

            //Customer Header
            docCustomerHeader(font, col, CustomerName, CustomerTaxID, CustomerAddress, CustomerTel, CustomerFax, CustomerEmail, sender, e);

            //List
            try
            {
                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblProductLoan where Id like '" + ID + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);
                int i = 0;

                foreach (DataRow dr in dt.Rows)
                {
                    String st = dr[2].ToString();
                    String st2 = "";
                    bool A = false;
                    for (int j = 0; j < st.Length; j++)
                    {
                        if (A)
                        {
                            st2 = st2 + st[j].ToString();
                        }
                        else if (st[j] == ':')
                        {
                            j++;
                            A = true;
                        }
                    }
                    String desc = st2;
                    Rectangle rec1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec5 = new Rectangle((int)(128 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    strformat.Alignment = StringAlignment.Center;
                    strformat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0}", i + 1), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec1, strformat);
                    strformat.Alignment = StringAlignment.Near;
                    e.Graphics.DrawString(desc, new Font(font, 12, FontStyle.Regular), Brushes.Black, rec2, strformat);
                    strformat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[3].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec3, strformat);
                    e.Graphics.DrawString(dr[4].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec4, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[5].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec5, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[6].ToString())) + "%", new Font(font, 12, FontStyle.Regular), Brushes.Black, rec6, strformat);
                    strformat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[7].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec7, strformat);
                    i++;
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }

            //Footer
            Rectangle rc1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc2 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rc1);
            e.Graphics.FillRectangle(col, rc2);

            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc3 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(103.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc3, strformat);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc5 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc6 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString("มูลค่าก่อนภาษี", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc4, strformat);
            e.Graphics.DrawString("ภาษีมูลค่าเพิ่ม", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc5, strformat);
            e.Graphics.DrawString("รวมสุทธิ", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc6, strformat);

            strformat.Alignment = StringAlignment.Far;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc7 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc8 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc9 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueBeforeTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc7, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc8, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc9, strformat);

            //P.S.
            e.Graphics.DrawString("หมายเหตุ : สินค้าตามรายการข้างต้น หากมีการเสียหายหรือขาดตกบกพร่อง โปรดแจ้งให้ทราบภายใน 3 วัน นับจากวันที่ได้รับสินค้า มิฉะนั้น ทางบริษัทฯ จะไม่รับผิดชอบใดๆ ทั้งสิ้น"
                , new Font(font, 10, FontStyle.Bold), Brushes.Black, (int)(14 / mmpi * dpi), (int)((251) / mmpi * dpi));


            //Signature
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rtg1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(93.1 / mmpi * dpi), (int)(27 / mmpi * dpi));
            Rectangle rtg2 = new Rectangle((int)(105.9 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(91.2 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้ส่งสินค้า\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg1, strformat);
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้รับสินค้า\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg2, strformat);


            //DrawLine

            /*Vertical*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(18.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(18.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(93.6 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(93.6 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(116.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(128 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(128 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(151.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(151.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(174.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(174.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            /*Horizintal*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(123.1 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(123.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(233 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(241.2 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(241.2 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));

        }

        public void printDelivery(String ID, String Date, String CustomerName, String CustomerTaxID, String CustomerAddress,
            String CustomerTel, String CustomerFax, String CustomerEmail, String Contact, String ValueBeforeTax,
            String ValueTax, String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e, Brush col)
        {
            String font = "TH Sarabun New";
            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;



            //Header
            Rectangle rect1 = new Rectangle((int)(122.2 / mmpi * dpi), (int)(10 / mmpi * dpi), (int)(87.8 / mmpi * dpi), (int)(15.6 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rect1);
            e.Graphics.DrawString("ใบส่งสินค้า(Delivery Note)", new Font(font, 28, FontStyle.Bold), Brushes.Black, rect1, strformat);

            Rectangle rect2 = new Rectangle((int)(122.2 / mmpi * dpi), (int)(25.6 / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(12.8 / mmpi * dpi));
            strformat.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("NO." + ID + "  DATE: " + Date, new Font(font, 14, FontStyle.Bold), Brushes.Black, rect2, strformat);

            Rectangle rect3 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(38.4 / mmpi * dpi), (int)(80.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            strformat.Alignment = StringAlignment.Near;
            e.Graphics.DrawString("YONO TOOLS CO.,LTD.", new Font(font, 24, FontStyle.Bold), Brushes.Black, rect3, strformat);

            Rectangle rect4 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(47.1 / mmpi * dpi), (int)(98 / mmpi * dpi), (int)(33.8 / mmpi * dpi));
            strformat.LineAlignment = StringAlignment.Near;
            e.Graphics.DrawString("103/314 M.5 T.Phanthai Norasing, A.Muang Samut Sakhon,\nSamut Sakhon 74000" +
                "\nTEL: 034-116655  Fax: 034-116656  MOBILE: 099-0568889\nE-MAIL: sale.yonotools@gmail.com\n" +
                "TAX-ID: 0125560000590", new Font(font, 13, FontStyle.Bold), Brushes.Black, rect4, strformat);

            Rectangle rect5 = new Rectangle((int)(12.8 / mmpi * dpi), (int)(10 / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(26.9 / mmpi * dpi));
            Image img = Image.FromFile(Connection.getLogo());
            e.Graphics.DrawImage(img, rect5);




            //Side Header
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle r1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(69.9 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r1);
            e.Graphics.DrawString("ติดต่อ: " + Contact, new Font(font, 14, FontStyle.Bold), Brushes.Black, r1, strformat);




            //Customer Header
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;

            Rectangle r3 = new Rectangle((int)(128.1 / mmpi * dpi), (int)((38.3) / mmpi * dpi), (int)(69.1 / mmpi * dpi), (int)(68 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r3);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Near;
            Rectangle r4 = new Rectangle((int)(134 / mmpi * dpi), (int)((42.4) / mmpi * dpi), (int)(57.5 / mmpi * dpi), (int)(17.5 / mmpi * dpi));
            e.Graphics.DrawString("เลขประจำตัวผู้เสียภาษีอากร\n" + CustomerTaxID, new Font(font, 12, FontStyle.Bold), Brushes.Black, r4, strformat);

            string fax = CustomerFax;
            if (fax == "")
            {
                fax = "-";
            }
            Rectangle r5 = new Rectangle((int)(134 / mmpi * dpi), (int)((59.7) / mmpi * dpi), (int)(63.3 / mmpi * dpi), (int)(46.7 / mmpi * dpi));
            e.Graphics.DrawString("ข้อมูลลูกค้า\n" + CustomerName + "\n" + CustomerAddress + "\nTel : " + CustomerTel + "\nFax : " + fax,
                new Font(font, 12, FontStyle.Bold), Brushes.Black, r5, strformat);






            //List Header
            Rectangle re1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re5 = new Rectangle((int)(128.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, re1);
            e.Graphics.FillRectangle(col, re2);
            e.Graphics.FillRectangle(col, re3);
            e.Graphics.FillRectangle(col, re4);
            e.Graphics.FillRectangle(col, re5);
            e.Graphics.FillRectangle(col, re6);
            e.Graphics.FillRectangle(col, re7);
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("NO", new Font(font, 12, FontStyle.Bold), Brushes.Black, re1, strformat);
            e.Graphics.DrawString("DESCRIPTION", new Font(font, 12, FontStyle.Bold), Brushes.Black, re2, strformat);
            e.Graphics.DrawString("QUANTITY", new Font(font, 12, FontStyle.Bold), Brushes.Black, re3, strformat);
            e.Graphics.DrawString("UNIT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re4, strformat);
            e.Graphics.DrawString("PRICE", new Font(font, 12, FontStyle.Bold), Brushes.Black, re5, strformat);
            e.Graphics.DrawString("DISCOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re6, strformat);
            e.Graphics.DrawString("AMOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re7, strformat);





            //List
            try
            {
                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblDelivery where Id like '" + ID + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);
                int i = 0;

                foreach (DataRow dr in dt.Rows)
                {
                    String st = dr[2].ToString();
                    String st2 = "";
                    bool A = false;
                    for (int j = 0; j < st.Length; j++)
                    {
                        if (A)
                        {
                            st2 = st2 + st[j].ToString();
                        }
                        else if (st[j] == ':')
                        {
                            j++;
                            A = true;
                        }
                    }
                    String desc = st2;
                    Rectangle rec1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(93.6 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23.3 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(11.4 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec5 = new Rectangle((int)(128 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec6 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec7 = new Rectangle((int)(174.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(23 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    strformat.Alignment = StringAlignment.Center;
                    strformat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0}", i + 1), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec1, strformat);
                    strformat.Alignment = StringAlignment.Near;
                    e.Graphics.DrawString(desc, new Font(font, 12, FontStyle.Regular), Brushes.Black, rec2, strformat);
                    strformat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[3].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec3, strformat);
                    e.Graphics.DrawString(dr[4].ToString(), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec4, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[5].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec5, strformat);
                    e.Graphics.DrawString(String.Format("{0:0,0}", float.Parse(dr[6].ToString())) + "%", new Font(font, 12, FontStyle.Regular), Brushes.Black, rec6, strformat);
                    strformat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[7].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec7, strformat);
                    i++;
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }






            //Footer
            Rectangle rc1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc2 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rc1);
            e.Graphics.FillRectangle(col, rc2);

            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc3 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(103.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc3, strformat);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc4 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc5 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc6 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString("มูลค่าก่อนภาษี", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc4, strformat);
            e.Graphics.DrawString("ภาษีมูลค่าเพิ่ม", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc5, strformat);
            e.Graphics.DrawString("รวมสุทธิ", new Font(font, 14, FontStyle.Bold), Brushes.Black, rc6, strformat);

            strformat.Alignment = StringAlignment.Far;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc7 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc8 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((233) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc9 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueBeforeTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc7, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc8, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font(font, 14, FontStyle.Bold), Brushes.Black, rc9, strformat);






            //P.S.
            e.Graphics.DrawString("หมายเหตุ : สินค้าตามรายการข้างต้น หากมีการเสียหายหรือขาดตกบกพร่อง โปรดแจ้งให้ทราบภายใน 3 วัน นับจากวันที่ได้รับสินค้า มิฉะนั้น ทางบริษัทฯ จะไม่รับผิดชอบใดๆ ทั้งสิ้น"
                , new Font(font, 10, FontStyle.Bold), Brushes.Black, (int)(14 / mmpi * dpi), (int)((251) / mmpi * dpi));






            //Signature
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rtg1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(93.1 / mmpi * dpi), (int)(27 / mmpi * dpi));
            Rectangle rtg2 = new Rectangle((int)(105.9 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(91.2 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้ส่งสินค้า\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg1, strformat);
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้รับสินค้า\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg2, strformat);







            //DrawLine

            /*Vertical*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(18.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(18.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(93.6 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(93.6 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(116.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(128 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(128 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(151.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(151.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(174.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(174.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            /*Horizintal*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(123.1 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(123.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(224.3 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(224.3 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(233 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(241.2 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(241.2 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));


        }

        public void printCreditNote(String ID, String Date, String CustomerName, String CustomerTaxID, String CustomerAddress,
            String CustomerTel, String CustomerFax, String CustomerEmail, String InvoiceID, String InvoiceDate, String ValueOld, String ValueReal,
            String ValueBeforeTax,  String ValueTax, String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String font = "TH Sarabun New";
            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Brush col = Brushes.Coral;

            //Header
            docHeader(font, col, "ใบลดหนี้ (Credit Note)", ID, Date, sender, e);

            //Side Header
            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle r1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(84.9 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.FillRectangle(col, r1);
            r1 = new Rectangle((int)(13 / mmpi * dpi), (int)((89.2) / mmpi * dpi), (int)(84.9 / mmpi * dpi), (int)(17.2 / mmpi * dpi));
            e.Graphics.DrawString("อ้างถึงเลขที่ใบกำกับภาษี(ฉบับเดิม) " + InvoiceID + "\nวันที่ตามใบกำกับภาษี(ฉบับเดิม) " + InvoiceDate, new Font(font, 12, FontStyle.Bold), Brushes.Black, r1, strformat);

            //Customer Header
            docCustomerHeader(font, col, CustomerName, CustomerTaxID, CustomerAddress, CustomerTel, CustomerFax, CustomerEmail, sender, e);
            
            //List Header
            Rectangle re1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(132.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle re3 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((114.6) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, re1);
            e.Graphics.FillRectangle(col, re2);
            e.Graphics.FillRectangle(col, re3);
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("NO", new Font(font, 12, FontStyle.Bold), Brushes.Black, re1, strformat);
            e.Graphics.DrawString("DESCRIPTION", new Font(font, 12, FontStyle.Bold), Brushes.Black, re2, strformat);
            e.Graphics.DrawString("AMOUNT", new Font(font, 12, FontStyle.Bold), Brushes.Black, re3, strformat);

            //List
            try
            {
                conn.Open();

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblCreditNote where Id like '" + ID + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);
                int i = 0;

                foreach (DataRow dr in dt.Rows)
                {
                    String st = dr[2].ToString();
                    String st2 = "";
                    bool A = false;
                    for (int j = 0; j < st.Length; j++)
                    {
                        if (A)
                        {
                            st2 = st2 + st[j].ToString();
                        }
                        else if (st[j] == ':')
                        {
                            j++;
                            A = true;
                        }
                    }
                    String desc = st2;
                    Rectangle rec1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(6.1 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(18.7 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(75 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((123.1 + (i * 6)) / mmpi * dpi), (int)(46.3 / mmpi * dpi), (int)(6 / mmpi * dpi));
                    strformat.Alignment = StringAlignment.Center;
                    strformat.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0}", i + 1), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec1, strformat);
                    strformat.Alignment = StringAlignment.Near;
                    e.Graphics.DrawString(desc, new Font(font, 12, FontStyle.Regular), Brushes.Black, rec2, strformat);
                    strformat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[3].ToString())), new Font(font, 12, FontStyle.Regular), Brushes.Black, rec3, strformat);
                    i++;
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }

            //Footer
            Rectangle rc1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((189.5) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc2 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((206.9) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc3 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(184.6 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rc1);
            e.Graphics.FillRectangle(col, rc2);
            e.Graphics.FillRectangle(col, rc3);

            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc4 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(103.8 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font(font, 12, FontStyle.Bold), Brushes.Black, rc4, strformat);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc5 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((189.5) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc6 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((198.2) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc7 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((206.9) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc8 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((215.6) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc9 = new Rectangle((int)(116.7 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString("มูลค่าตามใบกำกับภาษีเดิม", new Font(font, 12, FontStyle.Bold), Brushes.Black, rc5, strformat);
            e.Graphics.DrawString("มูลค่าที่ถูกต้อง", new Font(font, 12, FontStyle.Bold), Brushes.Black, rc6, strformat);
            e.Graphics.DrawString("มูลค่าผลต่าง(ก่อนVAT)", new Font(font, 12, FontStyle.Bold), Brushes.Black, rc7, strformat);
            e.Graphics.DrawString("VAT 7%", new Font(font, 12, FontStyle.Bold), Brushes.Black, rc8, strformat);
            e.Graphics.DrawString("รวมมูลค่า(รวมVAT)", new Font(font, 12, FontStyle.Bold), Brushes.Black, rc9, strformat);

            strformat.Alignment = StringAlignment.Far;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc10 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((189.5) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc11 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((198.2) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc12 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((206.9) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc13 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((215.6) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            Rectangle rc14 = new Rectangle((int)(151.1 / mmpi * dpi), (int)((224.3) / mmpi * dpi), (int)(46.1 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueOld)), new Font(font, 12, FontStyle.Bold), Brushes.Black, rc10, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueReal)), new Font(font, 12, FontStyle.Bold), Brushes.Black, rc11, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueBeforeTax)), new Font(font, 12, FontStyle.Bold), Brushes.Black, rc12, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueTax)), new Font(font, 12, FontStyle.Bold), Brushes.Black, rc13, strformat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font(font, 12, FontStyle.Bold), Brushes.Black, rc14, strformat);

            //Side Footer
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rct1 = new Rectangle((int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(50 / mmpi * dpi), (int)(8.7 / mmpi * dpi));
            e.Graphics.FillRectangle(col, rct1);
            e.Graphics.DrawString("เหตุผลการลดหนี้", new Font(font, 12, FontStyle.Bold), Brushes.Black, rct1, strformat);

            //Signature
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rtg1 = new Rectangle((int)(12.8 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(62.4 / mmpi * dpi), (int)(27 / mmpi * dpi));
            Rectangle rtg2 = new Rectangle((int)(75.2 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(61.5 / mmpi * dpi), (int)(27 / mmpi * dpi));
            Rectangle rtg3 = new Rectangle((int)(136.8 / mmpi * dpi), (int)((258.1) / mmpi * dpi), (int)(60.4 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้จัดทำ\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg1, strformat);
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้มีอำนาจลงนาม\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg2, strformat);
            e.Graphics.DrawString("ลงชื่อ.............................................\nผู้อนุมัติ\nวันที่....../....../......", new Font(font, 14, FontStyle.Bold), Brushes.Black, rtg3, strformat);


            //DrawLine

            /*Vertical*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(12.7 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(18.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(18.7 / mmpi * dpi), (int)(189.5 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(151.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(151.1 / mmpi * dpi), (int)(189.5 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(189.5 / mmpi * dpi), (int)(116.7 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(233 / mmpi * dpi));

            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(62.7 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(62.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(197.1 / mmpi * dpi), (int)((241.2) / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));


            /*Horizintal*/
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(114.6 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(114.6 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(123.1 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(123.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(189.5 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(189.5 / mmpi * dpi));

            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(198.1 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(198.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(206.8 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(206.8 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(215.5 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(215.5 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(116.7 / mmpi * dpi), (int)(224.2 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(224.2 / mmpi * dpi));

            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(233 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(233 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(241.2 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(241.2 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(12.7 / mmpi * dpi), (int)(249.9 / mmpi * dpi), (int)(197.1 / mmpi * dpi), (int)(249.9 / mmpi * dpi));

        }

        public void printInvoice(String ID, String Date, String CustomerName, String CustomerTaxID, String CustomerAddress,
            String CustomerTel, String CustomerFax, String CustomerEmail, String PONum, String OrderBy, String TermOfPayment, String DateDue,
            String Sales, String ValueBeforeTax, String ValueTax, String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            dpi = 105;

            //Header
            string fax = CustomerFax;
            if (fax == "")
            {
                fax = "-";
            }

            e.Graphics.DrawString(ID, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, (int)(155.4 / mmpi * dpi), (int)(55.6 / mmpi * dpi));
            e.Graphics.DrawString(Date, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, (int)(155.4 / mmpi * dpi), (int)(66.5 / mmpi * dpi));
            e.Graphics.DrawString("เลขประจำตัวผู้เสียภาษีอากร : " + CustomerTaxID, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, (int)(17.5 / mmpi * dpi), (int)(45 / mmpi * dpi));
            e.Graphics.DrawString(CustomerName + "\n" + CustomerAddress + "\nTel : " + CustomerTel + "  Fax : " + fax
                , new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, (int)(25.9 / mmpi * dpi), (int)(54 / mmpi * dpi));

            //Side Header
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;
            Rectangle rect1 = new Rectangle((int)(2.4 / mmpi * dpi), (int)(86 / mmpi * dpi), (int)(33.6 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle rect2 = new Rectangle((int)(36 / mmpi * dpi), (int)(86 / mmpi * dpi), (int)(37.6 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle rect3 = new Rectangle((int)(69 / mmpi * dpi), (int)(86 / mmpi * dpi), (int)(37.6 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle rect4 = new Rectangle((int)(105.1 / mmpi * dpi), (int)(86 / mmpi * dpi), (int)(44.7 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle rect5 = new Rectangle((int)(145.8 / mmpi * dpi), (int)(86 / mmpi * dpi), (int)(41.3 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            e.Graphics.DrawString(PONum, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rect1, stringFormat);
            e.Graphics.DrawString(OrderBy, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rect2, stringFormat);
            e.Graphics.DrawString(TermOfPayment, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rect3, stringFormat);
            e.Graphics.DrawString(DateDue, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rect4, stringFormat);
            e.Graphics.DrawString(Sales, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rect5, stringFormat);

            //List
            try
            {
                conn.Open();

                String str = ID;

                SqlDataAdapter sda = new SqlDataAdapter("select * from tblInvoice where Id like '" + str + "%'", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);

                int i = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    int a = 6;
                    String st = dr[2].ToString();
                    String st2 = "";
                    bool A = false;
                    for (int j = 0; j < st.Length; j++)
                    {
                        if (A)
                        {
                            st2 = st2 + st[j].ToString();
                        }
                        else if (st[j] == ':')
                        {
                            j++;
                            A = true;
                        }
                    }
                    String desc = st2;
                    //if (desc.Length > 40) a = 12;
                    Rectangle rec1 = new Rectangle((int)(2.4 / mmpi * dpi), (int)((102.7 + (i * a)) / mmpi * dpi), (int)(12.4 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(14.8 / mmpi * dpi), (int)((102.7 + (i * a)) / mmpi * dpi), (int)(90.2 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(100 / mmpi * dpi), (int)((102.7 + (i * a)) / mmpi * dpi), (int)(18 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec4 = new Rectangle((int)(116 / mmpi * dpi), (int)((102.7 + (i * a)) / mmpi * dpi), (int)(18 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec5 = new Rectangle((int)(133 / mmpi * dpi), (int)((102.7 + (i * a)) / mmpi * dpi), (int)(26.2 / mmpi * dpi), (int)(a / mmpi * dpi));
                    Rectangle rec6 = new Rectangle((int)(157.2 / mmpi * dpi), (int)((102.7 + (i * a)) / mmpi * dpi), (int)(29.9 / mmpi * dpi), (int)(a / mmpi * dpi));
                    stringFormat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(String.Format("{0:0}", i + 1), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rec1, stringFormat);
                    stringFormat.Alignment = StringAlignment.Near;
                    e.Graphics.DrawString(desc, new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rec2, stringFormat);
                    stringFormat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(dr[3].ToString(), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rec3, stringFormat);
                    e.Graphics.DrawString(dr[4].ToString(), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rec4, stringFormat);
                    stringFormat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[5].ToString())), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rec5, stringFormat);
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[6].ToString())), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, rec6, stringFormat);
                    i++;
                    //if (desc.Length > 40) i++;
                }

                conn.Close();
            }
            catch (Exception)
            {
                conn.Close();
            }

            //Footer
            Rectangle re1 = new Rectangle((int)(157.2 / mmpi * dpi), (int)(182.2 / mmpi * dpi), (int)(29.9 / mmpi * dpi), (int)(7.9 / mmpi * dpi));
            Rectangle re2 = new Rectangle((int)(157.2 / mmpi * dpi), (int)(190.1 / mmpi * dpi), (int)(29.9 / mmpi * dpi), (int)(7.7 / mmpi * dpi));
            Rectangle re3 = new Rectangle((int)(157.2 / mmpi * dpi), (int)(197.8 / mmpi * dpi), (int)(29.9 / mmpi * dpi), (int)(8.5 / mmpi * dpi));
            Rectangle re4 = new Rectangle((int)(17.5 / mmpi * dpi), (int)(197.8 / mmpi * dpi), (int)(123.6 / mmpi * dpi), (int)(8.5 / mmpi * dpi));
            stringFormat.Alignment = StringAlignment.Far;
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueBeforeTax)), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, re1, stringFormat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueTax)), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, re2, stringFormat);
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, re3, stringFormat);
            stringFormat.Alignment = StringAlignment.Near;
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font("AngsanaUPC", 14, FontStyle.Regular), Brushes.Black, re4, stringFormat);

            dpi = 96;
        }

        public void printBilling(String ID, String Date, String CustomerID, String CustomerName, String CustomerTaxID, String CustomerAddress,
            String CustomerTel, String CustomerFax, String CustomerEmail, String BillingBy, String BillingDate, String PS, 
            String ValueAfterTax, object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            String font = "AngsanaUPC";
            StringFormat strformat = new StringFormat();
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Far;
            Brush col = Brushes.LightCyan;

            //Header
            e.Graphics.DrawString("บริษัท โย โน ทูลส์ จำกัด (สำนักงานใหญ่)", new Font(font, 24, FontStyle.Bold), Brushes.Black, (int)(13.5 / mmpi * dpi), (int)(15.1 / mmpi * dpi));
            e.Graphics.DrawString("ต้นฉบับใบวางบิล", new Font(font, 24, FontStyle.Bold), Brushes.Black, (int)(151.3 / mmpi * dpi), (int)(15.1 / mmpi * dpi));
            e.Graphics.DrawString("108/314 หมู่ 5 ต.พันท้ายนรสิงส์ อ.เมืองสมุทรสาคร จ.สมุทรสาคร\nโทร.034-116655, 099-0568889 แฟ็กส์.034-116655"
                , new Font(font, 14, FontStyle.Regular), Brushes.Black, (int)(13.5 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("หน้า 1/1", new Font(font, 14, FontStyle.Regular), Brushes.Black, (int)(154 / mmpi * dpi), (int)(27 / mmpi * dpi));
            e.Graphics.DrawString("TEX :ID 0 1 2 5 5 6 0 0 0 0 5 9 0", new Font(font, 14, FontStyle.Regular), Brushes.Black, (int)(154 / mmpi * dpi), (int)(33.6 / mmpi * dpi));

            //Side Header
            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Near;
            Rectangle r1 = new Rectangle((int)(13.6 / mmpi * dpi), (int)((43) / mmpi * dpi), (int)(137.8 / mmpi * dpi), (int)(43.1 / mmpi * dpi));
            e.Graphics.DrawString("รหัสลูกค้า : " + CustomerID + "\n" + CustomerName + "\nที่อยู่ " + CustomerAddress
                + "\nโทร: " + CustomerTel + " แฟ็กส์ : " + CustomerFax + "\nTEX :ID " + CustomerTaxID + "\nหมายเหตุ " + PS
                , new Font(font, 14, FontStyle.Regular), Brushes.Black, r1, strformat);

            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Near;
            Rectangle r2 = new Rectangle((int)(154.1 / mmpi * dpi), (int)((43) / mmpi * dpi), (int)(46.6 / mmpi * dpi), (int)(43.1 / mmpi * dpi));
            e.Graphics.DrawString("เลขที่ใบวางบิล " + ID + "\nวันที่วางบิล " + BillingDate + "\nบันทึกโดย " + BillingBy
                + "\nวันที่บันทึก " + Date, new Font(font, 14, FontStyle.Regular), Brushes.Black, r2, strformat);

            //List Header
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle re1 = new Rectangle((int)(13.5 / mmpi * dpi), (int)((86) / mmpi * dpi), (int)(14.8 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle re2 = new Rectangle((int)(28.3 / mmpi * dpi), (int)((86) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle re3 = new Rectangle((int)(62.7 / mmpi * dpi), (int)((86) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle re4 = new Rectangle((int)(97.1 / mmpi * dpi), (int)((86) / mmpi * dpi), (int)(29.4 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle re5 = new Rectangle((int)(126.5 / mmpi * dpi), (int)((86) / mmpi * dpi), (int)(26.5 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle re6 = new Rectangle((int)(152.9 / mmpi * dpi), (int)((86) / mmpi * dpi), (int)(47.6 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            e.Graphics.DrawString("ลำดับที่", new Font(font, 14, FontStyle.Bold), Brushes.Black, re1, strformat);
            e.Graphics.DrawString("เลขที่ใบขาย", new Font(font, 14, FontStyle.Bold), Brushes.Black, re2, strformat);
            e.Graphics.DrawString("วันที่ขาย", new Font(font, 14, FontStyle.Bold), Brushes.Black, re3, strformat);
            e.Graphics.DrawString("วันที่ครบกำหนด", new Font(font, 14, FontStyle.Bold), Brushes.Black, re4, strformat);
            e.Graphics.DrawString("จำนวนเงิน", new Font(font, 14, FontStyle.Bold), Brushes.Black, re5, strformat);
            e.Graphics.DrawString("หมายเหตุ", new Font(font, 14, FontStyle.Bold), Brushes.Black, re6, strformat);

            int k = 0;

            //List
            try
            {

                conn.Open();


                SqlDataAdapter sda = new SqlDataAdapter("select * from tblBilling where Id like ('" + ID + "%') order by InvoiceID ASC", conn);

                DataTable dt = new DataTable();
                sda.Fill(dt);
                int i = 0;

                k = dt.Rows.Count;

                foreach (DataRow dr in dt.Rows)
                {
                    strformat.Alignment = StringAlignment.Center;
                    strformat.LineAlignment = StringAlignment.Center;
                    Rectangle rec1 = new Rectangle((int)(13.5 / mmpi * dpi), (int)((93.1 + (i * 6.3)) / mmpi * dpi), (int)(14.8 / mmpi * dpi), (int)(6.3 / mmpi * dpi));
                    Rectangle rec2 = new Rectangle((int)(28.3 / mmpi * dpi), (int)((93.1 + (i * 6.3)) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(6.3 / mmpi * dpi));
                    Rectangle rec3 = new Rectangle((int)(62.7 / mmpi * dpi), (int)((93.1 + (i * 6.3)) / mmpi * dpi), (int)(34.4 / mmpi * dpi), (int)(6.3 / mmpi * dpi));
                    Rectangle rec4 = new Rectangle((int)(97.1 / mmpi * dpi), (int)((93.1 + (i * 6.3)) / mmpi * dpi), (int)(29.4 / mmpi * dpi), (int)(6.3 / mmpi * dpi));
                    Rectangle rec5 = new Rectangle((int)(126.5 / mmpi * dpi), (int)((93.1 + (i * 6.3)) / mmpi * dpi), (int)(26.5 / mmpi * dpi), (int)(6.3 / mmpi * dpi));
                    Rectangle rec6 = new Rectangle((int)(152.9 / mmpi * dpi), (int)((93.1 + (i * 6.3)) / mmpi * dpi), (int)(47.6 / mmpi * dpi), (int)(6.3 / mmpi * dpi));
                    e.Graphics.DrawString(String.Format("{0:0}", i+1), new Font(font, 14, FontStyle.Regular), Brushes.Black, rec1, strformat);
                    e.Graphics.DrawString(dr[2].ToString(), new Font(font, 14, FontStyle.Regular), Brushes.Black, rec2, strformat);
                    e.Graphics.DrawString(dr[3].ToString(), new Font(font, 14, FontStyle.Regular), Brushes.Black, rec3, strformat);
                    e.Graphics.DrawString(dr[4].ToString(), new Font(font, 14, FontStyle.Regular), Brushes.Black, rec4, strformat);
                    strformat.Alignment = StringAlignment.Far;
                    e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(dr[5].ToString())), new Font(font, 14, FontStyle.Regular), Brushes.Black, rec5, strformat);
                    strformat.Alignment = StringAlignment.Center;
                    e.Graphics.DrawString(dr[6].ToString(), new Font(font, 14, FontStyle.Regular), Brushes.Black, rec6, strformat);
                    i++;

                }

                conn.Close();


            }
            catch (Exception)
            {
                conn.Close();
            }

            //Footer
            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Center;
            Rectangle rc1 = new Rectangle((int)(13.6 / mmpi * dpi), (int)((219.8) / mmpi * dpi), (int)(83.6 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle rc2 = new Rectangle((int)(97.1 / mmpi * dpi), (int)((219.8) / mmpi * dpi), (int)(29.4 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            Rectangle rc3 = new Rectangle((int)(126.5 / mmpi * dpi), (int)((219.8) / mmpi * dpi), (int)(26.5 / mmpi * dpi), (int)(7.1 / mmpi * dpi));
            e.Graphics.DrawString(ThaiBaht(ValueAfterTax), new Font(font, 14, FontStyle.Regular), Brushes.Black, rc1, strformat);
            strformat.Alignment = StringAlignment.Center;
            e.Graphics.DrawString("จำนวนเงินรวม", new Font(font, 14, FontStyle.Regular), Brushes.Black, rc2, strformat);
            strformat.Alignment = StringAlignment.Far;
            e.Graphics.DrawString(String.Format("{0:0,0.00}", float.Parse(ValueAfterTax)), new Font(font, 14, FontStyle.Regular), Brushes.Black, rc3, strformat);

            //Signature
            strformat.Alignment = StringAlignment.Near;
            strformat.LineAlignment = StringAlignment.Near;
            Rectangle rtg1 = new Rectangle((int)(13.5 / mmpi * dpi), (int)((227.1) / mmpi * dpi), (int)(123.8 / mmpi * dpi), (int)(40.7 / mmpi * dpi));
            Rectangle rtg2 = new Rectangle((int)(137.3 / mmpi * dpi), (int)((227.1) / mmpi * dpi), (int)(63.2 / mmpi * dpi), (int)(40.7 / mmpi * dpi));
            e.Graphics.DrawString("รวมทั้งสิ้น " + k + " ฉบับ\nชื่อผู้รับวางบิล.............................\n" +
                "วันที.......................................\nนัดรับเช็ค/โอนเงิน วันที่............................. เวลา.............................\n" +
                "หมายเหตุ.........................................................."
                , new Font(font, 14, FontStyle.Regular), Brushes.Black, rtg1, strformat);
            strformat.Alignment = StringAlignment.Center;
            strformat.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("ชื่อผู้วางบิล.............................................\nวันที่....../....../......"
                , new Font(font, 14, FontStyle.Regular), Brushes.Black, rtg2, strformat);

            //DrawLine

            /*Vertical*/
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(42.9 / mmpi * dpi), (int)(13.5 / mmpi * dpi), (int)(93.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(200.6 / mmpi * dpi), (int)(42.9 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(93.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(219.8 / mmpi * dpi), (int)(13.5 / mmpi * dpi), (int)(227 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(97.1 / mmpi * dpi), (int)(219.8 / mmpi * dpi), (int)(97.1 / mmpi * dpi), (int)(227 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(126.5 / mmpi * dpi), (int)(219.8 / mmpi * dpi), (int)(126.5 / mmpi * dpi), (int)(227 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(200.6 / mmpi * dpi), (int)(219.8 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(227 / mmpi * dpi));
            /*Horizintal*/
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(42.9 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(42.9 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(86 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(86 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(93.1 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(93.1 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(219.8 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(219.8 / mmpi * dpi));
            e.Graphics.DrawLine(Pens.Black, (int)(13.5 / mmpi * dpi), (int)(227 / mmpi * dpi), (int)(200.6 / mmpi * dpi), (int)(227 / mmpi * dpi));

        }




    }
}
