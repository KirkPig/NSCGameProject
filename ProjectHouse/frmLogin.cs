using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ProjectHouse
{
    public partial class frmLogin : MetroFramework.Forms.MetroForm
    {

        System.Threading.Thread th;

        public frmLogin()
        {
            InitializeComponent();
        }

        private void MetroButton1_Click(object sender, EventArgs e)
        {
            if (metroTextBox1.Text == "admin" && metroTextBox2.Text == "password")
            {
                MessageBox.Show("Login Success!");
                
                
            }
            else
            {
                MessageBox.Show("Username or Password Incorrect");
            }
        }
    }
}
