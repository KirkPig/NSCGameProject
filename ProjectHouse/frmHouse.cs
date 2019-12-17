using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectHouse
{
    public partial class frmHouse : MetroFramework.Forms.MetroForm
    {
        public frmHouse()
        {
            InitializeComponent();
        }

        private void MetroButton1_Click(object sender, EventArgs e)
        {
            Form f = new frmProduct();
            f.Show();
        }

        private void MetroButton2_Click(object sender, EventArgs e)
        {
            Form f = new frmCustomer();
            f.Show();
        }

        private void FrmHouse_Load(object sender, EventArgs e)
        {
            metroTabControl1.SelectedTab = metroTabPage1;
            metroTabControl2.SelectedTab = metroTabPage9;
            metroTabControl3.SelectedTab = metroTabPage29;
            metroTabControl5.SelectedTab = metroTabPage14;
            metroTabControl6.SelectedTab = metroTabPage15;
            metroTabControl7.SelectedTab = metroTabPage26;
            metroTabControl8.SelectedTab = metroTabPage27;
            metroTabControl9.SelectedTab = metroTabPage28;
        }
    }
}
