using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ProjectHouse
{
    class Connection
    {

        public static SqlConnection getConnection()
        {
            
            var database = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dtbDonkey.mdf");
            var connString = $"Data Source = (LocalDB)\\MSSQLLocalDB;AttachDbFilename={database};Integrated Security = True; Connect Timeout = 30";
            //MessageBox.Show(database);
            return new SqlConnection(connString);
            
            
            
            //return new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=\"C: \\Users\\Kirk Pig\\source\\repos\\ProjectHouse\\ProjectHouse\\dtbHouse.mdf\";Integrated Security=True");
        }

        public static String getLogo()
        {
            
            var pic = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "yono_logo.png");
            return pic;
            
            
            //return "C:\\Users\\Kirk Pig\\source\\repos\\ProjectHouse\\ProjectHouse\\yono_logo.png";
        }

    }
}
