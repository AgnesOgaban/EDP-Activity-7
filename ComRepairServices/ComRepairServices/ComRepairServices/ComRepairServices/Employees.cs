using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ComRepairServices
{
    public partial class Employees : Form
    {
        public Employees()
        {
            InitializeComponent();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Create an instance of the Employees form
            Employees employeesForm = new Employees();

            // Show the Employees form
            employeesForm.Show();
        }
    }
}
