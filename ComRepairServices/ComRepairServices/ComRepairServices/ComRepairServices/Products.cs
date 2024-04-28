using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComRepairServices;


namespace ComRepairServices
{
    public partial class Products : Form
    {
        public Products()
        {
            InitializeComponent();
        }

        //Function 
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Create an instance of the ProductsForm class
            Products productsForm = new Products();

            // Show the ProductsForm
            productsForm.Show();
        }
    }
}
