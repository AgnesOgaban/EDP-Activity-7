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
    public partial class About: Form
    {
        public About()
        {
            InitializeComponent();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Call the new function to perform the desired functionality
            PerformButtonClickFunction();
        }

        // Add a new private function to encapsulate the desired functionality
        private void PerformButtonClickFunction()
        {
            // Add your code here to perform the desired functionality
            // For example:
            MessageBox.Show("Button 7 clicked!");
        }
    }
}
