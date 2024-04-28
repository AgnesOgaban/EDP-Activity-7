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
    public partial class Dashboard : Form
    {
        public Dashboard()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }


        
        private void button2_Click(object sender, EventArgs e)
        {
            // Create an instance of the Employees form
            Employees employeesForm = new Employees();

            // Show the Employees form
            employeesForm.Show();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Create an instance of the Products form
            Products productsForm = new Products();

            // Show the Products form
            productsForm.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Create an instance of the Customers form
            Customers customersForm = new Customers();

            // Show the Customers form
            customersForm.Show();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Create an instance of the Transactions form
            Transactions transactionsForm = new Transactions();

            // Show the Transactions form
            transactionsForm.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {

            OpenReportsForm();
        }

        private void OpenReportsForm()
        {
            // Hide the current form
            this.Hide();

            // Create an instance of the Reports form
            Reports reportsForm = new Reports();

            // Show the Reports form
            reportsForm.Show();


        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenAboutForm();
        }
        private void OpenAboutForm()
        {
            // Hide the current form
            this.Hide();

            // Create an instance of the About form
            About aboutForm = new About();

            // Show the About form
            aboutForm.Show();
        }









        /*
        private void OpenForm<T>() where T : Form, new()
        {
            // Check if an instance of the form already exists
            Form existingForm = Application.OpenForms.OfType<T>().FirstOrDefault();

            if (existingForm != null)
            {
                // If the form is already open, bring it to the front
                existingForm.BringToFront();
            }
            else
            {
                // If the form is not open, create a new instance and show it
                T newForm = new T();
                newForm.Show();
            }
        }
    
        /* 
        private void button2_Click(object sender, EventArgs e)
        {
            OpenForm<Employees>();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenForm<Products>();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenForm<Customers>();
        }
        */







    }
}
