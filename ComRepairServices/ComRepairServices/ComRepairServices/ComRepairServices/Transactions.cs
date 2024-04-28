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
    public partial class Transactions : Form
    {
        public Transactions()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            // Create a font and brush for the text
            Font font = new Font("Arial", 12);
            SolidBrush brush = new SolidBrush(Color.Black);

            // Specify the text to be printed
            string textToPrint = "Hello, this is a sample text for printing.";

            // Specify the position where the text should be printed
            PointF position = new PointF(100, 100);

            // Draw the text on the print page
            e.Graphics.DrawString(textToPrint, font, brush, position);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Create an instance of the Dashboard form
            Dashboard dashboardForm = new Dashboard();

            // Show the Dashboard form
            dashboardForm.Show();

            // Optionally, you can close the Transactions form if needed
            this.Close();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
