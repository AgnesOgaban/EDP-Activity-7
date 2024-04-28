using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComRepairServices
{
    public partial class Transactions : Form
    {
        public List<Transact> Transaction { get; set; }

        public Transactions()
        {
            Transaction = Transaction1;
            InitializeComponent();

        }

        private List<Transact> Transaction1
        {
            get
            {
                var list = new List<Transact>();
                list.Add(new Transact()
                {
                    TransactionId = 1,
                    CustomerId = "1",
                    EmployeeId = "101",
                    TransactionDate = DateTime.Parse("2024-04-14 09:30:00"),
                    TotalAmount = 50m,
                    Description = "Purchase of electronics"
                });

                list.Add(new Transact()
                {
                    TransactionId = 2,
                    CustomerId = "2",
                    EmployeeId = "102",
                    TransactionDate = DateTime.Parse("2024-04-14 10:15:00"),
                    TotalAmount = 75.50m,
                    Description = "Payment for repair services"
                });

                list.Add(new Transact()
                {
                    TransactionId = 3,
                    CustomerId = "3",
                    EmployeeId = "103",
                    TransactionDate = DateTime.Parse("2024-04-14 11:00:00"),
                    TotalAmount = 30.25m,
                    Description = "Purchase of accessories"
                });

                list.Add(new Transact()
                {
                    TransactionId = 4,
                    CustomerId = "3",
                    EmployeeId = "103",
                    TransactionDate = DateTime.Parse("2024-04-14 11:00:00"),
                    TotalAmount = 30.25m,
                    Description = "Purchase of accessories"
                });

                list.Add(new Transact()
                {
                    TransactionId = 5,
                    CustomerId = "1",
                    EmployeeId = "104",
                    TransactionDate = DateTime.Parse("2024-04-14 12:00:00"),
                    TotalAmount = 120m,
                    Description = "Repair of laptop"
                });

                return list;
            }
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
            // Get the required values (example values shown here)
            string username = "root";
            string password = "shin29";
            string hostname = "localhost";
            int port = 3386;

            // Create an instance of the Dashboard form with the required arguments
            Dashboard dashboardForm = new Dashboard(username, password, hostname, port);

            // Show the Dashboard form
            dashboardForm.Show();

            // Optionally, you can close the Transactions form if needed
            this.Close();
        }


        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Transactions_Load(object sender, EventArgs e)
        {
            var Transaction = this.Transaction;
            dataGridTransact.DataSource = Transaction;
        }

        private void dataGridTransact_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private List<Transact> GetTransaction()
        {
            var list = new List<Transact>();
            list.Add(new Transact()
            {
                TransactionId = 1,
                CustomerId = "1",
                EmployeeId = "101",
                TransactionDate = DateTime.Parse("2024-04-14 09:30:00"),
                TotalAmount = 50m,
                Description = "Purchase of electronics"
            });

            list.Add(new Transact()
            {
                TransactionId = 2,
                CustomerId = "2",
                EmployeeId = "102",
                TransactionDate = DateTime.Parse("2024-04-14 10:15:00"),
                TotalAmount = 75.50m,
                Description = "Payment for repair services"
            });

            list.Add(new Transact()
            {
                TransactionId = 3,
                CustomerId = "3",
                EmployeeId = "103",
                TransactionDate = DateTime.Parse("2024-04-14 11:00:00"),
                TotalAmount = 30.25m,
                Description = "Purchase of accessories"
            });

            list.Add(new Transact()
            {
                TransactionId = 4,
                CustomerId = "3",
                EmployeeId = "103",
                TransactionDate = DateTime.Parse("2024-04-14 11:00:00"),
                TotalAmount = 30.25m,
                Description = "Purchase of accessories"
            });

            list.Add(new Transact()
            {
                TransactionId = 5,
                CustomerId = "1",
                EmployeeId = "104",
                TransactionDate = DateTime.Parse("2024-04-14 12:00:00"),
                TotalAmount = 120m,
                Description = "Repair of laptop"
            });

            return list;
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }
    }
}
