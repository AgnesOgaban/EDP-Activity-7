using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using MySql.Data.MySqlClient;



namespace ComRepairServices
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string enteredUsername = usern_text.Text;
            string enteredPassword = pass_text.Text;
            string hostname = "localhost";
            int port = 3386;

            if (AuthenticateUser(enteredUsername, enteredPassword))
            {
                string connectionString = $"Server={hostname};Port={port};Database=fixitshop_db;Uid={enteredUsername};Pwd={enteredPassword};AllowPublicKeyRetrieval=true;SslMode=Preferred;";

                MySqlConnection connection = new MySqlConnection(connectionString);

                try
                {
                    connection.Open();
                    MessageBox.Show("Connected to MySQL server!");

                    // Close the current form (Login)
                    this.Hide();

                    // Show the Dashboard form
                    ShowDashboard(enteredUsername, enteredPassword, hostname, port);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error connecting to the database: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Invalid username or password. Please try again.");
            }
        }

        private bool AuthenticateUser(string username, string password)
        {
            // Replace this with your actual authentication logic
            // For simplicity, let's assume a username "admin" and password "password" for now
            return username == "admin" && password == "password";
        }

        private void usern_text_TextChanged(object sender, EventArgs e)
        {
            // You can add additional logic for username text changes if needed
        }

        private void pass_text_TextChanged(object sender, EventArgs e)
        {
            // You can add additional logic for password text changes if needed
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            // Add any necessary logic for the pictureBox1 click event
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            // Add any necessary logic for the pictureBox2 click event
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            // Add any necessary logic for the Remember Me checkbox
        }

        // Method to show the Dashboard form
        private void ShowDashboard(string username, string password, string hostname, int port)
        {
            // Implement the logic to show the Dashboard form
            // Example: Dashboard dashboardForm = new Dashboard(username, password, hostname, port);
            // dashboardForm.Show();
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Call the function to navigate to the forgot password recovery page
            NavigateToForgotPasswordRecovery();
        }

        private void NavigateToForgotPasswordRecovery()
        {
            // Assuming you have another form called ForgotPasswordRecoveryForm
            PassRecovery forgotPasswordForm = new PassRecovery();

            // Show the forgot password recovery form and hide the login form (if needed)
            forgotPasswordForm.Show();
            this.Hide();
        }
    }
}