using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ComRepairServices
{
    partial class PassRecovery : Form
    {
        // Define user credentials (username, password, security question, and answer)
        private static readonly Dictionary<string, (string password, string securityQuestion, string securityAnswer)> UserCredentials = new Dictionary<string, (string, string, string)>
        {
            {"Ayaka28", ("kamisato", "Favorite Car?", "Nissan")},
            {"root", ("shin29", "Favorite Car?", "toyota")},
            {"Ayato26", ("Kamisato01", "Favorite Car?", "BMW")}
        };

        public PassRecovery()
        {
            InitializeComponent();
        }

        private void reset_Click(object sender, EventArgs e)
        {
            string username = userN.Text.Trim();
            string answer = securityQ.Text.Trim();

            // Check if the username exists in the dictionary
            if (UserCredentials.ContainsKey(username))
            {
                var (password, securityQuestion, securityAnswer) = UserCredentials[username];

                // Check if the entered answer matches the stored security answer
                if (answer.ToLower() == securityAnswer.ToLower())
                {
                    MessageBox.Show($"Username: {username}\nPassword: {password}", "Password Recovery", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Close the current form (PassRecovery)
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Incorrect security answer. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Unknown username. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PassRecovery_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Open the login form
            Login loginForm = new Login();
            loginForm.Show();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            // Close the current form
            this.Close();

            // Show the login form
            Login loginForm = new Login();
            loginForm.Show();
        }


        private void emailadd_TextChanged(object sender, EventArgs e)
        {
            // Add any necessary event handling code
        }
    }
}
