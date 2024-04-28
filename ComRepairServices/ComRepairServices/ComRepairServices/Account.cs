using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace ComRepairServices
{
    public partial class Account : Form
    {
        private readonly string connectionString = "Server=localhost;Port=3386;Database=fixitshop_db;Uid=root;Pwd=shin29;";

        public Account()
        {
            InitializeComponent();
            LoadActiveAccounts(); // Load active accounts by default
            // Subscribe the Click event of the searchu button to the searchu_Click event handler
            this.searchu.Click += new System.EventHandler(this.searchu_Click);
            // Subscribe the Click event of the add button to the add_Click event handler
            this.add.Click += new System.EventHandler(this.add_Click);
            // Subscribe the Click event of the update button to the update_Click event handler
            this.update.Click += new System.EventHandler(this.update_Click);
            // Subscribe the Click event of the delete button to the delete_Click event handler
            this.delete.Click += new System.EventHandler(this.delete_Click);
        }

        private void LoadActiveAccounts()
        {
            string sql = "SELECT * FROM users WHERE status = 'Active'"; // Assuming 'status' field indicates account status

            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                    {
                        using (MySqlDataAdapter adapter = new MySqlDataAdapter(cmd))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            dataA.DataSource = dataTable;
                        }
                    }
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error loading active accounts: " + ex.Message);
            }
        }

        private void AddAccount(string username, string password)
        {
            string sql = "INSERT INTO users (username, password, status) VALUES (@username, @password, 'Active')";

            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@username", username);
                        cmd.Parameters.AddWithValue("@password", password);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Account added successfully.");
                LoadActiveAccounts(); // Refresh DataGridView after adding account
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error adding account: " + ex.Message);
            }
        }

        private void UpdateAccount(int userId, string username, string password)
        {
            string sql = "UPDATE users SET username = @username, password = @password WHERE id = @id";

            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@id", userId);
                        cmd.Parameters.AddWithValue("@username", username);
                        cmd.Parameters.AddWithValue("@password", password);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Account updated successfully.");
                LoadActiveAccounts(); // Refresh DataGridView after updating account
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error updating account: " + ex.Message);
            }
        }

        private void DeleteAccount(int userId)
        {
            string sql = "DELETE FROM users WHERE id = @id";

            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@id", userId);
                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Account deleted successfully.");
                LoadActiveAccounts(); // Refresh DataGridView after deleting account
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error deleting account: " + ex.Message);
            }
        }
        private void SearchAccounts(string username)
        {
            string sql = "SELECT * FROM users WHERE username = @username";

            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectionString))
                {
                    conn.Open();
                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@username", username);
                        using (MySqlDataAdapter adapter = new MySqlDataAdapter(cmd))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);
                            dataA.DataSource = dataTable;
                        }
                    }
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error searching accounts: " + ex.Message);
            }
        }

        private void searchu_Click(object sender, EventArgs e)
        {
            string username = userbox.Text;
            SearchAccounts(username);
        }

        private void add_Click(object sender, EventArgs e)
        {
            string username = userbox.Text;
            string password = passbox.Text;
            AddAccount(username, password);
        }

        private void update_Click(object sender, EventArgs e)
        {
            if (dataA.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dataA.SelectedRows[0].Index;
                int userId = Convert.ToInt32(dataA.Rows[selectedRowIndex].Cells["id"].Value);
                string username = userbox.Text;
                string password = passbox.Text;
                UpdateAccount(userId, username, password);
            }
            else
            {
                MessageBox.Show("Please select a row to update.");
            }
        }

        private void delete_Click(object sender, EventArgs e)
        {
            if (dataA.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dataA.SelectedRows[0].Index;
                int userId = Convert.ToInt32(dataA.Rows[selectedRowIndex].Cells["id"].Value);
                DeleteAccount(userId);
            }
            else
            {
                MessageBox.Show("Please select a row to delete.");
            }
        }

        private void Account_Load(object sender, EventArgs e)
        {

        }
    }
}
