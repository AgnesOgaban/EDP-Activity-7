namespace ComRepairServices
{
    partial class PassRecovery
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.reset = new System.Windows.Forms.Button();
            this.userN = new System.Windows.Forms.TextBox();
            this.Username = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.details = new System.Windows.Forms.Label();
            this.backlogin = new System.Windows.Forms.Button();
            this.securityQ = new System.Windows.Forms.TextBox();
            this.FavC = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // reset
            // 
            this.reset.BackColor = System.Drawing.Color.SlateBlue;
            this.reset.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.reset.ForeColor = System.Drawing.Color.White;
            this.reset.Location = new System.Drawing.Point(186, 207);
            this.reset.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.reset.Name = "reset";
            this.reset.Size = new System.Drawing.Size(134, 27);
            this.reset.TabIndex = 16;
            this.reset.Text = "Confirm";
            this.reset.UseVisualStyleBackColor = false;
            this.reset.Click += new System.EventHandler(this.reset_Click);
            // 
            // userN
            // 
            this.userN.Location = new System.Drawing.Point(147, 113);
            this.userN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.userN.Multiline = true;
            this.userN.Name = "userN";
            this.userN.Size = new System.Drawing.Size(216, 28);
            this.userN.TabIndex = 13;
            this.userN.TextChanged += new System.EventHandler(this.emailadd_TextChanged);
            // 
            // Username
            // 
            this.Username.AutoSize = true;
            this.Username.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Username.Location = new System.Drawing.Point(144, 89);
            this.Username.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Username.Name = "Username";
            this.Username.Size = new System.Drawing.Size(111, 18);
            this.Username.TabIndex = 12;
            this.Username.Text = "Enter Username";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.SlateBlue;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(1, -2);
            this.panel1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(128, 285);
            this.panel1.TabIndex = 10;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ComRepairServices.Properties.Resources.Screenshot__2873_;
            this.pictureBox1.Location = new System.Drawing.Point(11, 91);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(84, 54);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Book Antiqua", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(168, 38);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(175, 20);
            this.label3.TabIndex = 21;
            this.label3.Text = "Forgot your Password?";
            // 
            // details
            // 
            this.details.AutoSize = true;
            this.details.Font = new System.Drawing.Font("Book Antiqua", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.details.Location = new System.Drawing.Point(145, 68);
            this.details.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.details.Name = "details";
            this.details.Size = new System.Drawing.Size(224, 15);
            this.details.TabIndex = 22;
            this.details.Text = "Please enter the username you use to sign in";
            // 
            // backlogin
            // 
            this.backlogin.BackColor = System.Drawing.Color.SlateBlue;
            this.backlogin.Font = new System.Drawing.Font("Book Antiqua", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.backlogin.ForeColor = System.Drawing.Color.White;
            this.backlogin.Location = new System.Drawing.Point(215, 241);
            this.backlogin.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.backlogin.Name = "backlogin";
            this.backlogin.Size = new System.Drawing.Size(80, 27);
            this.backlogin.TabIndex = 23;
            this.backlogin.Text = "Back To Login";
            this.backlogin.UseVisualStyleBackColor = false;
            this.backlogin.Click += new System.EventHandler(this.button2_Click);
            // 
            // securityQ
            // 
            this.securityQ.Location = new System.Drawing.Point(147, 171);
            this.securityQ.Margin = new System.Windows.Forms.Padding(2);
            this.securityQ.Multiline = true;
            this.securityQ.Name = "securityQ";
            this.securityQ.Size = new System.Drawing.Size(216, 28);
            this.securityQ.TabIndex = 24;
            // 
            // FavC
            // 
            this.FavC.AutoSize = true;
            this.FavC.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FavC.Location = new System.Drawing.Point(144, 149);
            this.FavC.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.FavC.Name = "FavC";
            this.FavC.Size = new System.Drawing.Size(88, 18);
            this.FavC.TabIndex = 25;
            this.FavC.Text = "Favorite Car";
            // 
            // PassRecovery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(379, 279);
            this.Controls.Add(this.FavC);
            this.Controls.Add(this.securityQ);
            this.Controls.Add(this.backlogin);
            this.Controls.Add(this.details);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.reset);
            this.Controls.Add(this.userN);
            this.Controls.Add(this.Username);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "PassRecovery";
            this.Text = "PassRecovery";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button reset;
        private System.Windows.Forms.TextBox userN;
        private System.Windows.Forms.Label Username;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label details;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button backlogin;
        private System.Windows.Forms.TextBox securityQ;
        private System.Windows.Forms.Label FavC;
    }
}