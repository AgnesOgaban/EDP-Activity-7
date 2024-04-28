using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ComRepairServices
{
    partial class Reports
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.button10 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.Accountbttn = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btnExportRep3 = new System.Windows.Forms.Button();
            this.btnExportRep2 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dataGridRep3 = new System.Windows.Forms.DataGridView();
            this.dataGridRep2 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.btnExportRep1 = new System.Windows.Forms.Button();
            this.dataGridRep1 = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridRep3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridRep2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridRep1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(0, 1);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(232, 455);
            this.panel1.TabIndex = 0;
            // 
            // button10
            // 
            this.button10.BackColor = System.Drawing.Color.PaleGreen;
            this.button10.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button10.Location = new System.Drawing.Point(390, 522);
            this.button10.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(65, 30);
            this.button10.TabIndex = 18;
            this.button10.Text = "Delete";
            this.button10.UseVisualStyleBackColor = false;
            this.button10.Click += new System.EventHandler(this.button9_Click);
            // 
            // button9
            // 
            this.button9.BackColor = System.Drawing.Color.PaleGreen;
            this.button9.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button9.Location = new System.Drawing.Point(305, 522);
            this.button9.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(67, 30);
            this.button9.TabIndex = 17;
            this.button9.Text = "Edit";
            this.button9.UseVisualStyleBackColor = false;
            // 
            // button8
            // 
            this.button8.BackColor = System.Drawing.Color.PaleGreen;
            this.button8.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button8.Location = new System.Drawing.Point(220, 522);
            this.button8.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(67, 30);
            this.button8.TabIndex = 16;
            this.button8.Text = "Save";
            this.button8.UseVisualStyleBackColor = false;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button7
            // 
            this.button7.BackColor = System.Drawing.Color.PaleGreen;
            this.button7.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button7.Location = new System.Drawing.Point(49, 380);
            this.button7.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(123, 30);
            this.button7.TabIndex = 14;
            this.button7.Text = "Abouts";
            this.button7.UseVisualStyleBackColor = false;
            // 
            // button6
            // 
            this.button6.BackColor = System.Drawing.Color.PaleGreen;
            this.button6.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button6.Location = new System.Drawing.Point(49, 345);
            this.button6.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(123, 30);
            this.button6.TabIndex = 13;
            this.button6.Text = "Reports";
            this.button6.UseVisualStyleBackColor = false;
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.Color.PaleGreen;
            this.button5.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.Location = new System.Drawing.Point(49, 310);
            this.button5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(123, 30);
            this.button5.TabIndex = 12;
            this.button5.Text = "Transactions";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.BackColor = System.Drawing.Color.PaleGreen;
            this.button4.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(49, 274);
            this.button4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(123, 30);
            this.button4.TabIndex = 11;
            this.button4.Text = "Customers";
            this.button4.UseVisualStyleBackColor = false;
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.PaleGreen;
            this.button3.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(49, 238);
            this.button3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(123, 30);
            this.button3.TabIndex = 10;
            this.button3.Text = "Products";
            this.button3.UseVisualStyleBackColor = false;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.PaleGreen;
            this.button2.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(49, 203);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(123, 30);
            this.button2.TabIndex = 9;
            this.button2.Text = "Employees";
            this.button2.UseVisualStyleBackColor = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.PaleGreen;
            this.button1.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(49, 167);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(123, 30);
            this.button1.TabIndex = 8;
            this.button1.Text = "Dashboard";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.SlateBlue;
            this.panel2.Controls.Add(this.Accountbttn);
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Controls.Add(this.button7);
            this.panel2.Controls.Add(this.button6);
            this.panel2.Controls.Add(this.button5);
            this.panel2.Controls.Add(this.button4);
            this.panel2.Controls.Add(this.button3);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.panel4);
            this.panel2.Location = new System.Drawing.Point(-4, -4);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(910, 575);
            this.panel2.TabIndex = 1;
            // 
            // Accountbttn
            // 
            this.Accountbttn.BackColor = System.Drawing.Color.PaleGreen;
            this.Accountbttn.Font = new System.Drawing.Font("Book Antiqua", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Accountbttn.Location = new System.Drawing.Point(49, 415);
            this.Accountbttn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Accountbttn.Name = "Accountbttn";
            this.Accountbttn.Size = new System.Drawing.Size(123, 30);
            this.Accountbttn.TabIndex = 16;
            this.Accountbttn.Text = "Account";
            this.Accountbttn.UseVisualStyleBackColor = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ComRepairServices.Properties.Resources.Screenshot__2873_;
            this.pictureBox1.Location = new System.Drawing.Point(12, 17);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(189, 127);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 15;
            this.pictureBox1.TabStop = false;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.Beige;
            this.panel4.Controls.Add(this.btnExportRep3);
            this.panel4.Controls.Add(this.btnExportRep2);
            this.panel4.Controls.Add(this.label4);
            this.panel4.Controls.Add(this.label3);
            this.panel4.Controls.Add(this.dataGridRep3);
            this.panel4.Controls.Add(this.dataGridRep2);
            this.panel4.Controls.Add(this.label2);
            this.panel4.Controls.Add(this.btnExportRep1);
            this.panel4.Controls.Add(this.button8);
            this.panel4.Controls.Add(this.button10);
            this.panel4.Controls.Add(this.button9);
            this.panel4.Controls.Add(this.dataGridRep1);
            this.panel4.Controls.Add(this.label1);
            this.panel4.Controls.Add(this.pictureBox5);
            this.panel4.Location = new System.Drawing.Point(207, 12);
            this.panel4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(700, 554);
            this.panel4.TabIndex = 1;
            this.panel4.Paint += new System.Windows.Forms.PaintEventHandler(this.panel4_Paint);
            // 
            // btnExportRep3
            // 
            this.btnExportRep3.BackColor = System.Drawing.Color.PaleGreen;
            this.btnExportRep3.Location = new System.Drawing.Point(618, 432);
            this.btnExportRep3.Name = "btnExportRep3";
            this.btnExportRep3.Size = new System.Drawing.Size(79, 35);
            this.btnExportRep3.TabIndex = 31;
            this.btnExportRep3.Text = "Export";
            this.btnExportRep3.UseVisualStyleBackColor = false;
            this.btnExportRep3.Click += new System.EventHandler(this.btnExportRep3_Click);
            // 
            // btnExportRep2
            // 
            this.btnExportRep2.BackColor = System.Drawing.Color.PaleGreen;
            this.btnExportRep2.Location = new System.Drawing.Point(618, 262);
            this.btnExportRep2.Name = "btnExportRep2";
            this.btnExportRep2.Size = new System.Drawing.Size(79, 35);
            this.btnExportRep2.TabIndex = 30;
            this.btnExportRep2.Text = "Export";
            this.btnExportRep2.UseVisualStyleBackColor = false;
            this.btnExportRep2.Click += new System.EventHandler(this.btnExportRep2_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 362);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(168, 22);
            this.label4.TabIndex = 29;
            this.label4.Text = "Repair Status update";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(15, 198);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 22);
            this.label3.TabIndex = 28;
            this.label3.Text = "Repair Request";
            // 
            // dataGridRep3
            // 
            this.dataGridRep3.BackgroundColor = System.Drawing.Color.White;
            this.dataGridRep3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridRep3.Location = new System.Drawing.Point(16, 387);
            this.dataGridRep3.Name = "dataGridRep3";
            this.dataGridRep3.RowHeadersWidth = 51;
            this.dataGridRep3.RowTemplate.Height = 24;
            this.dataGridRep3.Size = new System.Drawing.Size(596, 130);
            this.dataGridRep3.TabIndex = 27;
            // 
            // dataGridRep2
            // 
            this.dataGridRep2.BackgroundColor = System.Drawing.Color.White;
            this.dataGridRep2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridRep2.Location = new System.Drawing.Point(16, 223);
            this.dataGridRep2.Name = "dataGridRep2";
            this.dataGridRep2.RowHeadersWidth = 51;
            this.dataGridRep2.RowTemplate.Height = 24;
            this.dataGridRep2.Size = new System.Drawing.Size(596, 128);
            this.dataGridRep2.TabIndex = 26;
            this.dataGridRep2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridRep2_CellContentClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(121, 22);
            this.label2.TabIndex = 25;
            this.label2.Text = "Repair History";
            // 
            // btnExportRep1
            // 
            this.btnExportRep1.BackColor = System.Drawing.Color.PaleGreen;
            this.btnExportRep1.Location = new System.Drawing.Point(618, 111);
            this.btnExportRep1.Name = "btnExportRep1";
            this.btnExportRep1.Size = new System.Drawing.Size(79, 35);
            this.btnExportRep1.TabIndex = 24;
            this.btnExportRep1.Text = "Export";
            this.btnExportRep1.UseVisualStyleBackColor = false;
            this.btnExportRep1.Click += new System.EventHandler(this.btnExportRep1_Click);
            // 
            // dataGridRep1
            // 
            this.dataGridRep1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridRep1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridRep1.Location = new System.Drawing.Point(16, 66);
            this.dataGridRep1.Name = "dataGridRep1";
            this.dataGridRep1.RowHeadersWidth = 51;
            this.dataGridRep1.RowTemplate.Height = 24;
            this.dataGridRep1.Size = new System.Drawing.Size(596, 129);
            this.dataGridRep1.TabIndex = 23;
            this.dataGridRep1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Book Antiqua", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(437, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(157, 22);
            this.label1.TabIndex = 22;
            this.label1.Text = "Reports Generation";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // pictureBox5
            // 
            this.pictureBox5.Image = global::ComRepairServices.Properties.Resources.laptop_logo1;
            this.pictureBox5.Location = new System.Drawing.Point(629, 14);
            this.pictureBox5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(48, 35);
            this.pictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox5.TabIndex = 21;
            this.pictureBox5.TabStop = false;
            // 
            // Reports
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(912, 573);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Reports";
            this.Text = "Reports";
            this.Load += new System.EventHandler(this.Reports_Load);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridRep3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridRep2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridRep1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            this.ResumeLayout(false);

        }



        private void button9_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

       

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.PictureBox pictureBox5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Accountbttn;
        private System.Windows.Forms.DataGridView dataGridRep1;
        private Button btnExportRep1;
        private Label label2;
        private Label label4;
        private Label label3;
        private DataGridView dataGridRep3;
        private DataGridView dataGridRep2;
        private Button btnExportRep3;
        private Button btnExportRep2;
    }
}
