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
    public partial class Employees : Form
    {
        public List<Emp> Employee { get; set; }

        public Employees()
        {
            Employee = GetEmployee();
            InitializeComponent();
        }

        private List<Emp> GetEmployee()
        {
            var list = new List<Emp>();
            list.Add(new Emp()
            {
                EmployeeId = 1,
                Firstname = "Agnes",
                Lastname = "Ogaban",
                Email = "agnesogaban@gmail.com",
                PhoneNumber = "09928940225",
                Position = "Manager",
                HireDate = "2021-05-20",
                Salary = "50000"

            });


            list.Add(new Emp()
            {
                EmployeeId = 2,
                Firstname = "James",
                Lastname = "Fernandez",
                Email = "jamesfernandez@gmail.com",
                PhoneNumber = "09757545234",
                Position = "IT Technician",
                HireDate = "2023-04-16",
                Salary = "30000"

            });

            list.Add(new Emp()
            {
                EmployeeId = 3,
                Firstname = "Joshua",
                Lastname = "Jason",
                Email = "joshuajason@gmail.com",
                PhoneNumber = "09468734250",
                Position = "Hardware Technician",
                HireDate = "2022-06-25",
                Salary = "20000"

            });



            return list;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // Create an instance of the Employees form
            Employees employeesForm = new Employees();

            // Show the Employees form
            employeesForm.Show();
        }

        private void Panel2_Paint(object sender, PaintEventArgs e)
        {
            // Code for painting panel
        }


        private void DataGridEmp_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void Employees_Load(object sender, EventArgs e)
        {
            var Employee = this.Employee;
            DataGridEmp.DataSource = Employee;
        }


        //Exporting to Excel

        // Button click event handler for exporting report to Excel
        private void BtnExportEmp_Click(object sender, EventArgs e)
        {
            ExportToExcel(DataGridEmp);
        }

        // Method to export DataGridView data to Excel
        private void ExportToExcel(DataGridView dataGridView)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Add company logo at the header's left side
            AddCompanyLogo(worksheet);

            // Add company logo
            AddCompanyLogo(worksheet);

            // Merge cells for the company name
            Excel.Range companyNameRange = worksheet.Range["E4", "H4"];
            companyNameRange.Merge();
            companyNameRange.Value = "Computer Fix-IT-shop";
            companyNameRange.Font.Bold = true;
            companyNameRange.Font.Size = 16;
            companyNameRange.Font.Name = "Algerian"; // Set Algerian font
            companyNameRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Add "Employee List" above the headers
            worksheet.Cells[10, 2] = "EMPLOYEE LIST";
            // Merge cells for "Employee List" row
            Excel.Range employeeListRange = worksheet.Range[worksheet.Cells[10, 2], worksheet.Cells[10, dataGridView.Columns.Count + 1]];
            employeeListRange.Merge();
            // Format "Employee List" row
            employeeListRange.Font.Bold = true;
            employeeListRange.Font.Size = 14;
            employeeListRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
           


            // Add headers
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                worksheet.Cells[11, i + 2] = dataGridView.Columns[i].HeaderText;

                // Highlight with orange color
                worksheet.Cells[11, i + 2].Interior.Color = System.Drawing.Color.Orange;

                // Center the headers
                worksheet.Cells[11, i + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Set font to 12pt bold
                worksheet.Cells[11, i + 2].Font.Size = 12;
                worksheet.Cells[11, i + 2].Font.Bold = true;
            }

            // Add data
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    worksheet.Cells[i + 12, j + 2] = dataGridView.Rows[i].Cells[j].Value?.ToString();

                    // Center the data
                    worksheet.Cells[i + 12, j + 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    // Set font to 12pt bold
                    worksheet.Cells[i + 12, j + 2].Font.Size = 12;
                    worksheet.Cells[i + 12, j + 2].Font.Bold = false;

                    // Apply light gray background color
                    worksheet.Cells[i + 12, j + 2].Interior.Color = System.Drawing.Color.LightGray;
                }
            }


            // Add signature placeholder
            Excel.Range signaturePlaceholder = worksheet.Range["C40:G40"];
            signaturePlaceholder.Merge();
            signaturePlaceholder.Value = "SIGNED BY: ______________________________";
            signaturePlaceholder.Font.Bold = true;
            signaturePlaceholder.Font.Size = 14;

            // Add "Agnes V. Ogaban" below the line
            Excel.Range nameRange = worksheet.Range["C41:G41"];
            nameRange.Merge();
            nameRange.Value = "AGNES V. OGABAN";
            nameRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            nameRange.Font.Bold = true;


            // Add "Manager" below the name
            Excel.Range titleRange = worksheet.Range["C42:G42"];
            titleRange.Merge();
            titleRange.Value = "MANAGER";
            titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRange.Font.Bold = true;



            // Add Sheet 2 with graph
            Excel.Worksheet worksheet2 = workbook.Sheets.Add();
            worksheet2.Name = "Graph";
            Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet2.ChartObjects(Type.Missing);
            Excel.ChartObject chartObject = chartObjects.Add(50, 50, 300, 300);
            Excel.Chart chart = chartObject.Chart;
            Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[DataGridEmp.Rows.Count + 1, DataGridEmp.Columns.Count]];
            chart.SetSourceData(range, Type.Missing);
            chart.ChartType = Excel.XlChartType.xlColumnClustered;





            // Save Excel file
            try
            {
                // Prompt user to choose the location
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog1.Title = "Save Excel File";
                saveFileDialog1.ShowDialog();

                // If the file name is not empty, save the file
                if (saveFileDialog1.FileName != "")
                {
                    workbook.SaveAs(saveFileDialog1.FileName);
                    MessageBox.Show("Export Successful", $"Report exported successfully to {saveFileDialog1.FileName}");
                }
                else
                {
                    MessageBox.Show("Export Cancelled", "No file selected");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export Failed", $"An error occurred: {ex.Message}");
            }
            finally
            {
                workbook.Close();
                excelApp.Quit();
            }
        }


        private void AddCompanyLogo(Excel.Worksheet worksheet)
        {
            Excel.Pictures pictures = worksheet.Pictures(System.Reflection.Missing.Value) as Excel.Pictures;
            Excel.Picture picture = pictures.Insert(@"C:\Users\Agnes Ogaban\Downloads\ComRepairServices\ComRepairServices\logo.png", System.Reflection.Missing.Value);
            picture.Left = Convert.ToDouble(worksheet.Cells[1, 1].Left);
            picture.Top = Convert.ToDouble(worksheet.Cells[1, 1].Top);
            picture.Width = 80; // Set the width as needed
            picture.Height = 80; // Set the height as needed
        }


    }
}
