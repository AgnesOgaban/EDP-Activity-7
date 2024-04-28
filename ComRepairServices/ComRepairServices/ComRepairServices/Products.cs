using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComRepairServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace ComRepairServices
{
    public partial class Products : Form
    {
        private List<Prod> product;

        public List<Prod> GetProduct1()
        {
            return product;
        }

        public void SetProduct1(List<Prod> value)
        {
            product = value;
        }

        public Products()
        {
            InitializeComponent();
            SetProduct1(GetProduct());
            dataGridProd.DataSource = GetProduct1();
        }

        private List<Prod> GetProduct()
        {
            var list = new List<Prod>();
            list.Add(new Prod()
            {
                ProductId = 1,
                ProductName = "Laptop",
                Description = "15.6-inch HD display, Intel Core i5 processor, 8GB RAM, 256GB SSD",
                Price = 40000.00m,  // Decimal value instead of string
                StockQuantity = 100   // Integer value instead of string
            });

            list.Add(new Prod()
            {
                ProductId = 2,
                ProductName = "Desktop PC",
                Description = "Intel Core i7 processor, 16GB RAM, 1TB HDD, NVIDIA GeForce GTX 1660",
                Price = 50000.00m,   // Decimal value instead of string
                StockQuantity = 50    // Integer value instead of string
            });

            list.Add(new Prod()
            {
                ProductId = 3,
                ProductName = "Printer",
                Description = "Wireless all-in-one printer with scanning and copying capabilities",
                Price = 10000.00m,   // Decimal value instead of string
                StockQuantity = 30    // Integer value instead of string
            });

            return list;
        }

        private void dataGridProd_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Your implementation
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            // Your implementation
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Create an instance of the ProductsForm class
            Products productsForm = new Products();

            // Show the ProductsForm
            productsForm.Show();
        }


        private void Products_Load(object sender, EventArgs e)
        {
            var Product = this.GetProduct();
            dataGridProd.DataSource = GetProduct();
        }

        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }




        //Exporting to Excel

        // Button click event handler for exporting report to Excel
        private void BtnExportProd_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridProd);
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
            worksheet.Cells[10, 2] = "PRODUCTS LIST";
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
            Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[dataGridProd.Rows.Count + 1, dataGridProd.Columns.Count]];
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