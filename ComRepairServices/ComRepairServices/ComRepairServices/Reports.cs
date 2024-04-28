using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ComRepairServices
{
    public partial class Reports : Form
    {
        public Reports()
        {
            InitializeComponent();
            Report = GetReport(); // Initialize the Report property with sample data
        }

        public List<RepairHistory> Report { get; set; } // Property to hold report data

        private List<RepairHistory> GetReport()
        {
            var list = new List<RepairHistory>();

            list.Add(new RepairHistory()
            {
                RequestId = 1,
                CustomerName = "John Doe",
                DeviceBrand = "HP",
                DeviceModel = "Pavilion",
                IssueDescription = "Screen cracked",
                RequestDate = DateTime.Now // Assigning DateTime.Now directly
            });

            list.Add(new RepairHistory()
            {
                RequestId = 2,
                CustomerName = "Jane Smith",
                DeviceBrand = "Dell",
                DeviceModel = "Inspiron",
                IssueDescription = "Battery not charging",
                RequestDate = DateTime.Now // Assigning DateTime.Now directly
            });

            list.Add(new RepairHistory()
            {
                RequestId = 3,
                CustomerName = "Michael Johnson",
                DeviceBrand = "Apple",
                DeviceModel = "MacBook Pro",
                IssueDescription = "Keyboard not working",
                RequestDate = DateTime.Now // Assigning DateTime.Now directly
            });

            return list;
        }

        private List<RepairRequest> GetRepairRequestReport()
        {
            var list = new List<RepairRequest>();

            // Populate RepairRequest report
            list.Add(new RepairRequest()
            {
                RequestId = 1,
                CustomerName = "John Doe",
                DeviceBrand = "HP",
                DeviceModel = "Pavilion",
                IssueDescription = "Screen cracked",
                RequestDate = DateTime.Now // Assigning DateTime.Now directly
            });

            list.Add(new RepairRequest()
            {
                RequestId = 2,
                CustomerName = "Alice Johnson",
                DeviceBrand = "Apple",
                DeviceModel = "iPhone 12",
                IssueDescription = "Speaker not working",
                RequestDate = DateTime.Now.AddDays(-2) // Example: 2 days ago
            });

            list.Add(new RepairRequest()
            {
                RequestId = 3,
                CustomerName = "Bob Smith",
                DeviceBrand = "Samsung",
                DeviceModel = "Galaxy S20",
                IssueDescription = "Camera lens cracked",
                RequestDate = DateTime.Now.AddDays(-5) // Example: 5 days ago
            });

            // Add more RepairRequest items if needed

            return list;
        }


        private List<RepairStatusUpdate> GetRepairStatusUpdateReport()
        {
            var list = new List<RepairStatusUpdate>();

            // Populate RepairStatusUpdate report
            list.Add(new RepairStatusUpdate()
            {
                RequestId = 1,
                NewStatus = "In Progress",
                UpdateDate = DateTime.Now
            });

            list.Add(new RepairStatusUpdate()
            {
                RequestId = 2,
                NewStatus = "On Hold",
                UpdateDate = DateTime.Now.AddDays(-3) // Example: 3 days ago
            });

            list.Add(new RepairStatusUpdate()
            {
                RequestId = 3,
                NewStatus = "Completed",
                UpdateDate = DateTime.Now.AddDays(-7) // Example: 7 days ago
            });

            // Add more RepairStatusUpdate items if needed

            return list;
        }


        //Exporting to Excel

        // Button click event handler for exporting report to Excel
        private void btnExportRep1_Click(object sender, EventArgs e)
        {
            ExportToExcel1(dataGridRep1);
        }

        // Method to export DataGridView data to Excel
        private void ExportToExcel1(DataGridView dataGridView)
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
            worksheet.Cells[10, 2] = "REPORT LIST - REPAIR HISTORY";
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
            Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[dataGridRep1.Rows.Count + 1, dataGridRep1.Columns.Count]];
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


        private void AddCompanyLogo3(Excel.Worksheet worksheet)
        {
            Excel.Pictures pictures = worksheet.Pictures(System.Reflection.Missing.Value) as Excel.Pictures;
            Excel.Picture picture = pictures.Insert(@"C:\Users\Agnes Ogaban\Downloads\ComRepairServices\ComRepairServices\logo.png", System.Reflection.Missing.Value);
            picture.Left = Convert.ToDouble(worksheet.Cells[1, 1].Left);
            picture.Top = Convert.ToDouble(worksheet.Cells[1, 1].Top);
            picture.Width = 80; // Set the width as needed
            picture.Height = 80; // Set the height as needed
        }













        //Exporting to Excel

        // Button click event handler for exporting report to Excel
        private void btnExportRep2_Click(object sender, EventArgs e)
        {
            ExportToExcel2(dataGridRep2);
        }

        // Method to export DataGridView data to Excel
        private void ExportToExcel2(DataGridView dataGridView)
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
            worksheet.Cells[10, 2] = "REPORT LIST - REPAIR REQUEST";
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
            Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[dataGridRep2.Rows.Count + 1, dataGridRep2.Columns.Count]];
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


        private void AddCompanyLogo2(Excel.Worksheet worksheet)
        {
            Excel.Pictures pictures = worksheet.Pictures(System.Reflection.Missing.Value) as Excel.Pictures;
            Excel.Picture picture = pictures.Insert(@"C:\Users\Agnes Ogaban\Downloads\ComRepairServices\ComRepairServices\logo.png", System.Reflection.Missing.Value);
            picture.Left = Convert.ToDouble(worksheet.Cells[1, 1].Left);
            picture.Top = Convert.ToDouble(worksheet.Cells[1, 1].Top);
            picture.Width = 80; // Set the width as needed
            picture.Height = 80; // Set the height as needed
        }







        //Exporting to Excel

        // Button click event handler for exporting report to Excel
        private void btnExportRep3_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridRep3);
        }

        // Method to export DataGridView data to Excel
        private void ExportToExcel(DataGridView dataGridView)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Add company logo at the header's left side
            AddCompanyLogo3(worksheet);

            // Add company logo
            AddCompanyLogo3(worksheet);

            // Merge cells for the company name
            Excel.Range companyNameRange = worksheet.Range["E4", "H4"];
            companyNameRange.Merge();
            companyNameRange.Value = "Computer Fix-IT-shop";
            companyNameRange.Font.Bold = true;
            companyNameRange.Font.Size = 16;
            companyNameRange.Font.Name = "Algerian"; // Set Algerian font
            companyNameRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Add "Employee List" above the headers
            worksheet.Cells[10, 2] = "REPORT LIST - REPAIR STATUS UPDATE";
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
            Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[dataGridRep3.Rows.Count + 1, dataGridRep3.Columns.Count]];
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




        private void Reports_Load(object sender, EventArgs e)
        {
            var report1 = GetReport();
            dataGridRep1.DataSource = report1;

            var report2 = GetRepairRequestReport();
            dataGridRep2.DataSource = report2;

            var report3 = GetRepairStatusUpdateReport();
            dataGridRep3.DataSource = report3;
        }



        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }


        private void btnExportEmp_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void PopulateSampleData(DataGridView dataGridView2, object value)
        {
            throw new NotImplementedException();
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridRep2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

    
    }
}


