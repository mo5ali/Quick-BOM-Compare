using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace Quick_BOM_Compare
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static string GenerateFilePath(string fileName)
        {
            // Check if the file name ends with .dft .asm or .par or other.
            if (fileName.EndsWith(".dft", StringComparison.OrdinalIgnoreCase))
            {
                return $"Z:\\Zeichnungen\\DFT\\{fileName}";
            }
            else if (fileName.EndsWith(".asm", StringComparison.OrdinalIgnoreCase))
            {
                return $"Z:\\Zeichnungen\\{fileName}";
            }
            else
            {
                return null;
            }
        }
        static bool CheckIfPathExists(string filePath)
        {
            if (File.Exists(filePath))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private string searchZeichnungVerknupfung(string asmName)
        {
            string ZeichnungVerknupfungPath = @"Z:\Allgemein\Hilfsmittel\ExcelMakros_NM\ZeichnungenZuKuehler_000.2.xlsm";
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; //license context for EPPlus 
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(asmName); //Remove the extension from the asmName

            // Step 4: Open the Excel file and search for the name in column A
            try
            {
                // Use EPPlus to load the Excel file
                FileInfo fileInfo = new FileInfo(ZeichnungVerknupfungPath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming first worksheet

                    // Get the total number of rows in the worksheet
                    int rowCount = worksheet.Dimension.Rows;

                    // Step 5: Loop through column A to search for the file name
                    for (int row = 1; row <= rowCount; row++)
                    {
                        string columnAValue = worksheet.Cells[row, 1].Text.Trim(); // Column A (1st column)

                        if (columnAValue == fileNameWithoutExtension)
                        {
                            // Step 6: First, check if column M is filled
                            string colM = worksheet.Cells[row, 13].Text; // Column M (13th column)
                            string colN = worksheet.Cells[row, 14].Text; // Column N (14th column)

                            if (!string.IsNullOrEmpty(colM))
                            {
                                return $"{colM}-{colN}"; // Return M-N if M is filled
                            }

                            // Step 7: If M is empty, check if column J is filled
                            string colJ = worksheet.Cells[row, 10].Text; // Column J (10th column)
                            string colK = worksheet.Cells[row, 11].Text; // Column K (11th column)

                            if (!string.IsNullOrEmpty(colJ))
                            {
                                return $"{colJ}-{colK}"; // Return J-K if J is filled
                            }

                            // Step 8: If J is empty, return the content of D and E columns
                            string colD = worksheet.Cells[row, 4].Text; // Column D (4th column)
                            string colE = worksheet.Cells[row, 5].Text; // Column E (5th column)

                            if (!string.IsNullOrEmpty(colD) && !string.IsNullOrEmpty(colE))
                            {
                                return $"{colD}-{colE}"; // Return D-E if D and E are filled
                            }

                            // Step 9: If J and M are empty, return this message
                            return $"No DFT linked to {fileNameWithoutExtension} in ZeichnungVerknupfung";
                        }
                    }
                }

                // If not found, return this message
                return "Not in ZeichnungVerknupfung";
            }
            catch (Exception ex)
            {
                // Handle any exceptions (e.g., file not found or Excel file errors)
                return $"Error: {ex.Message}";
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string InputAsmName = TxtAsmName.Text;  // Get the Assembly name
            string AsmFilePath = GenerateFilePath(InputAsmName); //Generate file path
            bool pathExists = CheckIfPathExists(AsmFilePath);
            if (pathExists == true)
            {
                LblPathStatus.Text = $"{AsmFilePath} ---> Exists :) ";
            }
            else if (pathExists == false)  
            {
                LblPathStatus.Text = $"{AsmFilePath} ---> Does not exist :(";
            }
            else
            {
                LblPathStatus.Text = $"Please do not forget the extension :* ";
            }

            // Step 5: Call the Excel search function and display the result in label4
            string ZeichnungVerknupfungResult = searchZeichnungVerknupfung(InputAsmName);
            label4.Text = ZeichnungVerknupfungResult; // Update label4 with the result from Excel search
        }

    
    }
}
