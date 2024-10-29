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
using SolidEdgeFramework;  // Ensure SolidEdgeFramework is included
using SolidEdgeAssembly;   // Ensure SolidEdgeAssembly is included
using SolidEdgeDraft;      // Ensure SolidEdgeDraft is included
using SolidEdgePart;       // Ensure SolidEdgePart is included
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace Quick_BOM_Compare
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private SolidEdgeFramework.SolidEdgeDocument document = null;
        private SolidEdgeFramework.Application application = null;

        bool IsThereALinkedDFT = false;
        public static bool Extracted_From_SAP = false;
        string Linked_DFT = string.Empty;
        string Linked_DFT_Path = string.Empty;
        string Linked_DFT_Path_alte = string.Empty;
        public static string model_used_in_DFT = string.Empty;
        public static Dictionary<string, double> Extracted_3D_List = new Dictionary<string, double>();
        public static Dictionary<string, double> Translated_3D_List = new Dictionary<string, double>();
        public static Dictionary<string, double> Extracted_SAP_List = new Dictionary<string, double>();
        public static string InputAsmName = string.Empty;
        public List<string> disregardList = new List<string> { "DB-ERP", "ESMP", "ESDT", "ECLIPS", "OIL IN", "OIL OUT", };

        private void LogMessage(string message)
        {
            if (LblCnsl.InvokeRequired)
            {
                LblCnsl.Invoke(new MethodInvoker(() => LogMessage(message)));
            }
            else
            {
                LblCnsl.Text += System.Environment.NewLine + message;
            }
        }
        public SolidEdgeFramework.Application InitializeSolidEdge()
        {
            SolidEdgeFramework.Application application = null;

            try
            {
                try
                {
                    // Try to get an existing Solid Edge instance
                    application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

                    // If found, set visibility to false if it's not already
                    if (application != null && application.Visible)
                    {
                        application.Visible = false;
                    }
                }
                catch (COMException)
                {
                    // If no instance is running, start a new one
                    application = (SolidEdgeFramework.Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                    application.Visible = false; // Make sure it's invisible
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error initializing Solid Edge: {ex.Message}");
            }

            return application;
        }
        public SolidEdgeFramework.SolidEdgeDocument TryOpenDocument(SolidEdgeFramework.Application application, string filePath, int retryCount = 5)
        {
            SolidEdgeFramework.SolidEdgeDocument document = null;
            int attempts = 0;

            while (attempts < retryCount)
            {
                try
                {
                    // Attempt to open the document normally
                    document = (SolidEdgeFramework.SolidEdgeDocument)application.Documents.Open(filePath);
                    return document; // Exit the loop if successful
                }
                catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x8001010A)
                {
                    // Handle the application busy error
                    LogMessage($"Solid Edge is busy. Retrying... Attempt {attempts + 1} of {retryCount}");
                    System.Threading.Thread.Sleep(1000); // Wait for 1 second before retrying
                    attempts++;
                }
                catch (Exception ex)
                {
                    LogMessage($"Error opening document: {ex.Message}");
                    break;
                }
            }

            return document; // Returns null if it failed to open after retries
        }
        Dictionary<string, double> ExtractAssemblyProperties(SolidEdgeFramework.SolidEdgeDocument document)
        {
            int retryCount = 0;
            const int maxRetries = 5;
            const int retryDelay = 1000; // 1 second delay between retries

            while (retryCount < maxRetries)
            {
                try
                {
                    
                    SolidEdgeAssembly.AssemblyDocument assemblyDoc = (SolidEdgeAssembly.AssemblyDocument)document; // Cast the document
                    
                    Dictionary<string, double> basePartCounts = new Dictionary<string, double>(); // Create a dictionary to store base part names (e.g., MN387306012_10.par) and their counts

                    // Iterate through the occurrences in the assembly
                    foreach (SolidEdgeAssembly.Occurrence occurrence in assemblyDoc.Occurrences)
                    {
                        
                        string originalName = occurrence.Name; // Extract occurrence name (with suffix)
                        string baseName = originalName.Split(':')[0]; // remove index by removing everything after the :
                        
                        if (basePartCounts.ContainsKey(baseName)) //if the base part name is already in the dictionary
                        {
                            basePartCounts[baseName]++; // If yes, increase the count and skip
                        }
                        else
                        {
                            basePartCounts[baseName] = 1; // If no, add the base name to the dictionary with a count of 1
                        }
                    }
                    return basePartCounts;
                }
                catch (System.Runtime.InteropServices.COMException ex) when ((uint)ex.ErrorCode == 0x8001010A)
                {
                    // Error: RPC_E_SERVERCALL_RETRYLATER - Application is busy, retry after a delay
                    LogMessage($"Solid Edge is busy, retrying... ({retryCount + 1}/{maxRetries})");

                    // Wait for a while before retrying
                    System.Threading.Thread.Sleep(retryDelay);
                    retryCount++;
                }
                catch (Exception ex)
                {
                    LogMessage($"Error extracting assembly properties: {ex.Message}");
                    return new Dictionary<string, double>();
                }
            }

            // Return an empty dictionary if maximum retries are reached
            LogMessage("Failed to extract properties after multiple retries.");
            return new Dictionary<string, double>();
        }
        Dictionary<string, double> ExtractDraftProperties(SolidEdgeFramework.SolidEdgeDocument document)
        {
            int retryCount = 0;
            const int maxRetries = 5;
            const int retryDelay = 1000; // 1 second delay between retries

            // Initialize the dictionary to an empty dictionary
            Dictionary<string, double> PartList_DFT = new Dictionary<string, double>();

            while (retryCount < maxRetries)
            {
                try
                {
                    // Cast the document to a DraftDocument (SolidEdgeDraft.DraftDocument)
                    SolidEdgeDraft.DraftDocument draftDoc = (SolidEdgeDraft.DraftDocument)document;

                    string DFT_NAME = document.Name;
                    // Dictionary to store unique model references
                    Dictionary<string, SolidEdgeDraft.DrawingView> uniqueViews = new Dictionary<string, SolidEdgeDraft.DrawingView>();

                    foreach (SolidEdgeDraft.Sheet sheet in draftDoc.Sheets) // Iterate through the sheets
                    {
                        foreach (SolidEdgeDraft.DrawingView drawingView in sheet.DrawingViews) // Iterate through the DrawingViews on each sheet (model views)
                        {
                            // Access the ModelLink to get the file name of the referenced model
                            SolidEdgeDraft.ModelLink modelLink = drawingView.ModelLink as SolidEdgeDraft.ModelLink;

                            if (modelLink != null)
                            {
                                string modelFileName = modelLink.FileName;

                                // Check if the model file name is .par or a text part. skip if it does
                                if (!string.IsNullOrEmpty(modelFileName)
                                        && !modelFileName.EndsWith(".par", StringComparison.OrdinalIgnoreCase)
                                        && !modelFileName.Contains(@"\Texte für DFT"))
                                {
                                    // Add unique view to the dictionary based on the model file name
                                    if (!uniqueViews.ContainsKey(modelFileName))
                                    {
                                        uniqueViews[modelFileName] = drawingView; // Add unique view to the dictionary
                                    }
                                }
                            }
                        }
                    }

                    // Now uniqueViews contains all unique drawing views based on their referenced model
                    List<SolidEdgeDraft.DrawingView> viewsToProcess = new List<SolidEdgeDraft.DrawingView>(uniqueViews.Values);

                    // Process each unique drawing view
                    foreach (var view in viewsToProcess)
                    {
                        SolidEdgeDraft.ModelLink modelLink = view.ModelLink as SolidEdgeDraft.ModelLink;

                        if (modelLink != null)
                        {
                            string modelFileName = modelLink.FileName;
                            LogMessage($"Processing referenced model file: {modelFileName}");
                            

                            // Open the referenced model document
                            SolidEdgeFramework.SolidEdgeDocument referencedDoc = TryOpenDocument(modelLink.Application, modelFileName);

                            // Check if it's a part or assembly and extract properties accordingly
                            if (referencedDoc.Type == SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument)
                            {
                                model_used_in_DFT = modelFileName;
                                PartList_DFT = ExtractAssemblyProperties(referencedDoc);
                            }
                        }
                    }

                    break; // Exit the loop if processing is successful
                }
                catch (System.Runtime.InteropServices.COMException ex) when ((uint)ex.ErrorCode == 0x8001010A)
                {
                    // Error: RPC_E_SERVERCALL_RETRYLATER - Application is busy, retry after a delay
                    LogMessage($"Solid Edge is busy, retrying... ({retryCount + 1}/{maxRetries})");

                    // Wait for a while before retrying
                    System.Threading.Thread.Sleep(retryDelay);
                    retryCount++;
                }
                catch (Exception ex)
                {
                    LogMessage($"Error extracting draft properties: {ex.Message}");
                    break; // Break out of the loop if there's a different error
                }
            }

            if (retryCount == maxRetries)
            {
                LogMessage("Failed to extract properties after multiple retries.");
            }

            // Ensure that we return the dictionary, even if it's empty
            return PartList_DFT;
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
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(asmName); //Remove the extension from the asmName

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
                                IsThereALinkedDFT = true;
                                Linked_DFT = $"{colM}-{colN}";
                                return $"Linked to {colM}-{colN} in ZeichnungVerknupfung"; // Return M-N if M is filled
                            }

                            // Step 7: If M is empty, check if column J is filled
                            string colJ = worksheet.Cells[row, 10].Text; // Column J (10th column)
                            string colK = worksheet.Cells[row, 11].Text; // Column K (11th column)

                            if (!string.IsNullOrEmpty(colJ))
                            {
                                IsThereALinkedDFT = true; 
                                Linked_DFT = $"{colJ}-{colK}";
                                return $"Linked to {colJ}-{colK} in ZeichnungVerknupfung"; // Return J-K if J is filled
                            }

                            // Step 8: If J is empty, return the content of D and E columns
                            string colD = worksheet.Cells[row, 4].Text; // Column D (4th column)
                            string colE = worksheet.Cells[row, 5].Text; // Column E (5th column)

                            if (!string.IsNullOrEmpty(colD) && !string.IsNullOrEmpty(colE))
                            {
                                IsThereALinkedDFT = true;
                                Linked_DFT_Path = $"Z:\\Zeichnungen\\DFT\\{colD}-{colE}.dft";
                                return $"No F-drawing linked just {colD}-{colE} in ZeichnungVerknupfung"; // Return D-E if D and E are filled
                            }

                            // Step 9: If J and M are empty, return this message
                            IsThereALinkedDFT = false;
                            return $"No DFT linked to {fileNameWithoutExtension} in ZeichnungVerknupfung\nUsing Parts list from 3D in the BOM";
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
            IsThereALinkedDFT = false;
            Linked_DFT_Path = string.Empty;

            InputAsmName = TxtAsmName.Text;  // Get the Assembly name
            string InputAsmName_ext = $"{InputAsmName}.asm";
            string AsmFilePath = GenerateFilePath(InputAsmName_ext); //Generate file path
            bool pathExists = CheckIfPathExists(AsmFilePath);
            if (pathExists == true)
            {
                LblPathStatus.Text = $"{AsmFilePath} ---> Exists :) ";
            }
            else if (pathExists == false)  
            {
                LblPathStatus.Text = $"{AsmFilePath} ---> Does not exist :(";
            }


            string ZeichnungVerknupfungResult = searchZeichnungVerknupfung(InputAsmName); // Call the Excel search function
            label4.Text = ZeichnungVerknupfungResult; // Update label4 with the result from Excel search

            // Initial path setup
            Linked_DFT_Path = $"Z:\\Zeichnungen\\DFT\\{Linked_DFT}.dft";

            if (System.IO.File.Exists(Linked_DFT_Path))
            {
                label4.Text += System.Environment.NewLine + "Draft found in DFT :)";
            }
            else
            {
                // Change the path to the alternate location
                Linked_DFT_Path = $"Z:\\Zeichnungen\\alte DFT\\{Linked_DFT}.dft";

                // Check for existence in the alternate location
                if (System.IO.File.Exists(Linked_DFT_Path))
                {
                    label4.Text += System.Environment.NewLine + "Draft found in alte DFT :)";
                }
                else
                {
                    label4.Text += System.Environment.NewLine + "Cannot find DFT file";
                }
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            Extracted_SAP_List = GetSAP_BOM();
            if (Extracted_From_SAP == true)
            {
                Labelsapstatus.Text = $"Succesfully fetched SAP BOM";
            }
            else
            {
                Labelsapstatus.Text = $"Error: There was an issue fetching SAP BOM";
            }
            Label3dstatus.Text = $"Opening Solidedge to get 3D model data. please wait..";
            SolidEdgeFramework.SolidEdgeDocument document = null;
            var application = InitializeSolidEdge();
            int retryCount = 5; // Define retryCount here
            
            try // Open the document with retries
            {
                if (IsThereALinkedDFT == true)
                {
                    string filePath = Linked_DFT_Path;
              
                    document = TryOpenDocument(application, filePath, retryCount);
                    Extracted_3D_List = ExtractDraftProperties(document);
                }
                else if (IsThereALinkedDFT == false) // Process inputed assembly document
                {
                    Extracted_3D_List = ExtractAssemblyProperties(document);

                    //            SingleBOM_documentlevel = ExtractAssemblyProperties(DFT_NAME, document);
                    //            LogMessage();
                }
            }
            catch (Exception ex)
            {
                Label3dstatus.Text = $"Error: There was an issue fetching 3D Partlist";
            }
            finally
            {
                // Properly close and release the document without quitting Solid Edge
                if (document != null)
                {
                    document.Close(false); // Close without saving changes
                }
            }
            Label3dstatus.Text = $"Succesfully fetched 3D Partlist";
            return;
            }
        Dictionary<string, double> GetSAP_BOM()
        {
            Dictionary<string, double> Dict_SAP_BOM = new Dictionary<string, double>(); // Dict of part name and quantity

            string SAP_BOM_path = "Z:\\Allgemein\\Temp\\MA\\COMPARISON\\BOM.txt"; // Path to the BOM file

            try
            {
                string[] lines = File.ReadAllLines(SAP_BOM_path); // Read the file line by line

                // Determine the selected quantity filter from radio buttons
                string selectedQuantity = string.Empty;
                if (radioButton1000.Checked)        {   selectedQuantity = "1000";  }
                else if (radioButton2000.Checked)   {   selectedQuantity = "2000";  }
                else if (radioButton8000.Checked)   {   selectedQuantity = "8000";  }


                foreach (string line in lines)
                {
                    // Only process lines that start with the assembly name
                    if (line.StartsWith(InputAsmName, StringComparison.OrdinalIgnoreCase) &&
                        (line.Length == InputAsmName.Length || line[InputAsmName.Length] == '\t' || line[InputAsmName.Length] == ' '))
                    {
                        // Use regex to split the line by whitespace and capture relevant groups
                        var matches = Regex.Matches(line, @"\S+");
                        var columns = matches.Cast<Match>().Select(m => m.Value.Trim()).ToArray();

                        // Check if the line has enough columns (we need at least 6)
                        if (columns.Length >= 6)
                        {
                            // Check if the third column contains 'T'
                            if (columns[2].Equals("T", StringComparison.OrdinalIgnoreCase))
                            {
                                // Skip the line if the third column contains 'T'
                                LogMessage($"Info: Ignoring line with 'T' in the third column: {line}");
                                continue;
                            }

                            // Extract part name (column 4) and quantity (column 5)
                            string partName = columns[3];   // Column 4 (part name)
                            string quantity = columns[4];    // Column 5 (quantity)

                            // Check if part name and quantity are valid
                            if (!string.IsNullOrEmpty(partName) && !string.IsNullOrEmpty(quantity))
                            {
                                // Remove any potential decimal commas and trim the quantity
                                quantity = quantity.Replace(',', '.').Trim();
                                double quantity_d = Convert.ToDouble(quantity, System.Globalization.CultureInfo.InvariantCulture);
                                if (line.EndsWith(selectedQuantity))
                                {
                                    Dict_SAP_BOM[partName] = quantity_d; // Add or update the dictionary
                                }
                            }
                            else
                            {
                                LogMessage($"Warning: Missing part name or quantity for line: {line}");
                            }
                        }
                        else
                        {
                            LogMessage($"Warning: Skipped line with insufficient columns: {line}");
                        }
                    }
                }

            }
            catch (FileNotFoundException ex)
            {

            }
            catch (IOException ex)
            {
             
            }
            Extracted_From_SAP = true;
            return Dict_SAP_BOM;
        }
        public void Translate3DList( Dictionary<string, double> Extracted_3D_List, out Dictionary<string, double> Translated_3D_List)
        {
            Translated_3D_List = new Dictionary<string, double>();
            string Traslation_excel_path = $"Z:\\Entwicklung\\Intern\\Archiv\\SE und SAP Artikelnummern.xlsx";

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; //license context for EPPlus

            FileInfo fileInfo = new FileInfo(Traslation_excel_path); // Load the Excel file

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the first worksheet

                // Get the total number of rows in the worksheet
                int rowCount = worksheet.Dimension.Rows;

                // Loop through each entry in Extracted_3D_List
                foreach (var entry in Extracted_3D_List)
                {
                    string partName = entry.Key; // The key from the extracted list
                    double quantity = entry.Value; // The associated quantity

                    // Initialize a variable to track if a match was found
                    bool foundMatch = false;

                    // Loop through the rows in column B of the worksheet to find a match
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        // Get the value in column B for the current row
                        string cellValue = worksheet.Cells[row, 2].Text; // Column B

                        // Check if it matches the part name
                        if (cellValue.Equals(partName, StringComparison.OrdinalIgnoreCase))
                        {
                            // If a match is found, get the value from column C
                            string newKey = worksheet.Cells[row, 3].Text; // Column C

                            // Add to Translated_3D_List with newKey and quantity
                            Translated_3D_List[newKey] = quantity; // Quantity remains the same

                            foundMatch = true; // Mark as found
                            break; // Exit the loop as we found a match
                        }
                    }

                    // If not found, add the original string and value to Translated_3D_List
                    if (!foundMatch)
                    {
                        Translated_3D_List[partName] = quantity; // Keep the original value
                    }
                }
            }
        }
        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Check if the current column is the fourth column (index 3)
            if (e.ColumnIndex == 3 && e.Value != null)
            {
                // Apply colors based on the cell's content
                string cellValue = e.Value.ToString();

                if (cellValue == "IO")
                {
                    e.CellStyle.BackColor = Color.Green;
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (cellValue == "Not in SAP" || cellValue == "Not in 3D")
                {
                    e.CellStyle.BackColor = Color.Red;
                    e.CellStyle.ForeColor = Color.White;
                }
                else if (cellValue == "Diff Qu")
                {
                    e.CellStyle.BackColor = Color.Yellow;
                    e.CellStyle.ForeColor = Color.Black;
                }
                else
                {
                    // Set default color for cells that don't match the above criteria
                    e.CellStyle.BackColor = Color.White;
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Translate3DList(Extracted_3D_List, out Translated_3D_List);
            var singleBOM = new Dictionary<string, (double, double, string)>();
            singleBOM = Create_Single_BOM(Translated_3D_List, Extracted_SAP_List);
            DisplaySingleBOM(singleBOM);
            dataGridView1.CellFormatting += DataGridView1_CellFormatting;
            if (document != null)
            {
                document.Close(false); // Close without saving changes
            }
        }
        private Dictionary<string, (double, double, string)> Create_Single_BOM(Dictionary<string, double> basePartCounts, Dictionary<string, double> DictSAPBOM)
        {
            // This will hold the final results
            var singleBOM = new Dictionary<string, (double base3DQuantity, double sapQuantity, string someString)>();

            // First loop: Iterate over DictSAPBOM
            foreach (var sapEntry in DictSAPBOM)
            {
                string sapPartName = sapEntry.Key;
                double sapQuantity = sapEntry.Value;

                // Exact match check
                if (basePartCounts.ContainsKey(sapPartName))
                {
                    double base3DQuantity = basePartCounts[sapPartName];
                    singleBOM[sapPartName] = (base3DQuantity, sapQuantity, ""); // Indicator will be added later
                }

                // If no partial match was found, put name and quantity from SAP and 0 for 3D quantity, with indicator "Not in 3D"
                else
                {
                    singleBOM[sapPartName] = (0, sapQuantity, "Not in 3D"); // 0 for 3D quantity, indicator "Not in 3D"
                }
            }
            

            // Second loop: Iterate over basePartCounts for unmatched parts in SingleBOM
            foreach (var baseEntry in basePartCounts)
            {
                string basePartName = baseEntry.Key;
                double base3DQuantity = baseEntry.Value;

                // Add to SingleBOM only if basePartName is not present in singleBOM (no exact or partial match)
                bool foundMatch = singleBOM.ContainsKey(basePartName);

                if (!foundMatch)
                {
                    singleBOM[basePartName] = (base3DQuantity, 0, "Not in SAP"); // 0 for SAP quantity, indicator "Not in SAP"
                }
            }

            // Third loop: Add indicators to the singleBOM if the indicator is empty
            foreach (var bomEntry in singleBOM.ToList())
            {
                if (disregardList.Contains(bomEntry.Key) || bomEntry.Key.StartsWith("ILLEK") || bomEntry.Key.StartsWith("VE"))
                {
                    singleBOM.Remove(bomEntry.Key); // Remove this entry from singleBOM
                    continue; // Move to the next iteration without further processing this entry
                }
                double base3DQuantity = bomEntry.Value.Item1;
                double sapQuantity = bomEntry.Value.Item2;
                string currentIndicator = bomEntry.Value.Item3;

                // Only update the indicator if it's empty
                if (string.IsNullOrEmpty(currentIndicator))
                {
                    string indicator = (base3DQuantity == sapQuantity) ? "IO" : "Diff Qu";
                    singleBOM[bomEntry.Key] = (base3DQuantity, sapQuantity, indicator);
                }
            }

            // Sort the Single BOM by SAP part names, excluding the header
            var sortedSingleBOM = singleBOM
                .OrderBy(kvp => kvp.Key == "Header" ? "0" : kvp.Key) // Keep header first
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            // Add headers at the beginning of the singleBOM
            return sortedSingleBOM;
        }
        private void DisplaySingleBOM(Dictionary<string, (double base3DQuantity, double sapQuantity, string someString)> singleBOM)
        {
            // Clear any existing rows
            dataGridView1.Rows.Clear();

            // Add columns if not already defined in the designer
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("PartName", "Part Name");
            dataGridView1.Columns.Add("Base3DQuantity", "Base 3D Quantity");
            dataGridView1.Columns.Add("SAPQuantity", "SAP Quantity");
            dataGridView1.Columns.Add("Indicator", "Indicator");

            // Populate the DataGridView
            foreach (var entry in singleBOM)
            {
                dataGridView1.Rows.Add(entry.Key, entry.Value.base3DQuantity, entry.Value.sapQuantity, entry.Value.someString);
            }
        }
    }
}