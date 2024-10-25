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

namespace Quick_BOM_Compare
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        bool IsThereALinkedDFT = false;
        string Linked_DFT_Path = string.Empty;

        public static SolidEdgeFramework.Application InitializeSolidEdge()
        {
            SolidEdgeFramework.Application application = null;

            try
            {
                // Attempt to start Solid Edge application (Solid Edge must be installed)
                application = (SolidEdgeFramework.Application)Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));

                // Set to false to hide the Solid Edge window
                application.Visible = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing Solid Edge: {ex.Message}");
            }

            return application;
        }
        static SolidEdgeFramework.SolidEdgeDocument TryOpenDocument(SolidEdgeFramework.Application application, string filePath, int retryCount = 5)
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
                    Console.WriteLine($"Solid Edge is busy. Retrying... Attempt {attempts + 1} of {retryCount}");
                    System.Threading.Thread.Sleep(1000); // Wait for 1 second before retrying
                    attempts++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error opening document: {ex.Message}");
                    break;
                }
            }

            return document; // Returns null if it failed to open after retries
        }

        static Dictionary<string, double> ExtractAssemblyProperties(SolidEdgeFramework.SolidEdgeDocument document)
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
                        string baseName = originalName.Split('.')[0]; // remove extnsion by removing everything after the dot
                        
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
                    Console.WriteLine($"Solid Edge is busy, retrying... ({retryCount + 1}/{maxRetries})");

                    // Wait for a while before retrying
                    System.Threading.Thread.Sleep(retryDelay);
                    retryCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error extracting assembly properties: {ex.Message}");
                    return new Dictionary<string, double>();
                }
            }

            // Return an empty dictionary if maximum retries are reached
            Console.WriteLine("Failed to extract properties after multiple retries.");
            return new Dictionary<string, double>();
        }

        static Dictionary<string, double> ExtractDraftProperties(SolidEdgeFramework.SolidEdgeDocument document)
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
                            Console.WriteLine($"Processing referenced model file: {modelFileName}");
                            Console.WriteLine();

                            // Open the referenced model document
                            SolidEdgeFramework.SolidEdgeDocument referencedDoc = TryOpenDocument(modelLink.Application, modelFileName);

                            // Check if it's a part or assembly and extract properties accordingly
                            if (referencedDoc.Type == SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument)
                            {
                                PartList_DFT = ExtractAssemblyProperties(referencedDoc);
                            }
                        }
                    }

                    break; // Exit the loop if processing is successful
                }
                catch (System.Runtime.InteropServices.COMException ex) when ((uint)ex.ErrorCode == 0x8001010A)
                {
                    // Error: RPC_E_SERVERCALL_RETRYLATER - Application is busy, retry after a delay
                    Console.WriteLine($"Solid Edge is busy, retrying... ({retryCount + 1}/{maxRetries})");

                    // Wait for a while before retrying
                    System.Threading.Thread.Sleep(retryDelay);
                    retryCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error extracting draft properties: {ex.Message}");
                    break; // Break out of the loop if there's a different error
                }
            }

            if (retryCount == maxRetries)
            {
                Console.WriteLine("Failed to extract properties after multiple retries.");
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
                                IsThereALinkedDFT = true;
                                Linked_DFT_Path = $"Z:\\Zeichnungen\\DFT\\{colM}-{colN}.dft";
                                return $"Linked to {colM}-{colN} in ZeichnungVerknupfung"; // Return M-N if M is filled
                            }

                            // Step 7: If M is empty, check if column J is filled
                            string colJ = worksheet.Cells[row, 10].Text; // Column J (10th column)
                            string colK = worksheet.Cells[row, 11].Text; // Column K (11th column)

                            if (!string.IsNullOrEmpty(colJ))
                            {
                                IsThereALinkedDFT = true; 
                                Linked_DFT_Path = $"Z:\\Zeichnungen\\DFT\\{colJ}-{colK}.dft";
                                return $"Linked to {colJ}-{colK} in ZeichnungVerknupfung"; // Return J-K if J is filled
                            }

                            // Step 8: If J is empty, return the content of D and E columns
                            string colD = worksheet.Cells[row, 4].Text; // Column D (4th column)
                            string colE = worksheet.Cells[row, 5].Text; // Column E (5th column)

                            if (!string.IsNullOrEmpty(colD) && !string.IsNullOrEmpty(colE))
                            {
                                IsThereALinkedDFT = true;
                                Linked_DFT_Path = $"Z:\\Zeichnungen\\DFT\\{colD}-{colE}.dft";
                                return $"Linked to {colD}-{colE} in ZeichnungVerknupfung"; // Return D-E if D and E are filled
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

        private void button2_Click(object sender, EventArgs e)
        {
            SolidEdgeFramework.SolidEdgeDocument document = null;
            var application = InitializeSolidEdge();
            document = TryOpenDocument(application, filePath, retryCount);
            Dictionary<string, double> Extracted_3D_List = new Dictionary<string, double>();
        try // Open the document with retries
        {
            int retryCount = 5; // Define retryCount here
            if (IsThereALinkedDFT == true)
            {
                string filePath = Linked_DFT_Path;
                document = TryOpenDocument(application, filePath, retryCount);
                Extracted_3D_List = ExtractDraftProperties(document);
            }
            else if (IsThereALinkedDFT ==false) // Process inputed assembly document
            {
                Extracted_3D_List = ExtractAssemblyProperties(document);

                    //            SingleBOM_documentlevel = ExtractAssemblyProperties(DFT_NAME, document);
                    //            Console.WriteLine();
             return;
            }
        catch (Exception ex)
        {
             Console.WriteLine($"Error processing document: {ex.Message}");
        }
        finally
        {
        // Properly close and release the document without quitting Solid Edge
        if (document != null)
        {
         document.Close(false); // Close without saving changes
        }
                }
                return SingleBOM_documentlevel;
            }
            
        }
    }
}
