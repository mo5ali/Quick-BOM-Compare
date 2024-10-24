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
        }

    
    }
}
