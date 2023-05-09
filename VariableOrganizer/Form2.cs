using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;

namespace VariableOrganizer
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // create the table

            string name;
            string type;
            string startingValue;
            string use;
            string links;
            string comments;
            string projectName;
            string username = Environment.UserName;
            string variableCount = "1";
            int variableCountI = 1;
            projectName = textBox6.Text;

            System.IO.Directory.CreateDirectory("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Output " + projectName);

           

            string spreadSheetPath = "C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Output " + projectName + "\\" + "varOrganizer " + projectName + ".xlsx";
            File.Delete(spreadSheetPath);
            FileInfo spreadSheetInfo = new FileInfo(spreadSheetPath);
            ExcelPackage pck = new ExcelPackage(spreadSheetInfo);
            var activitiesWorksheet = pck.Workbook.Worksheets.Add("varOrganizer " + projectName);

            activitiesWorksheet.Cells["A1"].Value = "Alias";
            activitiesWorksheet.Cells["A1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(21, 105, 199));


            activitiesWorksheet.Cells["B1"].Value = "Type";
            activitiesWorksheet.Cells["B1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(199, 116, 22));
           


            activitiesWorksheet.Cells["C1"].Value = "Size (bytes)";
            activitiesWorksheet.Cells["C1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 232, 23));
         


            activitiesWorksheet.Cells["D1"].Value = "Range";
            activitiesWorksheet.Cells["D1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["D1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(232, 100, 199));
            


            activitiesWorksheet.Cells["E1"].Value = "Name";
            activitiesWorksheet.Cells["E1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 243, 128));
          


            activitiesWorksheet.Cells["F1"].Value = "Starting value";
            activitiesWorksheet.Cells["F1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["F1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 140, 255));
         


            activitiesWorksheet.Cells["G1"].Value = "Use";
            activitiesWorksheet.Cells["G1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["G1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 154, 77));
            


            activitiesWorksheet.Cells["H1"].Value = "Links";
            activitiesWorksheet.Cells["H1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["H1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(76, 159, 237));
           


            activitiesWorksheet.Cells["I1"].Value = "Comments";
            activitiesWorksheet.Cells["I1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            activitiesWorksheet.Cells["I1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 102, 0));
         




            int count = 1;
            int copy = variableCountI;

            int countC = 2;
            while (count < variableCountI)
            {

                string filename = projectName + count + ".txt"; ;

                StreamReader varN = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "N" + filename);
                StreamReader varT = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "T" + filename);
                StreamReader varS = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "S" + filename);
                StreamReader varU = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "U" + filename);
                StreamReader varL = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "L" + filename);
                StreamReader varC = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "C" + filename);

                string varNI = varN.ReadLine();
                string varTI = varT.ReadLine();
                string varSI = varS.ReadLine();
                string varUI = varU.ReadLine();
                string varLI = varL.ReadLine();
                string varCI = varC.ReadLine();

                string info1 = "";
                string info2 = "";
                string info3 = "";

                if (varTI == "byte")
                {
                    info1 = "0 to 255 ";
                    info2 = "Byte";
                    info3 = "8";
                }
                else if (varTI == "sbyte")
                {
                    info1 = "-128 to 127  ";
                    info2 = "SByte";
                    info3 = "8";
                }
                else if (varTI == "int  (Int32)")
                {
                    info1 = "-2,147,483,648 to 2,147,483,647 ";
                    info2 = "Int32";
                    info3 = "32";
                }
                else if (varTI == "uint  (UInt32)")
                {
                    info1 = "0 to 4294967295";
                    info2 = "UInt32";
                    info3 = "32";
                }
                else if (varTI == "short  (Int16)")
                {
                    info1 = "-32,768 to 32,767 ";
                    info2 = "Int16";
                    info3 = "16";
                }
                else if (varTI == "ushort  (UInt16)")
                {
                    info1 = "0 to 65,535 ";
                    info2 = "UInt16";
                    info3 = "16";
                }
                else if (varTI == "long  (Int64)")
                {
                    info1 = "-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 ";
                    info2 = "Int64";
                    info3 = "64";
                }
                else if (varTI == "ulong  (UInt64)")
                {
                    info1 = "0 to 18,446,744,073,709,551,615";
                    info2 = "UInt64";
                    info3 = "64";
                }
                else if (varTI == "float")
                {
                    info1 = "-3.402823e38 to 3.402823e38 ";
                    info2 = "Single";
                    info3 = "32";
                }
                else if (varTI == "double")
                {
                    info1 = "-1.79769313486232e308 to 1.79769313486232e308 ";
                    info2 = "Double";
                    info3 = "64";
                }
                else if (varTI == "char")
                {
                    info1 = "Unicode symbols used in text ";
                    info2 = "Char";
                    info3 = "16";
                }
                else if (varTI == "bool")
                {
                    info1 = "True or False ";
                    info2 = "Boolean";
                    info3 = "8";
                }
                else if (varTI == "object")
                {
                    info1 = "-";
                    info2 = "Object";
                    info3 = "-";
                }
                else if (varTI == "string")
                {
                    info1 = "-";
                    info2 = "String";
                    info3 = "-";
                }
                else if (varTI == "decimal")
                {
                    info1 = "(+ or -)1.0 x 10e-28 to 7.9 x 10e28 ";
                    info2 = "Decimal";
                    info3 = "128";
                }
                else if (varTI == "Date Time")
                {
                    info1 = "0:00:00am 1/1/01 to 11:59:59pm 12/31/9999 ";
                    info2 = "DateTime";
                    info3 = "-";
                }

                

                activitiesWorksheet.Cells["A" + countC].Value = varTI;   //alias
                activitiesWorksheet.Cells["A" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["A" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(21, 105, 199));

                activitiesWorksheet.Cells["B" + countC].Value = info2;   //type
                activitiesWorksheet.Cells["B" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["B" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(199, 116, 22));
               

                activitiesWorksheet.Cells["C" + countC].Value = info3;   // size 
                activitiesWorksheet.Cells["C" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["C" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(89, 232, 23));
              

                activitiesWorksheet.Cells["D" + countC].Value = info1;   // range   
                activitiesWorksheet.Cells["D" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["D" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(232, 100, 199));
                

                activitiesWorksheet.Cells["E" + countC].Value = varNI;   // name
                activitiesWorksheet.Cells["E" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["E" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 243, 128));
               

                activitiesWorksheet.Cells["F" + countC].Value = varSI;   // start val
                activitiesWorksheet.Cells["F" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["F" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 140, 255));
               

                activitiesWorksheet.Cells["G" + countC].Value = varUI;   // use
                activitiesWorksheet.Cells["G" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["G" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 154, 77));
               

                activitiesWorksheet.Cells["H" + countC].Value = varLI;   //links
                activitiesWorksheet.Cells["H" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["H" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(76, 159, 237));
                

                activitiesWorksheet.Cells["I" + countC].Value = varCI;    // comments
                activitiesWorksheet.Cells["I" + countC].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                activitiesWorksheet.Cells["I" + countC].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 102, 0));
               

                varN.Close();
                varT.Close();
                varS.Close();
                varU.Close();
                varL.Close();
                varC.Close();
                count = count + 1;
                countC = countC + 1;
            }
            pck.Save();
            System.Diagnostics.Process.Start("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Output " + projectName); // <- add the path
            System.Diagnostics.Process.Start("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Output " + projectName + "\\" + "varOrganizer " + projectName + ".xlsx"); // <- add the path
        }
    }
}
