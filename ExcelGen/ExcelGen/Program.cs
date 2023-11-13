using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelGen
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo outputDir = new DirectoryInfo(@"C:\Dev\POC Excel Writer\ExcelGen\Sample");
            if (!outputDir.Exists) throw new Exception("outputDir does not exist!");

            Console.WriteLine("Running sample");
            string output =  CreateExcelFile(outputDir);

            Console.WriteLine("Sample 1 created: {0}", output);
            Console.WriteLine();

            Console.ReadLine();
        }

        private static string CreateExcelFile(DirectoryInfo outputDir)
        {
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
            }
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");

                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Product";
                worksheet.Cells[1, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Price";
                worksheet.Cells[1, 5].Value = "Value";

                worksheet.Cells["A2"].Value = 12001;
                worksheet.Cells["B2"].Value = "PS4";
                worksheet.Cells["C2"].Value = 30;
                worksheet.Cells["D2"].Value = 5000.99;

                worksheet.Cells["A3"].Value = 12002;
                worksheet.Cells["B3"].Value = "Call of Duty";
                worksheet.Cells["C3"].Value = 5;
                worksheet.Cells["D3"].Value = 1200.10;

                //Add a formula for the value-column
                worksheet.Cells["E2:E4"].Formula = "C2*D2";


                // set some document properties
                package.Workbook.Properties.Title = "Invertory";
                package.Workbook.Properties.Author = "Andre Barnard";
                package.Workbook.Properties.Comments = "POC for saving excel file";

                // set some extended property values
                package.Workbook.Properties.Company = "Singular Systems";

                // save our new workbook and we are done!
                package.Save();

            }

            return newFile.FullName;
        }

    }
}
