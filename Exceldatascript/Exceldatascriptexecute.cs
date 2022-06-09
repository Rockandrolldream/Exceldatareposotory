using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Data;
using Newtonsoft.Json;
using RestSharp;

namespace Exceldatascript
{
    public class ExcelDataScriptExecute
    {
        List<String> exceldata = new List<String>();
        public List<String> GetDataTableFromExcel(int coloumnumber)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin)) {
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {  
                        for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                        {
                                var outputexcel = worksheet.Cells[i, coloumnumber].Value.ToString();
                                ValidateThatLinkWorks(outputexcel , i);
                        }
                    }
                }      
            }
            return exceldata;
        }

        public void ValidateThatLinkWorks(string outputexcel, int row)
        {

                switch (outputexcel)
                {
                    case null:
                        Console.WriteLine("string may not be null");
                        HelpereMethode(row);                       
                        break;
                    case "":
                        Console.WriteLine("string is empthy");
                        HelpereMethode(row);
                    break;
                    default:
                        exceldata.Add(outputexcel);
                        Console.WriteLine(outputexcel);
                        break;
                }    
        }

        public void HelpereMethode(int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                            var outputexcel = worksheet.Cells[row, 39].Value.ToString();
                            switch (outputexcel)
                            {
                                case null:
                                    Console.WriteLine("string may not be null");

                                    break;
                                case "":
                                    Console.WriteLine("string is empthy");
                                    break;
                                default:
                                    exceldata.Add(outputexcel);
                                    Console.WriteLine(outputexcel);
                                    break;
                            }

                    }
                }
            }
        }
    }
}
