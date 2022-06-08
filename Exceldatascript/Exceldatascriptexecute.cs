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

namespace Exceldatascript
{
    public class ExcelDataScriptExecute
    {

        public List<String> GetDataTableFromExcel(int coloumnumber)
        {
            string content = "";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
             

            List<String> exceldata = new List<String>();

            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin)) {
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                        {
                                if (worksheet.Cells[i, coloumnumber].Value != null)
                                {
                                    Console.WriteLine(worksheet.Cells[i, coloumnumber].Value.ToString());
                                    
                                    exceldata.Add(worksheet.Cells[i, coloumnumber].Value.ToString());
                                }
                        }
                    }
                }
            
            }
            Console.WriteLine(exceldata.Count);
            return exceldata;
        }
    }
}
