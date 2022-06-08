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
    public class Exceldatascriptexecute
    {

        public  void GetDataTableFromExcel()
        {
            string content = "";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //FileInfo fileInfo = new FileInfo(path);
            //ExcelPackage package = new ExcelPackage(fileInfo);
            // ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(); 

            List<String> exceldata = new List<String>();

            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin)) {
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        for (int i = worksheet.Dimension.Start.Row; i < worksheet.Dimension.End.Row; i++)
                        {
                            for (int j = worksheet.Dimension.Start.Column; j < worksheet.Dimension.End.Column; j++)
                            {
                                if (worksheet.Cells[i, j].Value != null)
                                {
                                    exceldata.Add(worksheet.Cells[i,j].Value.ToString());
                                }
                            }
                        }
                    }
                }
            
            }

            foreach (var item in exceldata)
            {
                Console.WriteLine(item);
            }

        }
    }
}
