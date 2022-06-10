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
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace Exceldatascript
{
    public class ExcelDataScriptExecute
    {
        List<String> exceldata = new List<String>();
        List<String> notdownloaded = new List<String>();
        Dictionary<string, string> goingwell = new Dictionary<string, string>();
        public async Task <Dictionary<string, string>> GetDataTableFromExcel(int coloumnumber)
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
                            Validatedata(outputexcel);
                        }
                    }
                }      
            }
            Console.WriteLine("You are done");
            return goingwell;
        }

        public void Validatedata(string outputexcel)
        {
            if (outputexcel.Contains("http"))
            {
                SendRequestvalidate(outputexcel);
            }
            else
            {
                LookIntoAnotherFile();
            }
            
        }

        public async Task SendRequestvalidate(string outputexcel)
        {
                var url = outputexcel;
                RestClient client = new RestClient(url);
                var request = new RestRequest(url, Method.Get);
                request.Timeout = 2000;
                RestResponse response = await client.ExecuteAsync(request);
                var Output = response.StatusCode.ToString();
            Console.WriteLine(Output + "  "  + outputexcel);
            if (Output == "OK")
            {
                goingwell.Add(outputexcel, Output);
            }
            else
            {  
              notdownloaded.Add(outputexcel);
            }
        }

        public async Task LookIntoAnotherFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                        {
                            var outputexcel3 = worksheet.Cells[i, 39].Value.ToString();
                            if (outputexcel3.Contains("http"))
                            {
                                SendRequestvalidate(outputexcel3);
                            }
                        }
                    }
                }
            }

        }

        public List<String> Displaylist()
        {
            Console.WriteLine("This is a list where there has been issusse with downloading");
            return notdownloaded;
        }
    }
}
