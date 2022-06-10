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
        List<int> notdownloaded = new List<int>();
        Dictionary<string, string> goingwell = new Dictionary<string, string>();
        public List<int> GetDataTableFromExcel(int coloumnumber)
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
                            Validatedata(outputexcel , i);
                        }
                    }
                }      
            }
            Console.WriteLine("You are done");
            return notdownloaded;
        }

        public void Validatedata(string outputexcel ,int row)
        {
            if (outputexcel.Contains("http") && outputexcel.Contains("pdf"))
            {
                SendRequestvalidate(outputexcel, row);
            }
            
        }

        public async Task SendRequestvalidate(string outputexcel, int row)
        {
                var url = outputexcel;
                RestClient client = new RestClient(url);
                var request = new RestRequest(url, Method.Get);
                request.Timeout = 5000;
                RestResponse response = await client.ExecuteAsync(request);
                var Output = response.StatusCode.ToString();
            Console.WriteLine(Output + "  "  + outputexcel);
            if (Output == "OK")
            {
                goingwell.Add(outputexcel, Output);
            }
            else
            {  
              notdownloaded.Add(row);
            }
        }

        public void LookIntoAnotherFile(List<int> notgoingwell)
        {
           ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        foreach (var item in notgoingwell)
                        {
                            var outputexcel3 = worksheet.Cells[item, 39].Value.ToString();
                            if (outputexcel3.Contains("http"))
                            {
                                SendRequestvalidate(outputexcel3, item);
                            }
                        }

                    }
                }
            }

        }

        public List<int> Displaylist()
        {
            Console.WriteLine("This is a list where there has been issusse with downloading");
            return notdownloaded;
        }
    }
}
