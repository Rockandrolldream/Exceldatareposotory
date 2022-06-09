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
                            Validatedata(outputexcel);
                        }
                    }
                }      
            }
            return exceldata;
        }

        public async void Validatedata( string outputexcel)
        {
            switch (outputexcel)
            {
                case null:
                  //  Console.WriteLine("string may not be null");
                    break;
                case "":
                 //   Console.WriteLine("string is empthy");
                    break;
                default: 

                    if (outputexcel.Contains("http"))
                    {
                        var sw = new Stopwatch();
                        sw.Start();
                        Task delay = Task.Delay(30000);
                        SendRequestvalidate(outputexcel);
                        await delay;
                    }
                    break;
            }
        }

        public async Task SendRequestvalidate(string outputexcel)
        {
                var url = outputexcel;
                RestClient client = new RestClient(url);
                var request = new RestRequest(url, Method.Get);
                RestResponse response = await client.ExecuteAsync(request);
                var Output = response.StatusCode.ToString();
                Console.WriteLine(Output);
            if (Output == "OK")
            {
                exceldata.Add(outputexcel);
            }
        }
    }
}
