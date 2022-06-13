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
using System.Collections.Concurrent;

namespace Exceldatascript
{
    public class ExcelDataScriptExecute
    {
        List<String> exceldata = new List<String>();
        List<int> notdownloaded = new List<int>();
        Dictionary<string, string> goingwell = new Dictionary<string, string>();
        List<ExcelObject> PDFdownloadMetadata2006_2016 = new List<ExcelObject>();
        ConcurrentBag<ExcelObject> PDFdownloadGRI_2017_2020 = new ConcurrentBag<ExcelObject>();
        public ConcurrentBag<ExcelObject> GetDataTableFromExcel(int coloumnumber)
        {
            Console.WriteLine("Begin to read PDF links");
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
                            var Brnum = worksheet.Cells[i, 1].Value.ToString();
                            Validatedata(outputexcel , i, Brnum);
                        }
                    }
                }      
            }
            Console.WriteLine("You are done");
            return PDFdownloadGRI_2017_2020;
        }

        public  async void Validatedata(string outputexcel ,int row, string Brnum)
        {
            if (outputexcel.Contains("http") && outputexcel.Contains("pdf") && outputexcel.Any(char.IsWhiteSpace) == false)
            {
              await SendRequestvalidate(outputexcel, row, Brnum);
            }
            
        }

        public async Task SendRequestvalidate(string outputexcel, int row, string Brnum)
        {
            var url = outputexcel;
                RestClient client = new RestClient(url);
                var request = new RestRequest(url, Method.Get);
                request.Timeout = 2000;
                RestResponse response = await client.ExecuteAsync(request);
                var Output = response.StatusCode.ToString();
          //  Console.WriteLine(Output + "  "  + outputexcel);
            if (Output == "OK")
            {
                PDFdownloadGRI_2017_2020.Add(new ExcelObject(outputexcel, true, row, Brnum));
            }
            else
            {
                PDFdownloadGRI_2017_2020.Add(new ExcelObject(outputexcel, false, row, Brnum));
            }
        }

        public async Task Downloadfiles(ConcurrentBag<ExcelObject> downloadinput)
        {
            foreach (var item in downloadinput)
            {
                for (int i = 0; i < downloadinput.Count; i++)
                {
                    if (item.Isdownloaded == true)
                    {
                        var url = item.Pdf_URL;
                        RestClient client = new RestClient(url);
                        var request = new RestRequest(url, Method.Get);
                        request.Timeout = 2000;
                        var response = client.DownloadData(new RestRequest(request.ToString()));
                        Console.WriteLine(response);
                        await File.WriteAllBytesAsync("C://Downloadedpdfs", response);
                    }
                }
               
            }
        }

        public ConcurrentBag<ExcelObject> LookIntoAnotherColoumn(ConcurrentBag<ExcelObject> exceldata)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            byte[] bin = File.ReadAllBytes("C:\\Users\\KOM\\Desktop\\Exceldatascriptopgave\\GRI_2017_2020.xlsx");

            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        foreach (var item in exceldata)
                        {
                            if (item.Isdownloaded == false)
                            {
                                var outputsecoundcoloumn = worksheet.Cells[item.Rownumber, 39].Value.ToString();

                                Validatedata(outputsecoundcoloumn, item.Rownumber, item.BRnum);
                            }
                            
                        }

                    }
                }
            }
            return PDFdownloadGRI_2017_2020;
        }

        public List<ExcelObject> Metadata2006_2016()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string path = "C:/Users/KOM/Desktop/Exceldatascriptopgave/Metadata2006_2016.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
            ExcelWorksheet worksheetdownlaoded = package.Workbook.Worksheets.FirstOrDefault();

            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                string pdflink = worksheet.Cells[i, 34].Value.ToString();
                string iddownloaded = worksheetdownlaoded.Cells[i, 46]?.Value?.ToString();

                if (iddownloaded == "YES")
                {
                    PDFdownloadMetadata2006_2016.Add(new ExcelObject(pdflink, true));
                } 
                else
                {
                    PDFdownloadMetadata2006_2016.Add(new ExcelObject(pdflink, false));
                }

            }
            return PDFdownloadMetadata2006_2016;
        }

        public void Helperdownlaod(ConcurrentBag<ExcelObject> downloadinput)
        {
           Downloadfiles(downloadinput);
        }

    }
}
