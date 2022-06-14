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
        List<ExcelObject> PDFdownloadMetadata2006_2016 = new List<ExcelObject>();
        ConcurrentBag<ExcelObject> PDFdownloadGRI_2017_2020 = new ConcurrentBag<ExcelObject>();
        public ConcurrentBag<ExcelObject> GetDataFromExcel(int coloumnumber)
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
                            var pdflinkname = worksheet.Cells[i, coloumnumber].Value.ToString(); 
                            var Brnum = worksheet.Cells[i, 1].Value.ToString();
                            Validatedata(pdflinkname, i, Brnum);
                        }
                    }
                }      
            }
            Console.WriteLine("You are done");
            return PDFdownloadGRI_2017_2020;
        }

        public  async void Validatedata(string pdflinkname, int row, string Brnum)
        {
            if (pdflinkname.Contains("http") && pdflinkname.Contains("pdf") && pdflinkname.Any(char.IsWhiteSpace) == false)
            {
              await SendRequestvalidate(pdflinkname, row, Brnum);
            }
            
        }

        public async Task SendRequestvalidate(string pdflinkname, int row, string Brnum)
        {
            var url = pdflinkname;
                RestClient client = new RestClient(url);
                var request = new RestRequest(url, Method.Get);
                request.Timeout = 2000;
                RestResponse response = await client.ExecuteAsync(request);
                var Output = response.StatusCode.ToString();
              Console.WriteLine(Output + "  "  + pdflinkname);
            if (Output == "OK")
            {
                PDFdownloadGRI_2017_2020.Add(new ExcelObject(pdflinkname, true, row, Brnum));
            }
            else
            {
                PDFdownloadGRI_2017_2020.Add(new ExcelObject(pdflinkname, false, row, Brnum));
            }
        }

        public async Task Downloadfiles(ConcurrentBag<ExcelObject> downloadinput)
        {
            foreach (var item in downloadinput)
            {
                    if (item.Isdownloaded == true)
                    {
                        var url = item.Pdf_URL;
                        RestClient client = new RestClient(url);
                        var request = new RestRequest(url, Method.Get);
                        request.Timeout = 2000;
                        var response = client.DownloadDataAsync(request);
                        Console.WriteLine(response);
                    if (response != null)
                    {
                       var pdfname =response.ToString().Split('/').Last(); 
                        String path = @"C:\Downloadedpdfs\";
                        string combinepath = path + pdfname;
                        Console.WriteLine("test");
                        File.WriteAllBytesAsync(combinepath, response.Result);
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
                                var Pdflinksecoundcoloumn = worksheet.Cells[item.Rownumber, 39].Value.ToString();

                                Validatedata(Pdflinksecoundcoloumn, item.Rownumber, item.BRnum);
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
                string isdownloaded = worksheetdownlaoded.Cells[i, 46]?.Value?.ToString();

                if (isdownloaded == "YES")
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

    }
}
