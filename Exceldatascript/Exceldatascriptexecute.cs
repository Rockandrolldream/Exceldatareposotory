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
using CsvHelper;
using System.Globalization;

namespace Exceldatascript
{
    public class ExcelDataScriptExecute
    {
        ConcurrentBag<ExcelObject> PDFLinksForcustomer = new ConcurrentBag<ExcelObject>(); 
        // GetDataFromExcelsheet
        public ConcurrentBag<ExcelObject> GetDataFromExcel()
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
                            var pdflinkname = worksheet.Cells[i, 38].Value.ToString();
                            var Brnum = worksheet.Cells[i, 1].Value.ToString();
                            ValidateDataAndSendsRequest(pdflinkname, i, Brnum);
                        }

                        foreach (var item in PDFLinksForcustomer)
                        {
                            if (item.Isdownloaded == "NotDownloaded")
                            {
                                var Pdflinksecoundcoloumn = worksheet.Cells[item.Rownumber, 39].Value.ToString();

                                ValidateDataAndSendsRequest(Pdflinksecoundcoloumn, item.Rownumber, item.BRnum);
                            }
                        }
                    }
                }
            }
            using (var writer = new StreamWriter("C:/CSVfiles/FileDatatoTheCustomer.csv"))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(PDFLinksForcustomer);
            }
            Console.WriteLine("You are done");
            return PDFLinksForcustomer;
        }
        // ValidateandSendRequest
        public async void ValidateDataAndSendsRequest(string pdflinkname, int row, string Brnum)
        {
            if (pdflinkname.Contains("http") && pdflinkname.Any(char.IsWhiteSpace) == false && pdflinkname.Length > 15)
            {
                await SendRequestvalidatePdflinks(pdflinkname, row, Brnum);
            }
            else
            {
                PDFLinksForcustomer.Add(new ExcelObject(pdflinkname, "NotDownloaded", row, Brnum));
            }
        }
        // SendsRequest and validates errors
        public async Task SendRequestvalidatePdflinks(string pdflinkname, int row, string Brnum)
        {
            var url = pdflinkname;
            RestClient client = new RestClient(url);
            var request = new RestRequest(url, Method.Get);
            request.Timeout = 10000;
            RestResponse response = await client.ExecuteGetAsync(request);
            var Output = response.StatusCode.ToString();
            Console.WriteLine(Output + "  " + pdflinkname);
            if (response.ContentType != null && Output == "OK" && response.ContentType.Contains("pdf"))
            {
                PDFLinksForcustomer.Add(new ExcelObject(pdflinkname, "IsDownloaded", row, Brnum));
                Downloadfiles(pdflinkname, row, Brnum);
            }
            else
            {
                PDFLinksForcustomer.Add(new ExcelObject(pdflinkname, "NotDownloaded", row, Brnum));
            }
        }
        // Downloadfiles 
        public async Task Downloadfiles(string pdflinkname, int row, string Brnum)
        {
            var url = pdflinkname;
            RestClient client = new RestClient(url);
            var request = new RestRequest(url, Method.Get);
            request.Timeout = 15000;
            var response = client.DownloadDataAsync(request);
            if (response != null)
            {
                String path = @"C:\Downloadedpdfs\";
                string combinepath = path + Brnum + ".pdf";
                Console.WriteLine(combinepath);
                File.WriteAllBytesAsync(combinepath, response.Result);
            }
        }
        // Class update metaDatafile  
        public List<ExcelObject> UpdateMetaData()
        {
            List<ExcelObject> UpdateMetaData = new List<ExcelObject>();
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
                if (isdownloaded == "yes")
                {
                    UpdateMetaData.Add(new ExcelObject(pdflink, "IsDownloaded"));
                }
                else
                {
                    UpdateMetaData.Add(new ExcelObject(pdflink, "NotDownloaded"));
                }

            }

            using (var writer = new StreamWriter("C:/CSVfiles/OldData.csv"))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(UpdateMetaData);
            }

            return UpdateMetaData;
        }
    }
}
