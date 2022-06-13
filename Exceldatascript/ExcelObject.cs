﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceldatascript
{
    public class ExcelObject
    {
        public string Pdf_URL { get; set; } 

        public bool Isdownloaded { get; set; } 

        public int Rownumber { get; set; } 

        public string BRnum { get; set; }

        public ExcelObject(string pdf_URL, bool isdownloaded)
        {
            Pdf_URL = pdf_URL;
            Isdownloaded = isdownloaded;
        } 

        public ExcelObject(string pdf_URL, bool isdownloaded, int rownumber, string bRnum)
        {
            Rownumber = rownumber;
            BRnum = bRnum;
            Pdf_URL = pdf_URL;
            Isdownloaded = isdownloaded;
        }
    }
}
