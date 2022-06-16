using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceldatascript
{
      public interface ExcelInterface
    {
        public string? Pdf_URL { get; set; }

        public string Isdownloaded { get; set; }

        public int Rownumber { get; set; }

        public string BRnum { get; set; }
    }
}
