using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CsvToExcelTransformer_Win
{
    class CsvLine
    {
        public int Id;
        public string Date;
        public string Status;
        public string Name;
        public string Lastname;
        public string Email;
        public string Tel;
        public string Interest;
        public bool FirstAgr;
        public bool SecondAgr;
        public bool ThirdAgr;

        public static CsvLine FromCsv(string line)
        {
            string[] strings = line.Replace("\"", "").Split(',');
            CsvLine csv = new CsvLine();
            csv.Id = int.Parse(strings[0]);
            csv.Date = strings[1];
            csv.Status = strings[2];
            csv.Name = strings[3].Trim();
            csv.Lastname = strings[4].Trim();
            csv.Email = strings[5];
            csv.Tel = strings[6];
            string flat = string.Empty;
            if (!string.IsNullOrEmpty(strings[7]) && (!string.IsNullOrEmpty(strings[11])))
            {
                flat = string.Join(',', strings[7], strings[11]);
            }
            else if (!string.IsNullOrEmpty(strings[7]) && (string.IsNullOrEmpty(strings[11])))
            {
                flat = strings[7];
            }
            else if (string.IsNullOrEmpty(strings[7]) && (!string.IsNullOrEmpty(strings[11])))
            {
                flat = strings[11];
            }
            bool fAgr;
            bool sAgr;
            bool tAgr;
            if (!string.IsNullOrEmpty(strings[8])) { fAgr = true; } else { fAgr = false; };
            if (!string.IsNullOrEmpty(strings[9])) { sAgr = true; } else { sAgr = false; };
            if (!string.IsNullOrEmpty(strings[10])) { tAgr = true; } else { tAgr = false; };
            csv.Interest = flat;
            csv.FirstAgr = fAgr;
            csv.SecondAgr = sAgr;
            csv.ThirdAgr = tAgr;
            return csv;
        }
    }
}
