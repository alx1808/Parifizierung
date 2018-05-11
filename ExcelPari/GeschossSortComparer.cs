using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelPari
{
    internal class GeschossSortComparer : IComparer<string>
    {
        private List<string> _LageOrderRegKeys = new List<string>();

        public GeschossSortComparer()
        {
            _LageOrderRegKeys.Add("UG");
            _LageOrderRegKeys.Add("EG");
            _LageOrderRegKeys.Add("OG");
            for (int i = 0; i < 100; i++)
            {
                _LageOrderRegKeys.Add(i.ToString() + @"[^\d]*" + "OG");
            }
            _LageOrderRegKeys.Add("DG");
            _LageOrderRegKeys.Add("GA");
        }


        public int Compare(string x, string y)
        {
            var intX = GetLageOrderIndex(x);
            var intY = GetLageOrderIndex(y);
            if (intX < intY) return -1;
            if (intX > intY) return 1;

            return 0;
        }

        private int GetLageOrderIndex(string lage)
        {
            if (lage == null) return int.MaxValue;

            for (int i = 0; i < _LageOrderRegKeys.Count; i++)
            {
                if (Regex.IsMatch(lage, _LageOrderRegKeys[i], RegexOptions.IgnoreCase))
                {
                    return i;
                }
            }
            return int.MaxValue;
        }
    }
}
