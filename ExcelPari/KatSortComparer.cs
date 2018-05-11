using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelPari
{
    internal class KatSortComparer : IComparer<IKategorieRecord>
    {
        private List<string> _LageOrderRegKeys = new List<string>();

        public KatSortComparer()
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


        public int Compare(IKategorieRecord x, IKategorieRecord y)
        {
            var intX = GetLageOrderIndex(x.Lage);
            var intY = GetLageOrderIndex(y.Lage);
            if (intX < intY) return -1;
            if (intX > intY) return 1;

            intX = GetBegrundungOrderIndex(x.Begrundung);
            intY = GetBegrundungOrderIndex(y.Begrundung);
            if (intX < intY) return -1;
            if (intX > intY) return 1;

            return 0;            
        }

        private int GetBegrundungOrderIndex(string begrundung)
        {
            if (begrundung == null) return int.MaxValue;
            if (begrundung.ToUpperInvariant().Contains("WOHNUNG")) return 0;
            if (begrundung.ToUpperInvariant().Contains("ZUSCHLAG")) return 1;
            if (begrundung.ToUpperInvariant().Contains("ZUBEHÖR")) return 2;
            return int.MaxValue;
        }

        private int GetLageOrderIndex(string lage)
        {
            if (lage == null) return int.MaxValue;

            for (int i = 0; i < _LageOrderRegKeys.Count; i++)
            {
                if (Regex.IsMatch(lage,_LageOrderRegKeys[i], RegexOptions.IgnoreCase))
                {
                    return i;
                }
            }
            return int.MaxValue;
        }
    }
}
