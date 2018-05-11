using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPari
{
    public class TopAllgComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (x == y) return 0;
            if (x == "ALLG") return 1;
            if (y == "ALLG") return -1;

            return string.Compare(x, y);
        }
    }
}
