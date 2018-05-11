using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPari
{
    public class TextNumSortComparer : IComparer<string>
    {
        private abstract class ComparePart
        {
            public abstract int Compare(ComparePart other);
            public abstract ComparePart AddOrNewPart(char c);
        }

        private class StringComparePart : ComparePart
        {
            public string Part = string.Empty;
            public StringComparePart(string part)
            {
                if (part == null) throw new ArgumentNullException(paramName: "part");
                Part = part;
            }
            public override int Compare(ComparePart other)
            {
                if (other == null) throw new ArgumentNullException(paramName: "other");
                StringComparePart scp = other as StringComparePart;
                if (scp == null)
                {
                    return Part.CompareTo(((IntComparePart)other).Part.ToString());
                }
                else
                {
                    return Part.CompareTo(scp.Part);
                }
            }
            public override ComparePart AddOrNewPart(char c)
            {
                if (Char.IsDigit(c)) return new IntComparePart(c.ToString());
                Part += c;
                return null;
            }
        }

        private class IntComparePart : ComparePart
        {
            public string Part = string.Empty;
            public IntComparePart(string part)
            {
                if (part == null) throw new ArgumentNullException(paramName: "part");
                int test = int.Parse(part);
                Part = part;
            }
            public override int Compare(ComparePart other)
            {
                if (other == null) throw new ArgumentNullException(paramName: "other");
                IntComparePart icp = other as IntComparePart;
                if (icp == null)
                {
                    return String.Compare(this.Part, ((StringComparePart)other).Part, StringComparison.Ordinal);
                }
                else
                {
                    var otherPart = icp.Part;
                    if (Part.Length > otherPart.Length) return 1;
                    if (otherPart.Length > Part.Length ) return -1;

                    var arr = Part.ToArray();
                    var arr2 = otherPart.ToArray();
                    for (int i = 0; i < arr.Length; i++)
                    {
                        int x = int.Parse(arr[i].ToString());
                        int y = int.Parse(arr2[i].ToString());
                        if (x < y) return -1;
                        if (y < x) return 1;
                    }
                    return 0;
                }
            }
            public override ComparePart AddOrNewPart(char c)
            {
                if (!Char.IsDigit(c)) return new StringComparePart(c.ToString());
                Part += c;
                return null;
            }
        }

        public int Compare(string x, string y)
        {
            var partsX = GetCompareParts(x);
            var partsY = GetCompareParts(y);

            var cnt = Math.Min(partsX.Count, partsY.Count);
            for (int i = 0; i < cnt; i++ )
            {
                var res = partsX[i].Compare(partsY[i]);
                if (res != 0) return res;
            }

            return partsX.Count.CompareTo(partsY.Count);
        }

        private List<ComparePart> GetCompareParts(string txt)
        {
            var parts = new List<ComparePart>();
            if (txt == null) throw new ArgumentNullException(paramName: "txt");
            if (txt == string.Empty) 
            {
                parts.Add(new StringComparePart(""));
                return parts;
            }

            var topArr = txt.ToArray();
            
            ComparePart fcp = new StringComparePart("");
            var cp = fcp.AddOrNewPart(topArr[0]);
            if (cp == null) cp = fcp;
            parts.Add(cp);

            for (int i = 1; i < topArr.Length; i++)
            {
                var cp2 = cp.AddOrNewPart(topArr[i]);
                if (cp2 != null)
                {
                    parts.Add(cp2);
                    cp = cp2;
                }
            }
            return parts;
        }

    }
}
