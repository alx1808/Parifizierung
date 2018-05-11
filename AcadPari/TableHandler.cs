using System;
using System.Collections.Generic;
using System.Linq;
using InterfacesPari;

namespace AcadPari
{
    public class TableHandler
    {
        public List<string> InvalidCategories
        {
            get { return _invalidCategories; }
            set { _invalidCategories = value; }
        }
        private List<string> _invalidCategories = new List<string>();

        public bool HasInvalidCategories
        {
            get { return InvalidCategories.Count > 0; }
        }

        public string JoinInvalidCatNames()
        {
            var catWithQuotes = InvalidCategories.Select(x => "'" + x + "'");
            var cats = string.Join("; ", catWithQuotes);
            return cats;
        }

        protected void CheckNutzwertPerKatOk(List<IRaumRecord> raume)
        {
            InvalidCategories.Clear();
            var raumePerKat = raume.GroupBy(x => x.KatIdentification);

            foreach (var raumGroup in raumePerKat)
            {
                var nutzwerte = new List<double>();
                foreach (var raumRecord in raumGroup)
                {
                    if (!ContainsDouble(nutzwerte, raumRecord.Nutzwert))
                    {
                        nutzwerte.Add(raumRecord.Nutzwert);
                    }
                }

                if (nutzwerte.Count > 1)
                {
                    InvalidCategories.Add(raumGroup.Key);
                }
            }
        }

        private const double CompareNutzwertEps = 0.0001;

        public static bool CompareNutzwert(double n1, double n2)
        {
            return Math.Abs(n1 - n2) <= CompareNutzwertEps;
        }

        private bool ContainsDouble(List<double> lst, double val)
        {
            return lst.Any(x => Math.Abs(x - val) <= CompareNutzwertEps);
        }
    }
}
