using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryPari
{
    public class KategorieZaRecord : KategorieRecord, IKategorieZaRecord
    {
        private List<IZuAbschlagRecord> _ZuAbschlaege = new List<IZuAbschlagRecord>();
        public List<IZuAbschlagRecord> ZuAbschlaege
        {
            get { return _ZuAbschlaege; }
            set { _ZuAbschlaege = value; }
        }
        public double SumProzent
        {
            get
            {
                return ZuAbschlaege.Sum(x => x.Prozent);
            }
        }
        public double ActualNutzwert
        {
            get { return this.Nutzwert + (this.Nutzwert * (this.SumProzent / 100)); }
        }
    }
}
