using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MdbPari
{
    public class SimpleRaumZaRecord : IRaumZaRecord
    {
        public int RaumId { get; set; }
        public string Top { get; set; }
        public string Lage { get; set; }
        public string Raum { get; set; }
        public string Widmung { get; set; }
        public string RNW { get; set; }
        private string _Begrundung = string.Empty;
        public string Begrundung { get; set; }
        public double Nutzwert { get; set; }
        public double Flaeche { get; set; }
        public int ProjektId { get; set; }
        public int KategorieId { get; set; }
        public string AcadHandle { get; set; }
        public IKategorieZaRecord Kategorie { get; set; }
        public string KatIdentification
        {
            get
            {
                return string.Join("|", new string[] { Top, Lage, Widmung, Begrundung });
            }
        }
    }
}
