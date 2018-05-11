using System.Globalization;
using InterfacesPari;

namespace FactoryPari
{
    public class KategorieRecord : IKategorieRecord
    {
        public KategorieRecord() { }

        public KategorieRecord(IRaumRecord raumRecord)
        {
            // folgende felder bilden die id einer kategorie
            Top = raumRecord.Top;
            Lage = raumRecord.Lage;
            Widmung = raumRecord.Widmung;
            Begrundung = raumRecord.Begrundung;

            RNW = raumRecord.RNW;
            if (string.IsNullOrEmpty(RNW))
            {
                RNW = "1,00";
            }
            Nutzwert = raumRecord.Nutzwert;
            ProjektId = raumRecord.ProjektId;
        }
        public int KategorieID { get; set; }
        public string Top { get; set; }
        public string Lage { get; set; }
        public string Widmung { get; set; }
        public string RNW { get; set; }
        public string Begrundung { get; set; }
        public double Nutzwert { get; set; }
        public int ProjektId { get; set; }
        public string Identification { get { return string.Join("|", new string[] { Top, Lage, Widmung, Begrundung }); } }
        public override string ToString()
        {
            return string.Format(CultureInfo.CurrentCulture,
                "ID: {0}, Top: {1}, Lage: {2}, Widmung: {3}, Begrundung: {4}, Nutzwert: {5}, ProjektID: {6}", KategorieID, Top, Lage,
                Widmung, Begrundung, Nutzwert, ProjektId);
        }
    }
}
