//using InterfacesPari;
//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace MdbPari
//{
//    public class SimpleRaumRecord : IRaumRecord
//    {
//        public int RaumId { get; set; }
//        public string Top { get; set; }
//        public string Lage { get; set; }
//        public string Raum { get; set; }
//        public string Widmung {get; set;}
//        public void UpdateValuesFrom(IBlockInfo acadBlockInfo)
//        {
//            Top = acadBlockInfo.Top;
//            Lage = acadBlockInfo.Geschoss;
//            Raum = acadBlockInfo.Raum;
//            RNW = acadBlockInfo.Nutzwert.Trim();
//            Begrundung = acadBlockInfo.Begrundung;
//            if (string.IsNullOrEmpty(RNW))
//            {
//                Nutzwert = 1.0;
//            }
//            else
//            {
//                double nutzwert;
//                var rnw = RNW.Replace(',', '.');
//                if (!double.TryParse(rnw, NumberStyles.Any, CultureInfo.CurrentCulture, out nutzwert))
//                {
//                    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Ungültiger Nutzwert {2} in {0}, Top {1}.", acadBlockInfo.Raum, acadBlockInfo.Top, acadBlockInfo.Nutzwert));
//                }
//                Nutzwert = nutzwert;
//            }
//            var m2S = GetM2(acadBlockInfo);
//            double m2;
//            if (!double.TryParse(m2S, out m2))
//            {
//                throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Ungültige Fläche {2} in {0}, Top {1}.", acadBlockInfo.Raum, acadBlockInfo.Top, acadBlockInfo.Flaeche));
//            }
//            Flaeche = m2;
//            AcadHandle = acadBlockInfo.Handle;
//        }

//        private static string GetM2(IBlockInfo bi)
//        {
//            var m2S = bi.Flaeche.Trim().Replace(".", "").Replace(',', '.');
//            int iom = m2S.IndexOf('m');
//            if (iom >= 0)
//            {
//                m2S = m2S.Substring(0, iom);
//            }
//            m2S = m2S.Trim();
//            return m2S;
//        }

//        public string RNW { get; set; }
//        private string _Begrundung = string.Empty;
//        public string Begrundung { get; set; }
//        public double Nutzwert { get; set; }
//        public double Flaeche { get; set; }
//        public int ProjektId { get; set; }
//        public int KategorieId { get; set; }
//        public string AcadHandle { get; set; }
//        public IKategorieRecord Kategorie { get; set; }
//        public string KatIdentification
//        {
//            get
//            {
//                return string.Join("|", new string[] { Top, Lage, Widmung, Begrundung });
//            }
//        }
//        public override string ToString()
//        {
//            return Top + "-" + Widmung + "-" + Raum;
//        }
//    }
//}
