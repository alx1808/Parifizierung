//using InterfacesPari;
//using System;
//using System.Globalization;

//namespace AcadPari
//{
//    public class RaumRecord : IRaumRecord
//    {
//        public int RaumId { get; set; }
//        public string Top { get; set; }
//        public string Lage { get; set; }
//        public string Raum { get; set; }
//        public string Widmung
//        {
//            get
//            {
//                if (string.IsNullOrEmpty(RNW))
//                {
//                    return "Wohnung";
//                }
//                else if (_begrundung.ToUpper().Contains("PKW"))
//                {
//                    return _begrundung;
//                }
//                else
//                {
//                    return Raum;
//                }
//            }
//            set { ; }
//        }

//        public void UpdateValuesFrom(IBlockInfo acadBlockInfo)
//        {
//            throw new NotImplementedException();
//        }

//        public string RNW { get; set; }
//        private string _begrundung = string.Empty;
//        public string Begrundung
//        {
//            get {
//                if (_begrundung == null || string.IsNullOrEmpty(_begrundung.Trim()) || _begrundung.ToUpper().Contains("PKW"))
//                {
//                    return "als Wohnungseigentumsobjekt";
//                }
//                else
//                {
//                    return _begrundung;
//                }
//            }
//            set { _begrundung = value; }
//        }
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
//    }
//}
