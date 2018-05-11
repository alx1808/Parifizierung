using InterfacesPari;
using System.Collections.Generic;

namespace AcadPari
{
    public class TableBuilder : TableHandler, ITableBuilder
    {
        private readonly IFactory _factory;

        public TableBuilder()
        {
            _factory = new FactoryPari.Factory();
        }

        private Dictionary<string, IKategorieRecord> _katDict = new Dictionary<string, IKategorieRecord>();
        public Dictionary<string, IKategorieRecord> KatDict
        {
            get { return _katDict; }
            set { _katDict = value; }
        }

        private List<IRaumRecord> _raumTable = new List<IRaumRecord>();
        public List<IRaumRecord> RaumTable
        {
            get { return _raumTable; }
            set { _raumTable = value; }
        }

        private List<IWohnungRecord> _wohnungTable = new List<IWohnungRecord>();

        public List<IWohnungRecord> WohnungTable
        {
            get { return _wohnungTable; }
            set { _wohnungTable = value; }
        }

        public void Build(List<IBlockInfo> blockInfos, List<IWohnungInfo> wohnungInfos)
        {
            foreach (var wi in wohnungInfos)
            {
                var wohnungRecord = new WohnungRecord();
                wohnungRecord.Top = wi.Top;
                wohnungRecord.Typ = wi.Typ;
                wohnungRecord.Widmung = wi.Widmung;
                wohnungRecord.Nutzwert = wi.Nutzwert;
                _wohnungTable.Add(wohnungRecord);
            }
            foreach (var bi in blockInfos)
            {
                var raumRecord = _factory.CreateRaumRecord(bi); // new  RaumRecord();
                //raumRecord.UpdateValuesFrom(bi);
                //raumRecord.Top = bi.Top;
                //raumRecord.Lage = bi.Geschoss;
                //raumRecord.Raum = bi.Raum;
                //raumRecord.RNW = bi.Nutzwert.Trim();
                //raumRecord.Begrundung = bi.Begrundung;
                //if (string.IsNullOrEmpty(raumRecord.RNW))
                //{
                //    raumRecord.Nutzwert = 1.0;
                //}
                //else
                //{
                //    double nutzwert;
                //    var rnw = raumRecord.RNW.Replace(',', '.');
                //    if (!double.TryParse(rnw, NumberStyles.Any, CultureInfo.CurrentCulture, out nutzwert))
                //    {
                //        throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Ungültiger Nutzwert {2} in {0}, Top {1}.", bi.Raum, bi.Top, bi.Nutzwert));
                //    }
                //    raumRecord.Nutzwert = nutzwert;
                //}

                //var m2s = GetM2(bi);
                //double m2;
                //if (!double.TryParse(m2s, out m2))
                //{
                //    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Ungültige Fläche {2} in {0}, Top {1}.", bi.Raum, bi.Top, bi.Flaeche));
                //}
                //raumRecord.Flaeche = m2;
                //raumRecord.AcadHandle = bi.Handle;

                var katIdent = raumRecord.KatIdentification;
                IKategorieRecord katRec;
                if (!_katDict.TryGetValue(katIdent, out katRec))
                {
                    katRec = _factory.CreateKategorie(raumRecord); // new KategorieRecord(raumRecord);
                    _katDict.Add(katIdent, katRec);
                }
                raumRecord.Kategorie = katRec;
                _raumTable.Add(raumRecord);
            }

            CheckNutzwertPerKatOk(_raumTable);
        }
    }
}
