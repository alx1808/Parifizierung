using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using InterfacesPari;

namespace AcadPari
{
    public class TableUpdater : TableHandler, ITableUpdater
    {
        #region log4net Initialization
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(typeof(TableUpdater));
        #endregion

        private readonly IFactory _factory;

        public TableUpdater()
        {
            _factory = new FactoryPari.Factory();
        }

        private Dictionary<string, IKategorieRecord> _dbKatDict = new Dictionary<string, IKategorieRecord>();
        public Dictionary<string, IKategorieRecord> DbKatDict
        {
            get { return _dbKatDict; }
            set { _dbKatDict = value; }
        }

        public List<IKategorieRecord> NewKats
        {
            get { return _newKats; }
        }

        public List<IKategorieRecord> UpdKats
        {
            get { return _updKats; }
        }

        public List<IKategorieRecord> DelKats
        {
            get { return _delKats; }
            set { _delKats = value; }
        }

        public List<IRaumRecord> UpdRaume
        {
            get { return _updRaume; }
        }

        public List<IRaumRecord> DelRaume
        {
            get { return _delRaume; }
        }

        public List<IRaumRecord> NewRaume
        {
            get { return _newRaume; }
        }

        public List<IWohnungRecord> WohnungRecords
        {
            get { return _wohnungRecords; }
        }

        public void Update(List<IBlockInfo> blockInfos, List<IWohnungInfo> wohnungInfos, IProjektInfo projektInfo, IPariDatabase database)
        {
            ClearAll();

            var projektId = GetProjektId(projektInfo, database);

            GetWohnungRecords(wohnungInfos, projektId);

            // Kategorien aus Datenbank
            GetKategoriesFromDatabase(database, projektId);

            // AutoCAD-Blockinfo
            var blockInfoDict = new Dictionary<string, IBlockInfo>();
            foreach (var blockInfo in blockInfos)
            {
                blockInfoDict[blockInfo.Handle] = blockInfo;
            }

            // Räume aus Datenkbank
            var dbRaume = database.GetRaeume(projektId);
            // Ermitteln der zu löschenden Räume
            foreach (var raumRecord in dbRaume)
            {
                _raumDict[raumRecord.AcadHandle] = raumRecord;
                if (!blockInfoDict.ContainsKey(raumRecord.AcadHandle))
                {
                    DelRaume.Add(raumRecord);
                }
            }

            // updRaume und newRaume
            foreach (var blockInfo in blockInfos)
            {
                IRaumRecord raumRecord;
                if (_raumDict.TryGetValue(blockInfo.Handle, out raumRecord))
                {
                    UpdRaume.Add(raumRecord);
                    var rr2 = raumRecord.ShallowCopy();
                    raumRecord.UpdateValuesFrom(blockInfo);
                    if (!raumRecord.IsEqualTo(rr2))
                    {
                        ChangedRaumRecords.Add(raumRecord);
                        NrOfChangedRaumRecords++;
                    }
                }
                else
                {
                    raumRecord = _factory.CreateRaumRecord();
                    raumRecord.ProjektId = projektId;
                    NewRaume.Add(raumRecord);
                    raumRecord.UpdateValuesFrom(blockInfo);
                }

                IKategorieRecord kat;
                if (!_dbKatDict.TryGetValue(raumRecord.KatIdentification, out kat))
                {
                    kat = _factory.CreateKategorie(raumRecord);
                    kat.ProjektId = projektId;
                    NewKats.Add(kat);
                    raumRecord.Kategorie = kat;
                    _dbKatDict.Add(raumRecord.KatIdentification, kat);
                }
                else
                {
                    if (!CompareNutzwert(raumRecord.Nutzwert, kat.Nutzwert))
                    {
                        // create kat because of rnw-handling
                        var tmpkat = _factory.CreateKategorie(raumRecord);
                        kat.RNW = tmpkat.RNW;
                        kat.Nutzwert = tmpkat.Nutzwert;
                        UpdKats.Add(kat);
                    }
                    DelKats.Remove(kat);
                }

                raumRecord.Kategorie = kat;
            }

            CheckNutzwertPerKatOk();
        }

        private void CheckNutzwertPerKatOk()
        {
            var raume = NewRaume.Select(x => x).ToList();
            raume.AddRange(UpdRaume);
            CheckNutzwertPerKatOk(raume);
        }

        private readonly List<IKategorieRecord> _newKats = new List<IKategorieRecord>();
        private readonly List<IKategorieRecord> _updKats = new List<IKategorieRecord>();
        private List<IKategorieRecord> _delKats = new List<IKategorieRecord>();

        private readonly Dictionary<string, IRaumRecord> _raumDict = new Dictionary<string, IRaumRecord>();
        private readonly List<IRaumRecord> _updRaume = new List<IRaumRecord>();
        private readonly List<IRaumRecord> _delRaume = new List<IRaumRecord>();
        private readonly List<IRaumRecord> _newRaume = new List<IRaumRecord>();
        private readonly List<IWohnungRecord> _wohnungRecords = new List<IWohnungRecord>();
        private List<IRaumRecord> _changedRaumRecords = new List<IRaumRecord>();
        public int NrOfChangedRaumRecords { get; private set; }

        internal void LogStatus(Logger logger)
        {
            //if (NewKats.Count > 0)
            //{
            logger.Info(string.Format(CultureInfo.CurrentCulture, "Anzahl neuer Kategorien: {0}", NewKats.Count));
            //}
            //if (DelKats.Count > 0)
            //{
            logger.Info(string.Format(CultureInfo.CurrentCulture, "Anzahl zu löschender Kategorien: {0}", DelKats.Count));
            
            logger.Info(string.Format(CultureInfo.CurrentCulture, "Anzahl geänderter Kategorien: {0}", UpdKats.Count));

            //}
            //if (NewRaume.Count > 0)
            //{

            logger.Info(string.Format(CultureInfo.CurrentCulture, "Anzahl neuer Räume: {0}", NewRaume.Count));
            //}
            //if (DelRaume.Count > 0)
            //{
            logger.Info(string.Format(CultureInfo.CurrentCulture, "Anzahl zu löschender Räume: {0}", DelRaume.Count));
            //}

            //if (NrOfChangedRaumRecords > 0)
            //{
            logger.Info(string.Format(CultureInfo.CurrentCulture, "Anzahl geänderter Räume: {0}", NrOfChangedRaumRecords));
            foreach (var changedRaumRecord in ChangedRaumRecords)
            {
                logger.Info("\t" + changedRaumRecord);
            }
            //}

            if (HasInvalidCategories)
            {
                var cats = JoinInvalidCatNames();
                logger.Warn(string.Format(CultureInfo.CurrentCulture, "Kategorien mit abweichenden Nutzwerten: {0}", cats));
            }
        }

        public List<IRaumRecord> ChangedRaumRecords
        {
            get { return _changedRaumRecords; }
            set { _changedRaumRecords = value; }
        }

        private void ClearAll()
        {
            _dbKatDict.Clear();
            NewKats.Clear();
            DelKats.Clear();
            _raumDict.Clear();
            UpdRaume.Clear();
            DelRaume.Clear();
            NewRaume.Clear();
            WohnungRecords.Clear();
            InvalidCategories.Clear();
            ChangedRaumRecords.Clear();
            NrOfChangedRaumRecords = 0;
        }

        private void GetKategoriesFromDatabase(IPariDatabase database, int projektId)
        {
            var dbKats = database.GetKategories(projektId);
            foreach (var kat in dbKats)
            {
                _dbKatDict[kat.Identification] = kat;
            }

            DelKats = dbKats.Select(x => x).ToList();
        }

        private void GetWohnungRecords(List<IWohnungInfo> wohnungInfos, int projektId)
        {
            foreach (var wi in wohnungInfos)
            {
                var wohnungRecord = new WohnungRecord();
                wohnungRecord.Top = wi.Top;
                wohnungRecord.Typ = wi.Typ;
                wohnungRecord.Widmung = wi.Widmung;
                wohnungRecord.Nutzwert = wi.Nutzwert;
                wohnungRecord.ProjektId = projektId;
                WohnungRecords.Add(wohnungRecord);
            }
        }

        private static int GetProjektId(IProjektInfo projektInfo, IPariDatabase database)
        {
            var projektId = database.GetProjektId(projektInfo);
            if (projektId < 0)
            {
                var msg = string.Format(CultureInfo.CurrentCulture, "Das Projekt '{0}' existiert nicht!",
                    projektInfo.Bauvorhaben);
                Log.Error(msg);
                throw new InvalidOperationException(msg);
            }

            return projektId;
        }
    }
}
