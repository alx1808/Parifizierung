using InterfacesPari;
using MdbPari.Properties;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;

namespace MdbPari
{
    public class MdbWriter : IPariDatabase
    {
        #region log4net Initialization
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(MdbWriter));
        static MdbWriter()
        {
            if (log4net.LogManager.GetRepository(System.Reflection.Assembly.GetExecutingAssembly()).Configured == false)
            {
                log4net.Config.XmlConfigurator.ConfigureAndWatch(
                    new System.IO.FileInfo(
                        System.IO.Path.Combine(
                            new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName,
                            "_log4net.config"
                        )
                    )
                );
            }
        }
        #endregion

        private IFactory _Factory = null;

        #region Lifecycle
        public MdbWriter(IFactory factory)
        {
            MdbName = Settings.Default.MdbName;
            log.Debug(string.Format(CultureInfo.InvariantCulture, "MdbWriter: '{0}'", MdbName));
            _Factory = factory;
        }
        public void SetDatabase(object dbo)
        {
            if (dbo == null) throw new ArgumentNullException(paramName: "dbo");
            var mdbName = dbo.ToString();
            if (!System.IO.File.Exists(mdbName))
            {
                throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Datei '{0}' nicht gefunden!", mdbName));
            }

            MdbName = mdbName;
            Settings.Default.MdbName = MdbName;
            Settings.Default.Save();
        }

        #endregion

        #region Public
        public string MdbName { get; set; }
        private string _ConnectionStringFormat = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Persist Security Info=False;";
        public string ConnectionStringFormat
        {
            get { return _ConnectionStringFormat; }
            set { _ConnectionStringFormat = value; }
        }
        #endregion

        private string ConnectionString
        {
            get
            {
                //return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + MdbName + ";Persist Security Info=False;";
                return string.Format(CultureInfo.InvariantCulture, _ConnectionStringFormat, MdbName);
            }
        }

        #region IPariDatabase
        public List<string> GetTableNames()
        {
            log.Debug("GetTableNames");
            try
            {
                List<string> tableNames = new List<string>();

                // Open OleDb Connection
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    //myConnection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + MdbName + ";Persist Security Info=False;"; // ConnectionString;
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();

                    var schema = myConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    foreach (var row in schema.Rows.OfType<DataRow>())
                    {
                        string tableName = row.ItemArray[2].ToString();
                        tableNames.Add(tableName);
                    }
                    return tableNames;
                }

            }
            catch (Exception ex)
            {
                log.Debug(ConnectionString);
                log.Error(ex.Message);
                throw new InvalidOperationException(ex.Message);
            }
        }

        public List<IRaumZaRecord> GetRaeumeWithZuAbschlag(int projektId)
        {
            log.Debug("GetRaeumeWithZuAbschlag");
            var kats = this.GetKategoriesWithZuAbschlag(projektId);
            var dict = new Dictionary<int, IKategorieZaRecord>();
            foreach (var kat in kats)
            {
                dict.Add(kat.KategorieID, kat);
            }
            var raume = GetRaeumeZaOhneKategorie(projektId);
            foreach (var rok in raume)
            {
                // Todo: passende Exception, wenn rok.KategorieId nicht existiert!
                IKategorieZaRecord katRec = null;
                if (!dict.TryGetValue(rok.KategorieId, out katRec))
                {
                    var msg = string.Format(CultureInfo.InvariantCulture, "Raum {0} has invalid KategorieId {1}!", rok.RaumId);
                    throw new InvalidOperationException(msg);
                }
                rok.Kategorie = katRec;
            }
            return raume;
        }

        public List<IRaumRecord> GetRaeume(int projektId)
        {
            log.Debug("GetRaeume");
            var kats = this.GetKategories(projektId);
            var dict = new Dictionary<int, IKategorieRecord>();
            foreach (var kat in kats)
            {
                dict.Add(kat.KategorieID, kat);
            }
            var raume = GetRaeumeOhneKategorie(projektId);
            foreach (var rok in raume)
            {
                // Todo: passende Exception, wenn rok.KategorieId nicht existiert!
                IKategorieRecord katRec = null;
                if (!dict.TryGetValue(rok.KategorieId, out katRec))
                {
                    var msg = string.Format(CultureInfo.InvariantCulture, "Raum {0} has invalid KategorieId {1}!", rok.RaumId, rok.KategorieId);
                    throw new InvalidOperationException(msg);
                }
                rok.Kategorie = katRec;
            }
            return raume;
        }

        public int UpdateRaum(IRaumRecord raum, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            // cmd.CommandText = "INSERT INTO Raum " + "([Top],[Lage],[Raum],[Widmung],[RNW],[Begrundung],[Nutzwert],[Flaeche],[ProjektId],[KategorieId],[AcadHandle]) "
            log.Debug("UpdateRaum");
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "UPDATE {0} SET Raum.[Top]=@Top,Raum.Lage=@Lage,Raum.Raum=@Raum,Raum.Widmung=@Widmung,Raum.RNW=@RNW,Raum.Begrundung=@Begrundung,Raum.Nutzwert=@Nutzwert,Raum.Flaeche=@Flaeche,Raum.KategorieId=@KategorieId where Raum.ProjektId=@ProjektId AND Raum.RaumID=@RaumID", "Raum");
                cmd.Parameters.AddRange(new OleDbParameter[]
                {
                new OleDbParameter("@Top", raum.Top??Convert.DBNull),
                new OleDbParameter("@Lage", raum.Lage??Convert.DBNull),
                new OleDbParameter("@Raum", raum.Raum??Convert.DBNull),
                new OleDbParameter("@Widmung", raum.Widmung??Convert.DBNull),
                new OleDbParameter("@RNW", raum.RNW??Convert.DBNull),
                new OleDbParameter("@Begrundung", raum.Begrundung??Convert.DBNull),
                new OleDbParameter("@Nutzwert", raum.Nutzwert),
                new OleDbParameter("@Flaeche", raum.Flaeche),
                new OleDbParameter("@KategorieId", raum.KategorieId),
                new OleDbParameter("@ProjektId", raum.ProjektId),
                new OleDbParameter("@RaumID", raum.RaumId),
                });
                var result = cmd.ExecuteNonQuery();
                return result;
            }
        }

        public int UpdateKategorie(IKategorieRecord kategorie)
        {
            log.Debug("UpdateKategorie");
            int result;
            var katInfos = new List<IKategorieRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "UPDATE {0} SET Nutzwert=@Nutzwert where ProjektId=@ProjektId AND KategorieId=@KategorieId", "KATEGORIE");
                    cmd.Parameters.Add(new OleDbParameter("@Nutzwert", kategorie.Nutzwert));
                    cmd.Parameters.Add(new OleDbParameter("@ProjektId", kategorie.ProjektId));
                    cmd.Parameters.Add(new OleDbParameter("@KategorieId", kategorie.KategorieID));
                    result = cmd.ExecuteNonQuery();
                }
                myConnection.Close();
            }
            return result;
        }


        private int UpdateKategorie(IKategorieRecord kategorie, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "UPDATE {0} SET Nutzwert=@Nutzwert,RNW=@RNW where ProjektId=@ProjektId AND KategorieId=@KategorieId", "KATEGORIE");
                cmd.Parameters.Add(new OleDbParameter("@Nutzwert", kategorie.Nutzwert));
                cmd.Parameters.Add(new OleDbParameter("@RNW", kategorie.RNW));
                cmd.Parameters.Add(new OleDbParameter("@ProjektId", kategorie.ProjektId));
                cmd.Parameters.Add(new OleDbParameter("@KategorieId", kategorie.KategorieID));
                return cmd.ExecuteNonQuery();
            }
        }

        public int UpdateZuAbschlag(IZuAbschlagRecord zuAbschlagRec)
        {
            log.Debug("UpdateZuAbschlag");
            int result;
            var katInfos = new List<IKategorieRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "UPDATE {0} SET Prozent=@Prozent,Beschreibung=@Beschreibung where ZuAbschlagId=@ZuAbschlagId", ZU_ABSCHLAG_TABELLE);
                    cmd.Parameters.Add(new OleDbParameter("@Prozent", zuAbschlagRec.Prozent));
                    cmd.Parameters.Add(new OleDbParameter("@Beschreibung", zuAbschlagRec.Beschreibung ?? Convert.DBNull));
                    cmd.Parameters.Add(new OleDbParameter("@ZuAbschlagId", zuAbschlagRec.ZuAbschlagId));
                    result = cmd.ExecuteNonQuery();
                }
                myConnection.Close();
            }
            return result;
        }

        //public List<IZuAbschlagRecord> GetZuAbschlagsForProjekt(int projektId)
        //{
        //    var zaInfos = new List<IZuAbschlagRecord>();
        //    using (OleDbConnection myConnection = new OleDbConnection())
        //    {
        //        myConnection.ConnectionString = ConnectionString;
        //        myConnection.Open();

        //        var cmd = myConnection.CreateCommand();
        //        cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from {0} where ProjektId={1} order by ZuAbschlagID", "ZuAbschlag", projektId);
        //        var reader = cmd.ExecuteReader();
        //        while (reader.Read())
        //        {
        //            var za = new SimpleZuAbschlagRecord();
        //            object o = reader["Beschreibung"];
        //            if (o != System.DBNull.Value) za.Beschreibung = o.ToString();
        //            o = reader["Prozent"];
        //            if (o != System.DBNull.Value) za.Prozent = (double)o;
        //            o = reader["ProjektId"];
        //            if (o != System.DBNull.Value) za.ProjektId = (int)o;
        //            o = reader["KategorieId"];
        //            if (o != System.DBNull.Value) za.KategorieId = (int)o;
        //            o = reader["ZuAbschlagId"];
        //            if (o != System.DBNull.Value) za.ZuAbschlagId = (int)o;

        //            zaInfos.Add(za);
        //        }
        //        myConnection.Close();
        //    }
        //    return zaInfos;
        //}
        public List<IZuAbschlagRecord> GetZuAbschlags(int kategorieId)
        {
            var zaInfos = new List<IZuAbschlagRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                GetZuAbschlags(kategorieId, zaInfos, myConnection);
                myConnection.Close();
            }
            return zaInfos;
        }

        public List<IZuAbschlagVorgabeRecord> GetZuAbschlagVorgaben()
        {
            var zaInfos = new List<IZuAbschlagVorgabeRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                GetZuAbschlagVorgaben(zaInfos, myConnection);
                myConnection.Close();
            }
            return zaInfos;
        }

        public List<IWohnungRecord> GetWohnungen(int projektId)
        {
            log.Debug("GetWohnungen");
            var wohnungInfos = new List<IWohnungRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from {0} where ProjektId={1}", "Wohnung", projektId);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var wohnung = new SimpleWohnungRecord();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) wohnung.ProjektId = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) wohnung.Top = o.ToString();

                        o = reader["Typ"];
                        if (o != System.DBNull.Value) wohnung.Typ = o.ToString();

                        o = reader["WohnungId"];
                        if (o != System.DBNull.Value) wohnung.WohnungId = (int)o;

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) wohnung.Widmung = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) wohnung.Nutzwert = o.ToString();

                        wohnungInfos.Add(wohnung);
                    }
                    reader.Close();
                }

                myConnection.Close();
            }
            return wohnungInfos;
        }

        public List<IKategorieRecord> GetKategories(int ProjektId)
        {
            log.Debug("GetKategories");
            var katInfos = new List<IKategorieRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from {0} where ProjektId={1} order by Top", "Kategorie", ProjektId);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var kat = _Factory.CreateKategorie();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) kat.ProjektId = (int)o;

                        o = reader["KategorieID"];
                        if (o != System.DBNull.Value) kat.KategorieID = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) kat.Top = o.ToString();

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) kat.Widmung = o.ToString();

                        o = reader["RNW"];
                        if (o != System.DBNull.Value) kat.RNW = o.ToString();

                        o = reader["Begrundung"];
                        if (o != System.DBNull.Value) kat.Begrundung = o.ToString();

                        o = reader["Lage"];
                        if (o != System.DBNull.Value) kat.Lage = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) kat.Nutzwert = Convert.ToDouble(o);

                        katInfos.Add(kat);
                    }
                    reader.Close();
                }
                myConnection.Close();
            }
            return katInfos;
        }

        public List<IKategorieZaRecord> GetKategoriesWithZuAbschlag(int ProjektId)
        {
            log.Debug("GetKategoriesWithZuAbschlag");
            var katInfos = new List<IKategorieZaRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from {0} where ProjektId={1} order by Top", "Kategorie", ProjektId);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var kat = _Factory.CreateZaKategorie();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) kat.ProjektId = (int)o;

                        o = reader["KategorieID"];
                        if (o != System.DBNull.Value) kat.KategorieID = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) kat.Top = o.ToString();

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) kat.Widmung = o.ToString();

                        o = reader["RNW"];
                        if (o != System.DBNull.Value) kat.RNW = o.ToString();

                        o = reader["Begrundung"];
                        if (o != System.DBNull.Value) kat.Begrundung = o.ToString();

                        o = reader["Lage"];
                        if (o != System.DBNull.Value) kat.Lage = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) kat.Nutzwert = Convert.ToDouble(o);

                        katInfos.Add(kat);
                    }
                    reader.Close();
                }
                myConnection.Close();
            }

            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();
                foreach (var kat in katInfos)
                {
                    var zaInfos = new List<IZuAbschlagRecord>();
                    GetZuAbschlags(kat.KategorieID, zaInfos, myConnection);
                    kat.ZuAbschlaege = zaInfos;
                }
                myConnection.Close();
            }
            return katInfos;
        }

        public List<IProjektInfo> ListProjInfos()
        {
            log.Debug("ListProjInfos");
            var projInfos = new List<IProjektInfo>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = "select * from Projekt order by ProjektID";
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var pi = _Factory.CreateProjectInfo();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) pi.ProjektId = (int)o;

                        o = reader["Bauvorhaben"];
                        if (o != System.DBNull.Value) pi.Bauvorhaben = o.ToString();

                        o = reader["DwgName"];
                        if (o != System.DBNull.Value) pi.DwgName = o.ToString();

                        o = reader["EZ"];
                        if (o != System.DBNull.Value) pi.EZ = o.ToString();

                        projInfos.Add(pi);
                    }
                    reader.Close();
                }
                myConnection.Close();
            }
            return projInfos;
        }

        public int GetProjektId(IProjektInfo projektInfo)
        {
            log.Debug("GetProjektId");
            int projektId = -1;

            // Open OleDb Connection
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                projektId = GetProjektId(projektInfo, myConnection);

                myConnection.Close();
            }
            return projektId;
        }

        public int DeleteKategorienAndZuAbschlag(List<IKategorieRecord> kategories, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Debug("DeleteKategorienAndZuAbschlag");
            int nrOfDeletedRows = 0;
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                foreach (var kategorieRecord in kategories)
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from {0} where KategorieId={1} AND ProjektId={2}", ZU_ABSCHLAG_TABELLE, kategorieRecord.KategorieID, kategorieRecord.ProjektId);
                    nrOfDeletedRows += cmd.ExecuteNonQuery();
                }

                foreach (var kategorieRecord in kategories)
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from {0} where KategorieId={1} AND ProjektId={2}", "Kategorie", kategorieRecord.KategorieID, kategorieRecord.ProjektId);
                    nrOfDeletedRows += cmd.ExecuteNonQuery();
                }
            }
            return nrOfDeletedRows;
        }

        public int DeleteZuAbschlag(IZuAbschlagRecord zuAbschlagRec)
        {
            log.Debug("DeleteKategorienAndZuAbschlag");
            int nrOfDeletedRows = 0;
            // Open OleDb Connection
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                // Execute Queries
                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from {0} where ZuAbschlagId={1}", ZU_ABSCHLAG_TABELLE, zuAbschlagRec.ZuAbschlagId);
                    nrOfDeletedRows = cmd.ExecuteNonQuery();
                }
                myConnection.Close();
            }

            return nrOfDeletedRows;
        }

        public int DeleteProjekt(int projektId)
        {
            log.Debug("DeleteProjekt");
            int NrOfDeletedRows = 0;
            // Open OleDb Connection
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();
                var transaction = myConnection.BeginTransaction();

                // Execute Queries
                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from Projekt where ProjektId={0}", projektId);
                    NrOfDeletedRows += cmd.ExecuteNonQuery();
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from Kategorie where ProjektId={0}", projektId);
                    NrOfDeletedRows += cmd.ExecuteNonQuery();
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from Raum where ProjektId={0}", projektId);
                    NrOfDeletedRows += cmd.ExecuteNonQuery();
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from GstInfo where ProjektId={0}", projektId);
                    NrOfDeletedRows += cmd.ExecuteNonQuery();
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from ZuAbschlag where ProjektId={0}", projektId);
                    NrOfDeletedRows += cmd.ExecuteNonQuery();
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from Wohnung where ProjektId={0}", projektId);
                    NrOfDeletedRows += cmd.ExecuteNonQuery();

                    transaction.Commit();
                }
                myConnection.Close();
            }

            return NrOfDeletedRows;
        }
        private const string ZU_ABSCHLAG_TABELLE = "ZuAbschlag";
        public int InsertZuAbschlag(IZuAbschlagRecord zuAbschlag)
        {
            log.Debug("InsertZuAbschlag");
            int zuAbschlagIdAfterInsert = -1;
            // Open OleDb Connection
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();
                var transaction = myConnection.BeginTransaction();

                int zuAbschlagIdBeforeInsert = GetHighestId(zuAbschlag, myConnection, transaction);
                if (InsertZuAbschlag(zuAbschlag, myConnection, transaction) <= 0)
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Konnte Record nicht zu Tabelle '{0}' hinzufügen!", ZU_ABSCHLAG_TABELLE));
                }
                zuAbschlagIdAfterInsert = GetHighestId(zuAbschlag, myConnection, transaction);
                if (zuAbschlagIdBeforeInsert == zuAbschlagIdAfterInsert)
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Konnte Record nicht zu Tabelle '{0}' hinzufügen! Id nicht gefunden.", ZU_ABSCHLAG_TABELLE));
                }
                zuAbschlag.ZuAbschlagId = zuAbschlagIdAfterInsert;

                transaction.Commit();
                myConnection.Close();
            }
            return zuAbschlagIdAfterInsert;
        }

        public void UpdateDatabase(ITableUpdater tableUpdater, IProjektInfo projektInfo)
        {
            log.Info("UpdateDatabase");
            var projektId = GetProjektId(projektInfo);
            if (projektId < 0)
            {
                throw new InvalidOperationException("Projekt existiert bereits!");
            }

            try
            {
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();
                    var transaction = myConnection.BeginTransaction();

                    //InsertProjekt(projektInfo, myConnection, transaction);
                    //projektId = GetProjektId(projektInfo, myConnection, transaction);
                    //projektInfo.ProjektId = projektId;
                    //InsertSubInfos(projektInfo, myConnection, transaction);

                    foreach (var kat in tableUpdater.NewKats)
                    {
                        InsertKategorieAndSetKategorieIdToKat(kat, myConnection, transaction);
                    }

                    DeleteKategorienAndZuAbschlag(tableUpdater.DelKats, myConnection, transaction);

                    foreach (var kat in tableUpdater.UpdKats)
                    {
                        UpdateKategorie(kat, myConnection, transaction);
                    }

                    foreach (var raum in tableUpdater.NewRaume)
                    {
                        raum.KategorieId = raum.Kategorie.KategorieID;
                        InsertRaum(raum, myConnection, transaction);
                    }

                    foreach (var raum in tableUpdater.UpdRaume)
                    {
                        raum.KategorieId = raum.Kategorie.KategorieID;
                        UpdateRaum(raum, myConnection, transaction);
                    }

                    DelRaeume(tableUpdater.DelRaume, projektId, myConnection, transaction);

                    // update wohnung
                    DelWohnungen(projektId, myConnection, transaction);
                    foreach (var wohnung in tableUpdater.WohnungRecords)
                    {
                        InsertWohnung(wohnung, myConnection, transaction);
                    }

                    transaction.Commit();
                    myConnection.Close();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }
        }

        public void SaveToDatabase(ITableBuilder tableBuilder, IProjektInfo projektInfo)
        {
            log.Info("SaveToDatabase");
            var projektId = GetProjektId(projektInfo);
            if (projektId >= 0)
            {
                throw new InvalidOperationException("Projekt existiert bereits!");
            }

            try
            {
                // Open OleDb Connection
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();
                    var transaction = myConnection.BeginTransaction();

                    InsertProjekt(projektInfo, myConnection, transaction);
                    projektId = GetProjektId(projektInfo, myConnection, transaction);
                    projektInfo.ProjektId = projektId;
                    InsertSubInfos(projektInfo, myConnection, transaction);

                    foreach (var kat in tableBuilder.KatDict.Values)
                    {
                        kat.ProjektId = projektId;
                        InsertKategorieAndSetKategorieIdToKat(kat, myConnection, transaction);
                    }

                    foreach (var raum in tableBuilder.RaumTable)
                    {
                        raum.ProjektId = projektId;
                        raum.KategorieId = raum.Kategorie.KategorieID;
                        InsertRaum(raum, myConnection, transaction);
                    }

                    foreach (var wohnung in tableBuilder.WohnungTable)
                    {
                        wohnung.ProjektId = projektId;
                        InsertWohnung(wohnung, myConnection, transaction);
                    }

                    transaction.Commit();
                    myConnection.Close();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }
        }

        #region Add Fields to Tables



        private class ColumnInfo
        {
            public string ColumnName;
            public string ColumnType;

            public ColumnInfo(string columnName, string columnType)
            {
                ColumnName = columnName;
                ColumnType = columnType;
            }
        }

        private class TableInfo
        {
            public string TableName;
            public List<ColumnInfo> ColumnInfos = new List<ColumnInfo>();

            public TableInfo(string tableName, List<ColumnInfo> columnInfos)
            {
                TableName = tableName;
                ColumnInfos = columnInfos;
            }
        }

        private List<TableInfo> _fieldsToCheckInTables = new List<TableInfo>()
        {
            new TableInfo("Wohnung", new List<ColumnInfo>()
            {
                new ColumnInfo("Widmung", "TEXT(200)"),
                new ColumnInfo("Nutzwert","TEXT(10)")
            })
        };


        public int CheckExistingFields()
        {
            int nrOfAddedFields = 0;
            foreach (var tableInfo in _fieldsToCheckInTables)
            {
                foreach (var columnInfo in tableInfo.ColumnInfos)
                {
                    if (!ColumnExists(tableInfo.TableName, columnInfo.ColumnName))
                    {
                        using (OleDbConnection myConnection = new OleDbConnection())
                        {
                            myConnection.ConnectionString = ConnectionString;
                            myConnection.Open();

                            using (var cmd = myConnection.CreateCommand())
                            {
                                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "ALTER TABLE {0} ADD COLUMN {1} {2}", tableInfo.TableName, columnInfo.ColumnName, columnInfo.ColumnType);
                                cmd.ExecuteNonQuery();
                            }
                            myConnection.Close();
                        }

                        nrOfAddedFields++;
                    }
                }
            }

            return nrOfAddedFields;
        }

        private bool ColumnExists(string tableName, string colName)
        {
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();
                var schema = myConnection.GetSchema("COLUMNS");
                var col = schema.Select("TABLE_NAME='" + tableName + "' AND COLUMN_NAME='" + colName + "'");

                if (col.Length > 0) return true;
                return false;
            }
        }
        #endregion

        #endregion

        #region Private
        private int InsertZuAbschlag(IZuAbschlagRecord zuAbschlag, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "INSERT INTO {0} " + "([Beschreibung],[Prozent],[KategorieId],[ProjektId]) "
                    + "VALUES(@Beschreibung,@Prozent,@KategorieId, @ProjektId)", ZU_ABSCHLAG_TABELLE);

                // add named parameters
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Beschreibung", zuAbschlag.Beschreibung??Convert.DBNull),
                               new OleDbParameter("@Prozent", zuAbschlag.Prozent),
                               new OleDbParameter("@KategorieId", zuAbschlag.KategorieId),
                               new OleDbParameter("@ProjektId", zuAbschlag.ProjektId),
                               });

                return cmd.ExecuteNonQuery();
            }
        }

        private int GetHighestId(IZuAbschlagRecord zuAbschlag, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            var zuAbschlagId = -1;
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                if (transaction != null) cmd.Transaction = transaction;
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture,
                    "SELECT ZuAbschlagId FROM {0} where KategorieId=@KategorieId AND ProjektId=@ProjektId ORDER BY ZuAbschlagID DESC",
                    ZU_ABSCHLAG_TABELLE);
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@KategorieId", zuAbschlag.KategorieId),
                               new OleDbParameter("@ProjektId", zuAbschlag.ProjektId),
                               });


                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    object o = reader["ZuAbschlagId"];
                    zuAbschlagId = (int)o;
                }
                reader.Close();
            }
            return zuAbschlagId;
        }

        private void GetZuAbschlagVorgaben(List<IZuAbschlagVorgabeRecord> zaInfos, OleDbConnection myConnection)
        {
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from {0}", "ZuAbschlagVorgabe");
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var za = new SimpleZuAbschlagVorgabeRecord();
                    object o = reader["Beschreibung"];
                    if (o != System.DBNull.Value) za.Beschreibung = o.ToString();
                    o = reader["Prozent"];
                    if (o != System.DBNull.Value) za.Prozent = (double)o;
                    zaInfos.Add(za);
                }
                reader.Close();
            }
        }

        private static void GetZuAbschlags(int kategorieId, List<IZuAbschlagRecord> zaInfos, OleDbConnection myConnection)
        {
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from {0} where KategorieId={1} order by ZuAbschlagID", "ZuAbschlag", kategorieId);
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var za = new SimpleZuAbschlagRecord();
                    object o = reader["Beschreibung"];
                    if (o != System.DBNull.Value) za.Beschreibung = o.ToString();
                    o = reader["Prozent"];
                    if (o != System.DBNull.Value) za.Prozent = (double)o;
                    o = reader["ProjektId"];
                    if (o != System.DBNull.Value) za.ProjektId = (int)o;
                    o = reader["KategorieId"];
                    if (o != System.DBNull.Value) za.KategorieId = (int)o;
                    o = reader["ZuAbschlagId"];
                    if (o != System.DBNull.Value) za.ZuAbschlagId = (int)o;

                    zaInfos.Add(za);
                }
                reader.Close();
            }
        }

        private List<IRaumZaRecord> GetRaeumeZaOhneKategorie(int ProjektId)
        {
            log.Debug("GetRaeumeZaOhneKategorie");
            var raumInfos = new List<IRaumZaRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from Raum where ProjektId={0} order by Top", ProjektId);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var raum = new SimpleRaumZaRecord();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) raum.ProjektId = (int)o;

                        o = reader["KategorieId"];
                        if (o != System.DBNull.Value) raum.KategorieId = (int)o;

                        o = reader["RaumId"];
                        if (o != System.DBNull.Value) raum.RaumId = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) raum.Top = o.ToString();

                        o = reader["Lage"];
                        if (o != System.DBNull.Value) raum.Lage = o.ToString();

                        o = reader["Raum"];
                        if (o != System.DBNull.Value) raum.Raum = o.ToString();

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) raum.Widmung = o.ToString();

                        o = reader["RNW"];
                        if (o != System.DBNull.Value) raum.RNW = o.ToString();

                        o = reader["Begrundung"];
                        if (o != System.DBNull.Value) raum.Begrundung = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) raum.Nutzwert = Convert.ToDouble(o);

                        o = reader["Flaeche"];
                        if (o != System.DBNull.Value) raum.Flaeche = Convert.ToDouble(o);

                        raumInfos.Add(raum);
                    }
                    reader.Close();
                }
                myConnection.Close();
            }
            return raumInfos;
        }

        private List<IRaumRecord> GetRaeumeOhneKategorie(int ProjektId)
        {
            log.Debug("GetRaeumeOhneKategorie");
            var raumInfos = new List<IRaumRecord>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConnectionString;
                myConnection.Open();

                using (var cmd = myConnection.CreateCommand())
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select * from Raum where ProjektId={0} order by Top", ProjektId);
                    var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var raum = _Factory.CreateRaumRecord(); // new SimpleRaumRecord();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) raum.ProjektId = (int)o;

                        o = reader["KategorieId"];
                        if (o != System.DBNull.Value) raum.KategorieId = (int)o;

                        o = reader["RaumId"];
                        if (o != System.DBNull.Value) raum.RaumId = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) raum.Top = o.ToString();

                        o = reader["Lage"];
                        if (o != System.DBNull.Value) raum.Lage = o.ToString();

                        o = reader["Raum"];
                        if (o != System.DBNull.Value) raum.Raum = o.ToString();

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) raum.Widmung = o.ToString();

                        o = reader["RNW"];
                        if (o != System.DBNull.Value) raum.RNW = o.ToString();

                        o = reader["Begrundung"];
                        if (o != System.DBNull.Value) raum.Begrundung = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) raum.Nutzwert = Convert.ToDouble(o);

                        o = reader["Flaeche"];
                        if (o != System.DBNull.Value) raum.Flaeche = Convert.ToDouble(o);

                        o = reader["AcadHandle"];
                        if (o != System.DBNull.Value) raum.AcadHandle = o.ToString();
                        raumInfos.Add(raum);
                    }
                    reader.Close();
                }
                myConnection.Close();
            }
            return raumInfos;
        }

        private int DelRaeume(List<IRaumRecord> raume, int projektId, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Debug("DelRaeume");
            int nrOfDeletedRows = 0;
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                foreach (var raumRecord in raume)
                {
                    cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from {0} where RaumID={1} AND ProjektId={2}", "Raum", raumRecord.RaumId, projektId);
                    nrOfDeletedRows += cmd.ExecuteNonQuery();
                }
            }

            return nrOfDeletedRows;
        }

        private int DelWohnungen(int projektId, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Debug("DelWohnung");
            int nrOfDeletedRows = 0;
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "Delete from Wohnung where ProjektId={0}", projektId);
                nrOfDeletedRows += cmd.ExecuteNonQuery();
            }
            return nrOfDeletedRows;
        }

        private void InsertWohnung(IWohnungRecord wohnung, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "INSERT INTO Wohnung " + "([Top],[Typ],[Widmung],[Nutzwert],[ProjektId]) " + "VALUES(@Top,@Typ,@Widmung,@Nutzwert,@ProjektId)";

                // add named parameters
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Top", wohnung.Top??Convert.DBNull),
                               new OleDbParameter("@Typ", wohnung.Typ??Convert.DBNull),
                               new OleDbParameter("@Widmung", wohnung.Widmung??Convert.DBNull),
                               new OleDbParameter("@Nutzwert", wohnung.Nutzwert??Convert.DBNull),
                               new OleDbParameter("@ProjektId", wohnung.ProjektId),
                               });

                cmd.ExecuteNonQuery();
            }
        }

        private void InsertRaum(IRaumRecord raum, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "INSERT INTO Raum " + "([Top],[Lage],[Raum],[Widmung],[RNW],[Begrundung],[Nutzwert],[Flaeche],[ProjektId],[KategorieId],[AcadHandle]) "
                    + "VALUES(@Top,@Lage,@Raum, @Widmung, @RNW, @Begrundung, @Nutzwert,@Flaeche, @ProjektId, @KategorieId,@AcadHandle)";

                // add named parameters
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Top", raum.Top??Convert.DBNull),
                               new OleDbParameter("@Lage", raum.Lage??Convert.DBNull),
                               new OleDbParameter("@Raum", raum.Raum??Convert.DBNull),
                               new OleDbParameter("@Widmung", raum.Widmung??Convert.DBNull),
                               new OleDbParameter("@RNW", raum.RNW??Convert.DBNull),
                               new OleDbParameter("@Begrundung", raum.Begrundung??Convert.DBNull),
                               new OleDbParameter("@Nutzwert", raum.Nutzwert),
                               new OleDbParameter("@Flaeche", raum.Flaeche),
                               new OleDbParameter("@ProjektId", raum.ProjektId),
                               new OleDbParameter("@KategorieId", raum.KategorieId),
                               new OleDbParameter("@AcadHandle", raum.AcadHandle??Convert.DBNull),
                               });

                cmd.ExecuteNonQuery();
            }

        }

        private void InsertKategorieAndSetKategorieIdToKat(IKategorieRecord kat, OleDbConnection myConnection, OleDbTransaction transaction)
        {
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "INSERT INTO Kategorie " + "([Top],[Lage],[Widmung],[RNW],[Begrundung],[Nutzwert],[ProjektId]) " + "VALUES(@Top,@Lage, @Widmung, @RNW, @Begrundung, @Nutzwert, @ProjektId)";

                // add named parameters
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Top", kat.Top??Convert.DBNull),
                               new OleDbParameter("@Lage", kat.Lage??Convert.DBNull),
                               new OleDbParameter("@Widmung", kat.Widmung??Convert.DBNull),
                               new OleDbParameter("@RNW", kat.RNW??Convert.DBNull),
                               new OleDbParameter("@Begrundung", kat.Begrundung??Convert.DBNull),
                               new OleDbParameter("@Nutzwert", kat.Nutzwert),
                               new OleDbParameter("@ProjektId", kat.ProjektId),
                               });

                cmd.ExecuteNonQuery();
            }
            kat.KategorieID = GetKatId(kat, myConnection, transaction);
            if (kat.KategorieID == -1)
            {
                var msg = string.Format(
                    CultureInfo.CurrentCulture,
                    "Konnte KategorieId nicht ermitteln von Kategorie mit Top='{0}', Lage='{1}', Widmung='{2}', RNW='{3}', Begrundung='{4}', ProjektId='{5}'",
                    kat.Top, kat.Lage, kat.Widmung, kat.RNW, kat.Begrundung, kat.ProjektId
                    );
                log.Error(msg);
                throw new InvalidOperationException(msg);
            }
        }

        private int GetKatId(IKategorieRecord kat, OleDbConnection myConnection, OleDbTransaction transaction = null)
        {
            int katId = -1;
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                if (transaction != null) cmd.Transaction = transaction;
                cmd.CommandText = "select KategorieId from Kategorie where Top=@Top and Lage=@Lage and Widmung=@Widmung and RNW=@RNW and Begrundung=@Begrundung and ProjektId=@ProjektId";
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Top", kat.Top??Convert.DBNull),
                               new OleDbParameter("@Lage", kat.Lage??Convert.DBNull),
                               new OleDbParameter("@Widmung", kat.Widmung??Convert.DBNull),
                               new OleDbParameter("@RNW", kat.RNW??Convert.DBNull),
                               new OleDbParameter("@Begrundung", kat.Begrundung??Convert.DBNull),
                               new OleDbParameter("@ProjektId", kat.ProjektId),
                               });


                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    object o = reader["KategorieId"];
                    katId = (int)o;
                }
                reader.Close();
            }
            return katId;
        }

        private static void InsertProjekt(IProjektInfo projektInfo, OleDbConnection myConnection, OleDbTransaction transaction = null)
        {
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                if (transaction != null) cmd.Transaction = transaction;
                cmd.CommandText = "INSERT INTO Projekt " + "([Bauvorhaben], [DwgName],[EZ],[DwgPrefix],[Katastralgemeinde]) " + "VALUES(@Bauvorhaben, @DwgName, @EZ,@DwgPrefix,@Katastralgemeinde)";

                // add named parameters
                cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Bauvorhaben", projektInfo.Bauvorhaben??Convert.DBNull),
                               new OleDbParameter("@DwgName", projektInfo.DwgName??Convert.DBNull),
                               new OleDbParameter("@EZ", projektInfo.EZ??Convert.DBNull),
                               new OleDbParameter("@DwgPrefix", projektInfo.DwgPrefix??Convert.DBNull),
                               new OleDbParameter("@Katastralgemeinde", projektInfo.Katastralgemeinde??Convert.DBNull),
                               });

                cmd.ExecuteNonQuery();
            }
        }

        private static void InsertSubInfos(IProjektInfo projektInfo, OleDbConnection myConnection, OleDbTransaction transaction = null)
        {

            foreach (var subInfo in projektInfo.SubInfos)
            {
                using (var cmd = myConnection.CreateCommand())
                {
                    if (transaction != null) cmd.Transaction = transaction;
                    cmd.CommandText = "INSERT INTO GstInfo " + "([Gstnr], [Flaeche],[AcadHandle],[ProjektId]) " + "VALUES(@Gstnr, @Flaeche, @AcadHandle,@ProjektId)";

                    // add named parameters
                    cmd.Parameters.AddRange(new OleDbParameter[]
                               {
                               new OleDbParameter("@Gstnr", subInfo.Gstnr??Convert.DBNull),
                               new OleDbParameter("@Flaeche", subInfo.Flaeche??Convert.DBNull),
                               new OleDbParameter("@AcadHandle", subInfo.AcadHandle??Convert.DBNull),
                               new OleDbParameter("@ProjektId", projektInfo.ProjektId),
                               });

                    // todo: trennung von add parameter und parameter-value
                    // Beispiel:
                    //command.Parameters.Add("@Schlagwort", System.Data.SqlDbType.NVarChar, 256);
                    //command.Parameters.Add("@SwId", System.Data.SqlDbType.Int);

                    //command.Parameters["@Schlagwort"].Value = schlagWort.ValueString;
                    //command.Parameters["@SwId"].Value = schlagWort.SwId;

                    cmd.ExecuteNonQuery();
                }
            }
        }

        private static int GetProjektId(IProjektInfo projectInfo, OleDbConnection myConnection, OleDbTransaction transaction = null)
        {
            int projektId = -1;
            // Execute Queries
            using (var cmd = myConnection.CreateCommand())
            {
                if (transaction != null) cmd.Transaction = transaction;
                cmd.CommandText = string.Format(CultureInfo.InvariantCulture, "select ProjektId from {0} where Bauvorhaben='{1}'", "Projekt", projectInfo.Bauvorhaben);
                var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    object o = reader["ProjektId"];
                    projektId = (int)o;
                }
                reader.Close();
            }
            return projektId;
        }
        #endregion

        #region Override
        public override string ToString()
        {
            return MdbName;
        }
        #endregion

        #region KonsistenzPrüfung
        public bool CheckConsistency()
        {
            log.Info("CheckConsistency");

            var errors = false;

            try
            {
                using (OleDbConnection myConnection = new OleDbConnection())
                {
                    myConnection.ConnectionString = ConnectionString;
                    myConnection.Open();
                    var transaction = myConnection.BeginTransaction();

                    var raumMissingKat = ConsistRaumMissingKat(myConnection, transaction);
                    var katMissingRaum = ConsistKatMissingRaum(myConnection, transaction);
                    var zuAbschlagMissingKat = ConsistZuAbschlagMissingKat(myConnection, transaction);
                    var raumKatTop = ConsistRaumKatTop(myConnection, transaction);
                    var raumKatLage = ConsistRaumKatLage(myConnection, transaction);
                    var raumKatWidmung = ConsistRaumKatWidmung(myConnection, transaction);
                    var raumKatBegrundung = ConsistRaumKatBegrundung(myConnection, transaction);
                    var raumKatProjektId = ConsistRaumKatProjektId(myConnection, transaction);
                    var raumKatNutzwert = ConsistRaumKatNutzwert(myConnection, transaction);

                    if (raumMissingKat.Count > 0 ||
                        katMissingRaum.Count > 0 ||
                        zuAbschlagMissingKat.Count > 0 ||
                        raumKatTop.Count > 0 ||
                        raumKatLage.Count > 0 ||
                        raumKatWidmung.Count > 0 ||
                        raumKatBegrundung.Count > 0 ||
                        raumKatProjektId.Count > 0 ||
                        raumKatNutzwert.Count > 0
                    )
                        errors = true;

                    transaction.Commit();
                    myConnection.Close();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }

            if (errors) log.Warn("Konsistenzprüfung hat Fehler gefunden.");
            else log.Info("Konsistenzprüfung war erfolgreich.");
            return !errors;
        }

        private class RaumAndKat : IRaumRecord
        {

            public int RaumId { get; set; }
            public string AcadHandle { get; set; }
            public string Begrundung { get; set; }
            public double Flaeche { get; set; }
            public IKategorieRecord Kategorie { get; set; }
            public int KategorieId { get; set; }
            public string KatIdentification { get; set; }
            public string Lage { get; set; }
            public double Nutzwert { get; set; }
            public int ProjektId { get; set; }
            public string Raum { get; set; }
            public string RNW { get; set; }
            public string Top { get; set; }
            public string Widmung { get; set; }
            public void UpdateValuesFrom(IBlockInfo acadBlockInfo)
            {
                throw new NotImplementedException();
            }
            public IRaumRecord ShallowCopy()
            {
                return (IRaumRecord)MemberwiseClone();
            }

            public bool IsEqualTo(IRaumRecord otherRaumRecord)
            {
                throw new NotImplementedException();
            }
        }

        private List<IRaumRecord> ConsistRaumKatTop(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumKatTop");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                //cmd.CommandText =
                //    "SELECT Raum.*, Kategorie.KategorieID, Kategorie.[Top] FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Top] = Kategorie.[Top]";
                cmd.CommandText =
                    "SELECT Raum.*, Kategorie.* FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Top] = Kategorie.[Top]";
                ConsistExecuteCmdAndAddToRaumInfos(cmd, raumInfos, "Top");
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Unterschiedliche Top für Raum und Kategorie: Raum: {0}, Kategorie: {1}", rec.ToString(), (rec.Kategorie != null) ? rec.Kategorie.ToString() : ""));
            }
            return raumInfos;
        }

        private List<IRaumRecord> ConsistRaumKatLage(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumKatLage");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Raum.*, Kategorie.* FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Lage] = Kategorie.[Lage]";
                //cmd.CommandText =
                //    "SELECT Raum.*, Kategorie.KategorieID, Kategorie.[Lage] FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Lage] = Kategorie.[Lage]";
                ConsistExecuteCmdAndAddToRaumInfos(cmd, raumInfos, "Lage");
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Unterschiedliche Lage für Raum und Kategorie: Raum: {0}, Kategorie: {1}", rec.ToString(), (rec.Kategorie != null) ? rec.Kategorie.ToString() : ""));
            }
            return raumInfos;
        }

        private List<IRaumRecord> ConsistRaumKatWidmung(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumKatWidmung");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Raum.*, Kategorie.* FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Widmung] = Kategorie.[Widmung]";
                ConsistExecuteCmdAndAddToRaumInfos(cmd, raumInfos, "Widmung");
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Unterschiedliche Widmung für Raum und Kategorie: Raum: {0}, Kategorie: {1}", rec.ToString(), (rec.Kategorie != null) ? rec.Kategorie.ToString() : ""));
            }
            return raumInfos;
        }

        private List<IRaumRecord> ConsistRaumKatBegrundung(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumKatBegrundung");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Raum.*, Kategorie.* FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Begrundung] = Kategorie.[Begrundung]";
                //cmd.CommandText =
                //    "SELECT Raum.*, Kategorie.KategorieID, Kategorie.[Begrundung] FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[Begrundung] = Kategorie.[Begrundung]";
                ConsistExecuteCmdAndAddToRaumInfos(cmd, raumInfos, "Begrundung");
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Unterschiedliche Begründung für Raum und Kategorie: Raum: {0}, Kategorie: {1}", rec.ToString(), (rec.Kategorie != null) ? rec.Kategorie.ToString() : ""));
            }
            return raumInfos;
        }

        private List<IRaumRecord> ConsistRaumKatProjektId(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumKatProjektId");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Raum.*, Kategorie.* FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[ProjektID] = Kategorie.[ProjektID]";
                //cmd.CommandText =
                //    "SELECT Raum.*, Kategorie.KategorieID, Kategorie.[ProjektID] FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE NOT Raum.[ProjektID] = Kategorie.[ProjektID]";
                ConsistExecuteCmdAndAddToRaumInfos(cmd, raumInfos, "ProjektId");
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Unterschiedliche ProjektId für Raum und Kategorie: Raum: {0}, Kategorie: {1}", rec.ToString(), (rec.Kategorie != null) ? rec.Kategorie.ToString() : ""));
            }
            return raumInfos;
        }

        private List<IRaumRecord> ConsistRaumKatNutzwert(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumKatNutzwert");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Raum.*, Kategorie.* FROM Raum INNER JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE abs(Raum.Nutzwert-Kategorie.Nutzwert)>0.0001";
                ConsistExecuteCmdAndAddToRaumInfos(cmd, raumInfos, "Nutzwert");
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Unterschiedlicher Nutzwert für Raum und Kategorie: Raum: {0}, Kategorie: {1}", rec.ToString(), (rec.Kategorie != null) ? rec.Kategorie.ToString() : ""));
            }
            return raumInfos;
        }

        private void ConsistExecuteCmdAndAddToRaumInfos(OleDbCommand cmd, List<IRaumRecord> raumInfos, string fieldName)
        {
            var reader = cmd.ExecuteReader();

            if (reader != null)
            {
                var schemaTable = reader.GetSchemaTable();


                var raumNameDict = new Dictionary<string, string>();
                // ReSharper disable once PossibleNullReferenceException
                foreach (DataRow myField in schemaTable.Rows)
                {
                    var colName = myField["ColumnName"].ToString();
                    if (!colName.StartsWith("Kategorie", StringComparison.OrdinalIgnoreCase))
                    {
                        var arr = colName.Split(new char[] { '.' });
                        string getName = string.Empty;
                        if (arr.Length == 2) getName = arr[1];
                        else getName = colName;

                        raumNameDict.Add(getName, colName);
                    }
                }

                while (reader.Read())
                {
                    var raum = _Factory.CreateRaumRecord();

                    var o = reader[raumNameDict["ProjektID"]];
                    if (o != System.DBNull.Value) raum.ProjektId = (int)o;

                    o = reader[raumNameDict["KategorieID"]];
                    if (o != System.DBNull.Value) raum.KategorieId = (int)o;

                    o = reader[raumNameDict["RaumID"]];
                    if (o != System.DBNull.Value) raum.RaumId = (int)o;

                    o = reader[raumNameDict["Top"]];
                    if (o != System.DBNull.Value) raum.Top = o.ToString();

                    o = reader[raumNameDict["Lage"]];
                    if (o != System.DBNull.Value) raum.Lage = o.ToString();

                    o = reader[raumNameDict["Raum"]];
                    if (o != System.DBNull.Value) raum.Raum = o.ToString();

                    o = reader[raumNameDict["Widmung"]];
                    if (o != System.DBNull.Value) raum.Widmung = o.ToString();

                    o = reader[raumNameDict["RNW"]];
                    if (o != System.DBNull.Value) raum.RNW = o.ToString();

                    o = reader[raumNameDict["Begrundung"]];
                    if (o != System.DBNull.Value) raum.Begrundung = o.ToString();

                    o = reader[raumNameDict["Nutzwert"]];
                    if (o != System.DBNull.Value) raum.Nutzwert = Convert.ToDouble(o);

                    o = reader[raumNameDict["Flaeche"]];
                    if (o != System.DBNull.Value) raum.Flaeche = Convert.ToDouble(o);

                    o = reader[raumNameDict["AcadHandle"]];
                    if (o != System.DBNull.Value) raum.AcadHandle = o.ToString();

                    var kat = _Factory.CreateKategorie();
                    raum.Kategorie = kat;
                    o = reader["Kategorie.KategorieID"];
                    if (o != System.DBNull.Value) kat.KategorieID = (int)o;

                    o = reader["Kategorie.Top"];
                    if (o != System.DBNull.Value) kat.Top = o.ToString();

                    o = reader["Kategorie.Lage"];
                    if (o != System.DBNull.Value) kat.Lage = o.ToString();

                    o = reader["Kategorie.Widmung"];
                    if (o != System.DBNull.Value) kat.Widmung = o.ToString();

                    o = reader["Kategorie.Begrundung"];
                    if (o != System.DBNull.Value) kat.Begrundung = o.ToString();

                    o = reader["Kategorie.Nutzwert"];
                    if (o != System.DBNull.Value) kat.Nutzwert = Convert.ToDouble(o);

                    o = reader["Kategorie.ProjektID"];
                    if (o != System.DBNull.Value) kat.ProjektId = (int)o;

                    raumInfos.Add(raum);
                }
            }
            reader.Close();
        }

        private void ConsistExecuteCmdAndAddToRaumInfos(OleDbCommand cmd, List<IRaumRecord> raumInfos, string fieldName, string abc)
        {
            var reader = cmd.ExecuteReader();
            if (reader != null)
            {
                while (reader.Read())
                {
                    var raum = _Factory.CreateRaumRecord();
                    var o = reader[ConsistGetReadName("Raum", "ProjektID", fieldName)];
                    if (o != System.DBNull.Value) raum.ProjektId = (int)o;

                    o = reader["Raum.KategorieID"];
                    if (o != System.DBNull.Value) raum.KategorieId = (int)o;

                    o = reader["RaumID"];
                    if (o != System.DBNull.Value) raum.RaumId = (int)o;

                    o = reader[ConsistGetReadName("Raum", "Top", fieldName)];
                    if (o != System.DBNull.Value) raum.Top = o.ToString();

                    o = reader[ConsistGetReadName("Raum", "Lage", fieldName)];
                    if (o != System.DBNull.Value) raum.Lage = o.ToString();

                    o = reader["Raum"];
                    if (o != System.DBNull.Value) raum.Raum = o.ToString();

                    o = reader[ConsistGetReadName("Raum", "Widmung", fieldName)];
                    if (o != System.DBNull.Value) raum.Widmung = o.ToString();

                    o = reader[ConsistGetReadName("Raum", "RNW", fieldName)];
                    if (o != System.DBNull.Value) raum.RNW = o.ToString();

                    o = reader[ConsistGetReadName("Raum", "Begrundung", fieldName)];
                    if (o != System.DBNull.Value) raum.Begrundung = o.ToString();

                    o = reader[ConsistGetReadName("Raum", "Nutzwert", fieldName)];
                    if (o != System.DBNull.Value) raum.Nutzwert = Convert.ToDouble(o);

                    o = reader[ConsistGetReadName("Raum", "Flaeche", fieldName)];
                    if (o != System.DBNull.Value) raum.Flaeche = Convert.ToDouble(o);

                    o = reader["AcadHandle"];
                    if (o != System.DBNull.Value) raum.AcadHandle = o.ToString();

                    var kat = _Factory.CreateKategorie();
                    raum.Kategorie = kat;
                    o = reader["Kategorie.KategorieID"];
                    if (o != System.DBNull.Value) kat.KategorieID = (int)o;

                    if (string.Compare(fieldName, "Top", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        o = reader[ConsistGetReadName("Kategorie", "Top", fieldName)];
                        if (o != System.DBNull.Value) kat.Top = o.ToString();
                    }
                    if (string.Compare(fieldName, "Lage", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        o = reader[ConsistGetReadName("Kategorie", "Lage", fieldName)];
                        if (o != System.DBNull.Value) kat.Lage = o.ToString();
                    }
                    if (string.Compare(fieldName, "Widmung", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        o = reader[ConsistGetReadName("Kategorie", "Widmung", fieldName)];
                        if (o != System.DBNull.Value) kat.Widmung = o.ToString();
                    }
                    if (string.Compare(fieldName, "Begrundung", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        o = reader[ConsistGetReadName("Kategorie", "Begrundung", fieldName)];
                        if (o != System.DBNull.Value) kat.Begrundung = o.ToString();
                    }
                    if (string.Compare(fieldName, "Nutzwert", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        o = reader[ConsistGetReadName("Kategorie", "Nutzwert", fieldName)];
                        if (o != System.DBNull.Value) kat.Nutzwert = Convert.ToDouble(o);
                    }
                    if (string.Compare(fieldName, "ProjektID", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        o = reader[ConsistGetReadName("Kategorie", "ProjektID", fieldName)];
                        if (o != System.DBNull.Value) kat.ProjektId = (int)o;
                    }
                    raumInfos.Add(raum);
                }
            }
            reader.Close();
        }

        private string ConsistGetReadName(string prefix, string name, string fieldName)
        {
            if (string.Compare(name, fieldName, StringComparison.OrdinalIgnoreCase) == 0)
            {
                return prefix + "." + name;
            }
            return name;
        }

        private List<IKategorieRecord> ConsistKatMissingRaum(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistKatMissingRaum");
            var katInfos = new List<IKategorieRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Kategorie.* FROM Raum RIGHT JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE Raum.KategorieID Is Null";
                var reader = cmd.ExecuteReader();
                if (reader != null)
                {
                    while (reader.Read())
                    {
                        var kat = _Factory.CreateKategorie();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) kat.ProjektId = (int)o;

                        o = reader["KategorieID"];
                        if (o != System.DBNull.Value) kat.KategorieID = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) kat.Top = o.ToString();

                        o = reader["Lage"];
                        if (o != System.DBNull.Value) kat.Lage = o.ToString();

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) kat.Widmung = o.ToString();

                        o = reader["RNW"];
                        if (o != System.DBNull.Value) kat.RNW = o.ToString();

                        o = reader["Begrundung"];
                        if (o != System.DBNull.Value) kat.Begrundung = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) kat.Nutzwert = Convert.ToDouble(o);

                        katInfos.Add(kat);
                    }
                }
                reader.Close();
            }
            foreach (var rec in katInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Kategorie ohne Raum: {0}", rec.ToString()));
            }
            return katInfos;
        }
        private List<int> ConsistZuAbschlagMissingKat(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistZuAbschlagMissingKat");
            var zuAbschlagInfos = new List<int>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT ZuAbschlag.* FROM ZuAbschlag LEFT JOIN Kategorie ON ZuAbschlag.KategorieId = Kategorie.KategorieID WHERE Kategorie.KategorieID Is Null";
                var reader = cmd.ExecuteReader();
                if (reader != null)
                {
                    while (reader.Read())
                    {
                        object o = reader["ZuAbschlagId"];
                        if (o != System.DBNull.Value)
                        {
                            zuAbschlagInfos.Add((int)o);
                        }
                    }
                }
                reader.Close();
            }
            foreach (var rec in zuAbschlagInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: ZuAbschlag ohne Kategorie: Id={0}", rec.ToString()));
            }
            return zuAbschlagInfos;
        }

        private List<IRaumRecord> ConsistRaumMissingKat(OleDbConnection myConnection, OleDbTransaction transaction)
        {
            log.Info("ConsistRaumMissingKat");
            var raumInfos = new List<IRaumRecord>();
            using (var cmd = myConnection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "SELECT Raum.* FROM Raum LEFT JOIN Kategorie ON Raum.KategorieID = Kategorie.KategorieID WHERE Kategorie.KategorieID Is Null";
                var reader = cmd.ExecuteReader();
                if (reader != null)
                {
                    while (reader.Read())
                    {
                        var raum = _Factory.CreateRaumRecord();
                        object o = reader["ProjektID"];
                        if (o != System.DBNull.Value) raum.ProjektId = (int)o;

                        o = reader["KategorieId"];
                        if (o != System.DBNull.Value) raum.KategorieId = (int)o;

                        o = reader["RaumId"];
                        if (o != System.DBNull.Value) raum.RaumId = (int)o;

                        o = reader["Top"];
                        if (o != System.DBNull.Value) raum.Top = o.ToString();

                        o = reader["Lage"];
                        if (o != System.DBNull.Value) raum.Lage = o.ToString();

                        o = reader["Raum"];
                        if (o != System.DBNull.Value) raum.Raum = o.ToString();

                        o = reader["Widmung"];
                        if (o != System.DBNull.Value) raum.Widmung = o.ToString();

                        o = reader["RNW"];
                        if (o != System.DBNull.Value) raum.RNW = o.ToString();

                        o = reader["Begrundung"];
                        if (o != System.DBNull.Value) raum.Begrundung = o.ToString();

                        o = reader["Nutzwert"];
                        if (o != System.DBNull.Value) raum.Nutzwert = Convert.ToDouble(o);

                        o = reader["Flaeche"];
                        if (o != System.DBNull.Value) raum.Flaeche = Convert.ToDouble(o);

                        o = reader["AcadHandle"];
                        if (o != System.DBNull.Value) raum.AcadHandle = o.ToString();
                        raumInfos.Add(raum);
                    }
                }
                reader.Close();
            }
            foreach (var rec in raumInfos)
            {
                log.Warn(string.Format(CultureInfo.CurrentCulture, "Konsistenz: Raum ohne Kategorie: {0}", rec.ToString()));
            }
            return raumInfos;
        }

        #endregion
    }
}
