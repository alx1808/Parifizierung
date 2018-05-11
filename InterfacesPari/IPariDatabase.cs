using System;
using System.Collections.Generic;
namespace InterfacesPari
{
    public interface IPariDatabase
    {
        int DeleteProjekt(int projektId);
        int GetProjektId(IProjektInfo projektInfo);
        List<string> GetTableNames();
        List<IProjektInfo> ListProjInfos();
        void SaveToDatabase(ITableBuilder tableBuilder, IProjektInfo projektInfo);
        void UpdateDatabase(ITableUpdater tableUpdater, IProjektInfo projektInfo);
        List<IKategorieRecord> GetKategories(int projektId);
        List<IZuAbschlagRecord> GetZuAbschlags(int kategorieId);
        List<IKategorieZaRecord> GetKategoriesWithZuAbschlag(int projektId);
        List<IWohnungRecord> GetWohnungen(int projektId);
        List<IZuAbschlagVorgabeRecord> GetZuAbschlagVorgaben();
        //List<IZuAbschlagRecord> GetZuAbschlagsForProjekt(int projektId);
        List<IRaumRecord> GetRaeume(int projektId);
        List<IRaumZaRecord> GetRaeumeWithZuAbschlag(int projektId);
        int UpdateKategorie(IKategorieRecord kategorie);
        int InsertZuAbschlag(IZuAbschlagRecord zuAbschlag);
        int UpdateZuAbschlag(IZuAbschlagRecord zuAbschlagRec);
        int DeleteZuAbschlag(IZuAbschlagRecord zuAbschlagRec);
        void SetDatabase(object dbo);
        bool CheckConsistency();
        int CheckExistingFields();
    }
}
