using System;
namespace InterfacesPari
{
    public interface IFactory
    {
        IPariDatabase CreatePariDatabase();
        IProjektInfo CreateProjectInfo();
        ISubInfo CreateSubInfo();
        IKategorieRecord CreateKategorie();
        IKategorieZaRecord CreateZaKategorie();
        IKategorieRecord CreateKategorie(IRaumRecord raumRecord);
        IVisualOutput CreateVisualOutputHandler();
        IRaumRecord CreateRaumRecord();
        IRaumRecord CreateRaumRecord(IBlockInfo acadBlockInfo);
    }
}
