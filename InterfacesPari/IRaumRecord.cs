using System;
namespace InterfacesPari
{
    public interface IRaumRecord
    {
        int RaumId { get; set; }
        string AcadHandle { get; set; }
        string Begrundung { get; set; }
        double Flaeche { get; set; }
        IKategorieRecord Kategorie { get; set; }
        int KategorieId { get; set; }
        string KatIdentification { get; }
        string Lage { get; set; }
        double Nutzwert { get; set; }
        int ProjektId { get; set; }
        string Raum { get; set; }
        string RNW { get; set; }
        string Top { get; set; }
        string Widmung { get; set; }
        void UpdateValuesFrom(IBlockInfo acadBlockInfo);
        IRaumRecord ShallowCopy();
        bool IsEqualTo(IRaumRecord otherRaumRecord);
    }
}
