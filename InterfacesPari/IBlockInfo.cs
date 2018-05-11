using System;
namespace InterfacesPari
{
    public interface IBlockInfo
    {
        string Begrundung { get; set; }
        string Flaeche { get; set; }
        string Geschoss { get; set; }
        string Handle { get; set; }
        string Nutzwert { get; set; }
        string Raum { get; set; }
        string Top { get; set; }
        string Zusatz { get; set; }
        string Widmung { get; set; }
    }
}
