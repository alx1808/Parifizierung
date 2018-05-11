using System;
using System.Collections.Generic;
namespace InterfacesPari
{
    public interface IProjektInfo
    {
        string Bauvorhaben { get; set; }
        string DwgName { get; set; }
        string DwgPrefix { get; set; }
        string EZ { get; set; }
        string Katastralgemeinde { get; set; }
        int ProjektId { get; set; }
        List<ISubInfo> SubInfos { get; set; }
    }
}
