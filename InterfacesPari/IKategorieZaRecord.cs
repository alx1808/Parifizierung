using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface IKategorieZaRecord : IKategorieRecord
    {
        List<IZuAbschlagRecord> ZuAbschlaege { get; set; }
        double SumProzent { get; }
        double ActualNutzwert { get; }
    }
}
