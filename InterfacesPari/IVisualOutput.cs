using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfacesPari
{
    public interface IVisualOutput
    {
        void ExportNW(IPariDatabase database, string locationHint, int projektId);
        void ExportNF(IPariDatabase database, string locationHint, int projektId);
        void SetTemplates(object oTemplate);
    }
}
