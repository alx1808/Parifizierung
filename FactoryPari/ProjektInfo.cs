using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FactoryPari
{
    public class ProjektInfo : IProjektInfo
    {
        //private const string BLOCK_NAME = "Grundstücksinfo";

        #region From Dwg
        private string _DwgName = string.Empty;
        public string DwgName
        {
            get { return _DwgName; }
            set { _DwgName = value; }
        }

        private string _DwgPrefix = string.Empty;
        public string DwgPrefix
        {
            get { return _DwgPrefix; }
            set { _DwgPrefix = value; }
        }
        #endregion

        #region From Database
        private int _ProjektId = -1;
        public int ProjektId
        {
            get { return _ProjektId; }
            set { _ProjektId = value; }
        }
        #endregion

        #region Properties von Grundstücksinfo.dwg
        public string Bauvorhaben { get; set; }

        public string Katastralgemeinde { get; set; }
        
        public string EZ {get;set;}

        public class SubInfo : ISubInfo
        {
            public string Gstnr { get; set; }
            public string Flaeche { get; set; }
            public string AcadHandle { get; set; }
        }

        private List<ISubInfo> _SubInfos = new List<ISubInfo>();
        public List<ISubInfo> SubInfos
        {
            get { return _SubInfos; }
            set { _SubInfos = value; }
        }


        #endregion

    }
}
