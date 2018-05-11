using InterfacesPari;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UiPari.ViewModel
{
    internal class ProjektInfo :  INotifyPropertyChanged
    {
        public ProjektInfo() { }
        public ProjektInfo(IProjektInfo projektInfo)
        {
            this.Bauvorhaben = projektInfo.Bauvorhaben;
            this.ProjektId = projektInfo.ProjektId;
        }

        private string _Bauvorhaben = string.Empty;
        public string Bauvorhaben
        {
            get
            {
                return _Bauvorhaben;
            }
            set
            {
                if (_Bauvorhaben != value)
                {
                    _Bauvorhaben = value;
                    OnPropertyChanged("Bauvorhaben");
                }
            }
        }

        public int ProjektId { get; set;}

        public ProjektInfo TheProjekt { get { return this; } set { int i = 1; } }

        //public string DwgName
        //{
        //    get
        //    {
        //        throw new NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new NotImplementedException();
        //    }
        //}

        //public string DwgPrefix
        //{
        //    get
        //    {
        //        throw new NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new NotImplementedException();
        //    }
        //}

        //public string EZ
        //{
        //    get
        //    {
        //        throw new NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new NotImplementedException();
        //    }
        //}

        //public string Katastralgemeinde
        //{
        //    get
        //    {
        //        throw new NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new NotImplementedException();
        //    }
        //}

        //public List<ISubInfo> SubInfos
        //{
        //    get
        //    {
        //        throw new NotImplementedException();
        //    }
        //    set
        //    {
        //        throw new NotImplementedException();
        //    }
        //}
        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

    }
}
