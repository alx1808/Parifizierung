using InterfacesPari;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UiPari.ViewModel
{
    public class KategorieRecord : IKategorieRecord, INotifyPropertyChanged
    {
        public KategorieRecord() { }
        public KategorieRecord(IKategorieRecord kat)
        {
            this.Top = kat.Top;
            this.Lage = kat.Lage;
            this.Widmung = kat.Widmung;
            this.RNW = kat.RNW;
            this.Begrundung = kat.Begrundung;
            this.Nutzwert = kat.Nutzwert;
            this.KategorieID = kat.KategorieID;
            this.ProjektId = kat.ProjektId;
        }
        private string _Top = string.Empty;
        public string Top
        {
            get
            {
                return _Top;
            }
            set
            {
                if (_Top != value)
                {
                    _Top = value;
                    OnPropertyChanged("Top");
                }
            }
        }

        private string _Lage = string.Empty;
        public string Lage
        {
            get
            {
                return _Lage;
            }
            set
            {
                if (_Lage != value)
                {
                    _Lage = value;
                    OnPropertyChanged("Lage");
                }
            }
        }

        private string _Widmung = string.Empty;
        public string Widmung
        {
            get
            {
                return _Widmung;
            }
            set
            {
                if (_Widmung != value)
                {
                    _Widmung = value;
                    OnPropertyChanged("Widmung");
                }
            }
        }

        public string Identification { get { return string.Join("|", new string[] { Top, Lage, Widmung, Begrundung }); } }

        private string _RNW = string.Empty;
        public string RNW
        {
            get
            {
                return _RNW;
            }
            set
            {
                if (_RNW != value)
                {
                    _RNW = value;
                    OnPropertyChanged("RNW");
                }
            }
        }

        private string _Begrundung = string.Empty;
        public string Begrundung
        {
            get
            {
                return _Begrundung;
            }
            set
            {
                if (_Begrundung != value)
                {
                    _Begrundung = value;
                    OnPropertyChanged("Begrundung");
                }
            }
        }

        private double _Nutzwert;
        public double Nutzwert
        {
            get
            {
                return _Nutzwert;
            }
            set
            {
                _Nutzwert = value;
                OnPropertyChanged("Nutzwert");
            }
        }

        private int _KategorieID;
        public int KategorieID
        {
            get { return _KategorieID; }
            set { _KategorieID = value;
            OnPropertyChanged("KategorieID");
            }
        }

        private int _ProjektId;
        public int ProjektId
        {
            get { return _ProjektId; }
            set { _ProjektId = value;
            OnPropertyChanged("ProjektId");
            }
        }

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
