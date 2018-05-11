using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UiPari.ViewModel
{
    internal class ZuAbschlagRecord : IZuAbschlagRecord, INotifyPropertyChanged
    {
        public ZuAbschlagRecord() { }
        public ZuAbschlagRecord(IZuAbschlagRecord zaRec)
        {
            this.Beschreibung = zaRec.Beschreibung;
            this.Prozent = zaRec.Prozent;
            this.KategorieId = zaRec.KategorieId;
            this.ProjektId = zaRec.ProjektId;
            this.ZuAbschlagId = zaRec.ZuAbschlagId;
        }

        private string _Beschreibung;
        public string Beschreibung
        {
            get { return _Beschreibung; }
            set
            {
                _Beschreibung = value;
                OnPropertyChanged("Beschreibung");
            }
        }

        private double _Prozent;
        public double Prozent
        {
            get { return _Prozent; }
            set
            {
                _Prozent = value;
                OnPropertyChanged("Prozent");
            }
        }

        private int _KategorieId;
        public int KategorieId
        {
            get { return _KategorieId; }
            set
            {
                _KategorieId = value;
                OnPropertyChanged("KategorieId");
            }
        }

        private int _ProjektId;
        public int ProjektId
        {
            get { return _ProjektId; }
            set
            {
                _ProjektId = value;
                OnPropertyChanged("ProjektId");
            }
        }

        private int _ZuAbschlagId;
        public int ZuAbschlagId
        {
            get { return _ZuAbschlagId; }
            set
            {
                _ZuAbschlagId = value;
                OnPropertyChanged("ZuAbschlagId");
            }
        }

        private IZuAbschlagVorgabeRecord _Zav = null;
        public IZuAbschlagVorgabeRecord Zav
        {
            get { return _Zav; }
            set { _Zav = value;
            Beschreibung = _Zav.Beschreibung;
            Prozent = _Zav.Prozent;
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
