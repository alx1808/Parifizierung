using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UiPari.ViewModel
{
    internal class ProjektInfoContainer : INotifyPropertyChanged
    {
        public ProjektInfoContainer() { }

        public ProjektInfoContainer(List<InterfacesPari.IProjektInfo> projInfos)
        {
            if (projInfos == null || projInfos.Count == 0) return;
            foreach (var pi in projInfos)
            {
                ProjektInfos.Add(new ProjektInfo(pi));
            }
        }

        private ProjektInfo _TheProjektInfo = null;
        public ProjektInfo TheProjektInfo
        {
            get { return _TheProjektInfo; }
            set { _TheProjektInfo = value;
            OnPropertyChanged("TheProjektInfo");
            }
        }

        private ObservableCollection<ProjektInfo> _ProjektInfos = new ObservableCollection<ProjektInfo>();
        public ObservableCollection<ProjektInfo> ProjektInfos
        {
            get { return _ProjektInfos; }
            set { _ProjektInfos = value; }
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
