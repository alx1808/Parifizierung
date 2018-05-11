using FactoryPari;
using InterfacesPari;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using UiPari.ViewModel;

namespace UiPari
{
    /// <summary>
    /// Interaction logic for ChangeDataUc.xaml
    /// </summary>
    public partial class ChangeDataUc : UserControl, INotifyPropertyChanged
    {
        #region log4net Initialization
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(ChangeDataUc));
        static ChangeDataUc()
        {
            if (log4net.LogManager.GetRepository(System.Reflection.Assembly.GetExecutingAssembly()).Configured == false)
            {
                log4net.Config.XmlConfigurator.ConfigureAndWatch(
                    new System.IO.FileInfo(
                        System.IO.Path.Combine(
                            new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName,
                            "_log4net.config"
                        )
                    )
                );
            }
        }
        #endregion

        private ProjektInfoContainer _ProjektInfoContainer = null;
        private Factory _Factory;
        private IPariDatabase _Database;
        private ObservableCollection<ViewModel.KategorieRecord> _Kategories = new ObservableCollection<ViewModel.KategorieRecord>();
        private ObservableCollection<ViewModel.ZuAbschlagRecord> _ZuAbschlags = new ObservableCollection<ViewModel.ZuAbschlagRecord>();

        private ObservableCollection<ZuAbschlagVorgabe> _ZAV = new ObservableCollection<ZuAbschlagVorgabe>();
        public ObservableCollection<ZuAbschlagVorgabe> ZAV
        {
            get { return _ZAV; }
            set { _ZAV = value; }
        }
        private ZuAbschlagVorgabe _TheZav = null;
        public ZuAbschlagVorgabe TheZav
        {
            get { return _TheZav; }
            set
            {
                _TheZav = value;
                var si = dgKategorien.SelectedItem;
                if (si != null)
                {
                    if (!string.IsNullOrEmpty(_TheZav.Beschreibung))
                    {
                        _ZuAbschlags.Add(new ZuAbschlagRecord() { Beschreibung = _TheZav.Beschreibung, Prozent = _TheZav.Prozent });
                    }
                    OnPropertyChanged("TheZav");
                    OnPropertyChanged("ZAV");
                }
            }
        }


        public ChangeDataUc()
        {
            InitializeComponent();

            log.Debug("New ChangeDataUc Window.");

            try
            {
                _Factory = new Factory();
                _Database = _Factory.CreatePariDatabase();

                var zav = _Database.GetZuAbschlagVorgaben();
                _ZAV.Add(new ZuAbschlagVorgabe() { Beschreibung = "", Prozent = 0.0 });
                foreach (var z in zav)
                {
                    _ZAV.Add(new ZuAbschlagVorgabe() { Beschreibung = z.Beschreibung, Prozent = z.Prozent });
                }
                cmbZaVorgaben.DataContext = this;
                TheZav = _ZAV[0];

                var projInfos = _Database.ListProjInfos();
                _ProjektInfoContainer = new ProjektInfoContainer(projInfos);
                ProjektCombo.DataContext = _ProjektInfoContainer;
                _ProjektInfoContainer.PropertyChanged += ProjektInfoContainer_PropertyChanged;

                dgKategorien.ItemsSource = _Kategories;
                dgKategorien.RowEditEnding += dgKategorien_RowEditEnding;
                dgKategorien.SelectionChanged += dgKategorien_SelectionChanged;

                dgZuAbschlag.ItemsSource = _ZuAbschlags;
                dgZuAbschlag.RowEditEnding += dgZuAbschlag_RowEditEnding;
                dgZuAbschlag.CellEditEnding += dgZuAbschlag_CellEditEnding;
                _ZuAbschlags.CollectionChanged += _ZuAbschlags_CollectionChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
                log.Error(ex.Message, ex);
            }
        }

        void dgZuAbschlag_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }


        void dgZuAbschlag_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            try
            {
                if (dgKategorien.SelectedItem == null) return;
                var row = e.Row;
                var zaRec = row.Item as ViewModel.ZuAbschlagRecord;
                UpdateZaRec(zaRec);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
            }
        }

        private void UpdateZaRec(IZuAbschlagRecord zaRec)
        {
            var nr = _Database.UpdateZuAbschlag((IZuAbschlagRecord)zaRec);
            if (nr <= 0)
            {
                var msg = string.Format(CultureInfo.CurrentCulture, "Der Zu-Abschlag {0} konnte nicht gespeichert werden!", zaRec.KategorieId);
                log.Error(msg);
                throw new InvalidOperationException(msg);
            }
            else if (nr > 1)
            {
                var msg = string.Format(CultureInfo.CurrentCulture, "Der Zu-Abschlag {0} kommt mehrmals vor: {1}", zaRec.ZuAbschlagId, nr);
                log.Error(msg);
                throw new InvalidOperationException(msg);
            }
        }

        void _ZuAbschlags_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            try
            {
                switch (e.Action)
                {
                    case System.Collections.Specialized.NotifyCollectionChangedAction.Add:
                        if (IgnoreAddZuabschlagAdd) return;
                        var newItems = e.NewItems;
                        var zaRecToAdd = (IZuAbschlagRecord)newItems[0];
                        var kat = (IKategorieRecord)dgKategorien.SelectedItem;
                        if (kat == null) return;
                        zaRecToAdd.KategorieId = kat.KategorieID;
                        zaRecToAdd.ProjektId = kat.ProjektId;
                        _Database.InsertZuAbschlag(zaRecToAdd);
                        break;
                    case System.Collections.Specialized.NotifyCollectionChangedAction.Move:
                        break;
                    case System.Collections.Specialized.NotifyCollectionChangedAction.Remove:
                        var oldItems = e.OldItems;
                        var delZaRec = (IZuAbschlagRecord)oldItems[0];
                        if (_Database.DeleteZuAbschlag(delZaRec) <= 0)
                        {
                            var msg = string.Format(CultureInfo.CurrentCulture, "Konnte Record id={0} in Zu/Abschlag-Tabelle nicht löschen!", delZaRec.ZuAbschlagId);
                            log.Error(msg);
                        }
                        break;
                    case System.Collections.Specialized.NotifyCollectionChangedAction.Replace:
                        break;
                    case System.Collections.Specialized.NotifyCollectionChangedAction.Reset:
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
            }
        }

        private bool IgnoreAddZuabschlagAdd = false;
        void dgKategorien_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                TheZav = _ZAV[0];

                var index = dgKategorien.SelectedIndex;
                var kat = (IKategorieRecord)dgKategorien.SelectedItem;
                if (kat == null) return;

                _ZuAbschlags.Clear();
                var zaInfos = _Database.GetZuAbschlags(kat.KategorieID);
                try
                {
                    IgnoreAddZuabschlagAdd = true;
                    foreach (var za in zaInfos)
                    {
                        _ZuAbschlags.Add(new ZuAbschlagRecord(za));
                    }
                }
                finally
                {
                    IgnoreAddZuabschlagAdd = false;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
            }
        }

        void dgKategorien_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            try
            {
                var row = e.Row;
                var kat = row.Item as ViewModel.KategorieRecord;
                var nr = _Database.UpdateKategorie(kat);
                if (nr <= 0)
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "Die Kategorie {0} konnte nicht gespeichert werden!", kat.KategorieID);
                    log.Error(msg);
                    throw new InvalidOperationException(msg);
                }
                else if (nr > 1)
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "Die Kombination KategorieId={0}, ProjektId={1} kommt mehrmals vor: {2}", kat.KategorieID, kat.ProjektId, nr);
                    log.Error(msg);
                    throw new InvalidOperationException(msg);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
            }
        }

        void ProjektInfoContainer_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "TheProjektInfo")
            {
                _Kategories.Clear();
                _ZuAbschlags.Clear();
                var kats = _Database.GetKategories(_ProjektInfoContainer.TheProjektInfo.ProjektId);
                foreach (var kat in kats)
                {
                    _Kategories.Add(new ViewModel.KategorieRecord(kat));
                }
            }
        }

        private void btnExportNW_Click(object sender, RoutedEventArgs e)
        {
            Cursor defCursor = Cursor;
            try
            {
                Cursor = Cursors.Wait;
                if (_ProjektInfoContainer == null || _ProjektInfoContainer.TheProjektInfo == null || _ProjektInfoContainer.TheProjektInfo.ProjektId < 0) return;
                var exporter = _Factory.CreateVisualOutputHandler();
                exporter.ExportNW(_Database, null, _ProjektInfoContainer.TheProjektInfo.ProjektId);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
            }
            finally
            {
                Cursor = defCursor;
            }
        }

        private void btnExportNF_Click(object sender, RoutedEventArgs e)
        {
            Cursor defCursor = Cursor;
            try
            {
                Cursor = Cursors.Wait;
                if (_ProjektInfoContainer == null || _ProjektInfoContainer.TheProjektInfo == null || _ProjektInfoContainer.TheProjektInfo.ProjektId < 0) return;
                var exporter = _Factory.CreateVisualOutputHandler();
                exporter.ExportNF(_Database, null, _ProjektInfoContainer.TheProjektInfo.ProjektId);
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
            }
            finally
            {
                Cursor = defCursor;
            }
        }

        private void btnDbLocation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_Database == null) throw new InvalidOperationException("Database is null!");

                Microsoft.Win32.OpenFileDialog fd = new Microsoft.Win32.OpenFileDialog();
                fd.Title = "Datenbank wählen";
                fd.Filter = "Access Datenbank|*.accdb";
                // todo: this
                //bool? res = fd.ShowDialog(this);
                bool? res = fd.ShowDialog();
                if (res.Value == true)
                {
                    _Database.SetDatabase(fd.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
                log.Error(ex.Message, ex);
            }
        }

        private void btnExcelTemplateLocation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_Database == null) throw new InvalidOperationException("Database is null!");

                Microsoft.Win32.OpenFileDialog fd = new Microsoft.Win32.OpenFileDialog();
                fd.Title = "Ein Exceltemplate wählen";
                fd.Filter = "Excel-File|*.xlsx";
                // todo: this
                //bool? res = fd.ShowDialog(this);
                bool? res = fd.ShowDialog();
                if (res.Value == true)
                {
                    var exporter = _Factory.CreateVisualOutputHandler();
                    exporter.SetTemplates(System.IO.Path.GetDirectoryName(fd.FileName));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Properties.Resources.MsgBoxTitle);
                log.Error(ex.Message, ex);
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
