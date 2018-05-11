using ExcelPari.Properties;
using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPari
{
    internal class ExcelizerNW : IDisposable
    {
        #region log4net Initialization
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(ExcelizerNW));
        static ExcelizerNW()
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

        private const string TEMPLATE_FILENAME = "Template_NW-Pari.xlsx";
        private Excel.Application _MyApp = null;
        private Excel.Workbook _WorkBook = null;
        private Excel.Worksheet _SheetNW = null;
        private Excel.Worksheet _SheetPari = null;
        IPariDatabase _Database;
        private string _TargetFile;
        private string _TemplateFile = null;
        private Dictionary<string, IWohnungRecord> _WohnungInfos = new Dictionary<string, IWohnungRecord>();
        private KatSortComparer _KatSortComparer = new KatSortComparer();
        private TextNumSortComparer _TextNumSortComparer = new TextNumSortComparer();

        public ExcelizerNW(IPariDatabase database, string locationHint)
        {
            this._Database = database;
            this._TargetFile = locationHint;

            _TemplateFile = Path.Combine(Settings.Default.TemplateLocation, TEMPLATE_FILENAME);
            log.Debug(string.Format(CultureInfo.InvariantCulture, "Settings.TemplateLocation: '{0}'", _TemplateFile));
            if (!File.Exists(_TemplateFile))
            {
                //var sourceDirName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                //_TemplateNWFile = Path.Combine(sourceDirName, TEMPLATE_NW_FILENAME);
                if (!File.Exists(_TemplateFile)) throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "File '{0}' doesn't exist!", _TemplateFile));
            }
        }

        internal void ExportNW(int projektId)
        {
            log.Info("ExportNW");
            var pi = GetProjektInfo(projektId);
            if (pi == null)
            {
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "ProjektInfo with id {0} doesn't exist in  Database!", projektId));
            }
            log.Debug("Getting Excel.Application");
            _MyApp = new Excel.Application();
            log.Debug("Got Excel.Application");
            try
            {
                _WorkBook = _MyApp.Workbooks.Open(_TemplateFile, Missing.Value, ReadOnly: false); //, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //_MyApp.Visible = true;
                //_MyApp.ScreenUpdating = true;
                _SheetNW = GetWorksheet("Template_Nutzwertfestlegung");
                _SheetPari = GetWorksheet("Template_Parifizierung");

                _WohnungInfos.Clear();
                var wohnungen = _Database.GetWohnungen(projektId);
                foreach (var w in wohnungen)
                {
                    _WohnungInfos[w.Top] = w;
                }

                WriteNW(projektId, pi);
                WritePari(projektId, pi); // Nutzflächenanteile

                log.Debug("Deleting Template-Sheets.");
                _MyApp.DisplayAlerts = false;
                _SheetNW.Delete();
                _SheetPari.Delete();
                _MyApp.DisplayAlerts = true;

                if (_TargetFile != null)
                {
                    log.Debug(string.Format(CultureInfo.InvariantCulture, "Saving to '{0}'", _TargetFile));
                    _WorkBook.SaveAs(_TargetFile);
                }
            }
            finally
            {
                if (_TargetFile == null)
                {
                    log.Debug("_MyApp.Visible = true;");
                    _MyApp.Visible = true;
                    log.Debug("_MyApp.ScreenUpdating = true;");
                    _MyApp.ScreenUpdating = true;
                }
            }
        }

        private Excel.Worksheet GetWorksheet(object indexOrName)
        {
            try
            {
                return (Excel.Worksheet)_WorkBook.Worksheets.get_Item(indexOrName);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "Unable to get tab {0} in '{1}'!", indexOrName, _TemplateFile));
            }
        }

        private void WritePari(int projektId, IProjektInfo pi)
        {
            log.Debug("WritePari");
            //Excel.Worksheet targetSheet = _WorkBook.Worksheets.Add();
            //targetSheet.Name = "Parifizierung";
            var targetSheet = GetWorksheet("Parifizierung");
            const int MAX_COL_INDEX = 10;

            // Überschrift Nutzflächenanteile
            CopyCells(_SheetPari, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);

            // bauvorhaben
            CopyCells(_SheetPari, targetSheet, rowIndex1: 1, colIndex1: 0, rowIndex2: 1, colIndex2: 0, copyColumnWidth: false);
            targetSheet.Cells[2, 1] = pi.Bauvorhaben;

            // räume mit kategorien mit zuabschlaginfo
            var raeume = _Database.GetRaeumeWithZuAbschlag(projektId).Where(x => IsNwTopName(x.Top)).OrderBy(x => x.Top, _TextNumSortComparer).ToList();
            int gesSumNutzEinz = GetSumNutz(raeume);
            var raeumeGroupByTop = raeume.GroupBy(x => x.Top);

            // wohnungseigentumsobjekte
            CopyCells(_SheetPari, targetSheet, rowIndex1: 5, colIndex1: 0, rowIndex2: 1, colIndex2: 0, copyColumnWidth: false);
            var nrOfWEO = raeumeGroupByTop.Count();
            targetSheet.Cells[6, 1] = nrOfWEO + " Wohnungseigentumsobjekte";

            // header: Fieldnames
            CopyCells(_SheetPari, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);
            CopyCells(_SheetPari, targetSheet, rowIndex1: 7, colIndex1: 0, rowIndex2: 7, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);

            int targetRowIndex = 9;
            var matrix = new ExcelMatrix(targetRowIndex, MAX_COL_INDEX + 1);
            foreach (var topGroup in raeumeGroupByTop)
            {
                var top = topGroup.Key;
                // Top-Header
                CopyCells(_SheetPari, targetSheet, rowIndex1: 9, colIndex1: 0, rowIndex2: 9, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, top);

                var wohnTypIndex = targetRowIndex;
                IWohnungRecord wohnungRec = null;
                if (_WohnungInfos.TryGetValue(top, out wohnungRec))
                {
                    var wohnTyp = wohnungRec.Typ ?? "";
                    matrix.Add(wohnTypIndex, 10, wohnTyp);
                }

                var rGbKat = topGroup.GroupBy(x => x.Kategorie);
                targetRowIndex += 2;
                int sumNutzEinz = 0;
                foreach (var katGroup in rGbKat.OrderBy(x => x.Key, _KatSortComparer))
                {
                    IKategorieZaRecord kat = katGroup.Key;

                    //bool isPkw = false;

                    if (wohnungRec == null)
                    {
                        // check pkw
                        string pkwWohnTyp = null;
                        if (GetPkwWohnTyp(kat.Widmung, out pkwWohnTyp))
                        {
                            //isPkw = true;
                            matrix.Add(wohnTypIndex, 10, pkwWohnTyp);
                        }
                    }

                    double m2 = 0.0;
                    foreach (var raum in katGroup)
                    {
                        m2 += raum.Flaeche;
                    }

                    var nutzEinz = (int)Math.Round(m2 * kat.ActualNutzwert);
                    sumNutzEinz += nutzEinz;

                    // Kategorie
                    CopyCells(_SheetPari, targetSheet, rowIndex1: 11, colIndex1: 0, rowIndex2: 11, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 1, kat.Lage);
                    matrix.Add(targetRowIndex, 2, kat.Widmung);
                    matrix.Add(targetRowIndex, 3, m2);
                    matrix.Add(targetRowIndex, 4, "m²");
                    matrix.Add(targetRowIndex, 5, kat.ActualNutzwert);
                    matrix.Add(targetRowIndex, 7, nutzEinz);
                    targetRowIndex += 1;

                    var begr = kat.Begrundung.Trim();
                    //if (isPkw)
                    //{
                    //    begr = "als Wohnungseigentumsobjekt";
                    //}
                    //else 
                    if (string.Compare(begr, "als Zuschlag", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        begr = "als Wohnungseigentumszuschlag";
                    }
                    else if (string.Compare(begr, "als Zubehör", StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        begr = "als Wohnungseigentumszubehör";
                    }
                    matrix.Add(targetRowIndex, 2, begr);
                    targetRowIndex += 1;
                }
                // Summe
                // 18-20 wegen Abschlusslinie. Deshalt targetRowIndex - 1
                CopyCells(_SheetPari, targetSheet, rowIndex1: 18, colIndex1: 0, rowIndex2: 20, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex-1);
                targetRowIndex += 1;
                matrix.Add(targetRowIndex, 2, "Summe Nutzwert");
                matrix.Add(targetRowIndex, 3, top);
                matrix.Add(targetRowIndex, 7, sumNutzEinz);
                matrix.Add(targetRowIndex, 8, sumNutzEinz);
                matrix.Add(targetRowIndex, 9, sumNutzEinz * 2);
                matrix.Add(targetRowIndex, 10, gesSumNutzEinz * 2);
                targetRowIndex += 3;
            }
            // Summe Gesamt
            targetRowIndex += 1;
            CopyCells(_SheetPari, targetSheet, rowIndex1: 20, colIndex1: 0, rowIndex2: 20, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 2, "SUMME MINDESTANTEILE");
            matrix.Add(targetRowIndex, 7, gesSumNutzEinz);
            matrix.Add(targetRowIndex, 8, gesSumNutzEinz);
            matrix.Add(targetRowIndex, 9, gesSumNutzEinz * 2);
            matrix.Add(targetRowIndex, 10, gesSumNutzEinz * 2);

            matrix.Write(targetSheet);

        }

        private bool IsNwTopName(string top)
        {
            if (string.IsNullOrEmpty(top)) return false;
            if (Regex.IsMatch(top, "ALLG", RegexOptions.IgnoreCase)) return false;
            return true;
        }

        private int GetSumNutz(List<IRaumZaRecord> raeume)
        {
            int gesSumNutzEinz = 0;
            var raeumeGroupByTop = raeume.GroupBy(x => x.Top);
            {
                foreach (var topGroup in raeumeGroupByTop)
                {
                    var rGbKat = topGroup.GroupBy(x => x.Kategorie);
                    foreach (var katGroup in rGbKat)
                    {
                        IKategorieZaRecord kat = katGroup.Key;
                        double m2 = 0.0;
                        foreach (var raum in katGroup)
                        {
                            m2 += raum.Flaeche;
                        }

                        var nutzEinz = (int)Math.Round(m2 * kat.ActualNutzwert);
                        gesSumNutzEinz += nutzEinz;
                    }
                }
            }
            return gesSumNutzEinz;
        }

        private void WriteNW(int projektId, IProjektInfo pi)
        {
            log.Debug("WriteNW");
            //Excel.Worksheet targetSheet = _WorkBook.Worksheets.Add();
            //targetSheet.Name = "Nutzwertfestlegung";
            var targetSheet = GetWorksheet("Nutzwertfestlegung");

            const int MAX_COL_INDEX = 6;

            // Überschrift Nutzflächenbewertung
            CopyCells(_SheetNW, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);

            // bauvorhaben
            CopyCells(_SheetNW, targetSheet, rowIndex1: 1, colIndex1: 0, rowIndex2: 1, colIndex2: 0, copyColumnWidth: false);
            targetSheet.Cells[2, 1] = pi.Bauvorhaben;

            // header: Fieldnames
            CopyCells(_SheetNW, targetSheet, rowIndex1: 4, colIndex1: 0, rowIndex2: 4, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);

            // kategorien
            //var kategories = _Database.GetKategories(projektId);
            var kategories = _Database.GetKategoriesWithZuAbschlag(projektId).Where(x => IsNwTopName(x.Top)).OrderBy(x => x.Top, _TextNumSortComparer);
            var katGroupByTop = kategories.GroupBy(x => x.Top);
            int targetRowIndex = 6;
            foreach (var topGroup in katGroupByTop)
            {
                var top = topGroup.Key;
                // Top-Header
                CopyCells(_SheetNW, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                targetSheet.Cells[targetRowIndex + 1, 1] = top;

                IWohnungRecord wohnungRec = null;
                var wohnTypRowIndex = targetRowIndex;
                if (_WohnungInfos.TryGetValue(top, out wohnungRec))
                {
                    var wohnTyp = wohnungRec.Typ ?? "";
                    targetSheet.Cells[wohnTypRowIndex + 1, 7] = wohnTyp;
                }

                targetRowIndex += 2;
                foreach (var kat in topGroup.OrderBy(x => x, _KatSortComparer))
                {
                    //var isPkw = false;
                    if (wohnungRec == null)
                    {
                        // check pkw
                        string pkwWohnTyp = null;
                        if (GetPkwWohnTyp(kat.Widmung, out pkwWohnTyp))
                        {
                            //isPkw = true;
                            targetSheet.Cells[wohnTypRowIndex + 1, 7] = pkwWohnTyp;
                        }
                    }

                    // Kategorie
                    CopyCells(_SheetNW, targetSheet, rowIndex1: 8, colIndex1: 0, rowIndex2: 8, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    targetSheet.Cells[targetRowIndex + 1, 2] = kat.Lage;
                    targetSheet.Cells[targetRowIndex + 1, 3] = kat.Widmung;
                    double rnwd;
                    var rnw = kat.RNW.Replace(',', '.');
                    if (!double.TryParse(rnw, NumberStyles.Any, CultureInfo.InvariantCulture, out rnwd))
                    {
                        targetSheet.Cells[targetRowIndex + 1, 4] = kat.RNW;
                    }
                    else
                    {
                        targetSheet.Cells[targetRowIndex + 1, 4] = rnwd;
                    }

                    //if (isPkw)
                    //{
                    //    targetSheet.Cells[targetRowIndex + 1, 6] = "als Wohnungseigentumsobjekt";
                    //}
                    //else
                    //{
                        targetSheet.Cells[targetRowIndex + 1, 6] = kat.Begrundung;
                    //}

                    if (kat.ZuAbschlaege.Count > 0)
                    {
                        targetSheet.Cells[targetRowIndex + 1, 7] = "";
                        // auflistung der prozente
                        foreach (var zuAb in kat.ZuAbschlaege)
                        {
                            targetRowIndex += 1;
                            CopyCells(_SheetNW, targetSheet, rowIndex1: 8, colIndex1: 0, rowIndex2: 8, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                            targetSheet.Cells[targetRowIndex + 1, 2] = "";
                            targetSheet.Cells[targetRowIndex + 1, 3] = "";
                            targetSheet.Cells[targetRowIndex + 1, 4] = "";
                            targetSheet.Cells[targetRowIndex + 1, 5] = "";
                            targetSheet.Cells[targetRowIndex + 1, 6] = string.Format(CultureInfo.CurrentCulture, "{0} {1}%", zuAb.Beschreibung, zuAb.Prozent);
                            targetSheet.Cells[targetRowIndex + 1, 7] = "";
                        }
                        targetSheet.Cells[targetRowIndex + 1, 7] = kat.ActualNutzwert;
                    }
                    else
                    {
                        targetSheet.Cells[targetRowIndex + 1, 7] = kat.Nutzwert;
                    }
                    targetRowIndex += 2;
                }
                targetRowIndex += 1;
            }
        }

        private bool GetPkwWohnTyp(string widmung, out string pkwWohnTyp)
        {
            pkwWohnTyp = null;
            if (widmung == null) return false;

            var wid = widmung.ToUpperInvariant();
            if (wid.Contains("PKW"))
            {
                if (wid.Contains("STELLPLATZ"))
                {
                    pkwWohnTyp = "PKW nicht überdacht";
                }
                else
                {
                    pkwWohnTyp = "PKW überdacht";
                }
                return true;
            }
            return false;
        }


        private void CopyCells(Excel.Worksheet sourceSheet, Excel.Worksheet targetSheet, int rowIndex1, int colIndex1, int rowIndex2, int colIndex2, bool copyColumnWidth = false, int targetRowIndex1 = -1)
        {
            var cell1Bez = GetCellBez(rowIndex1, colIndex1);
            var cell2Bez = GetCellBez(rowIndex2, colIndex2);
            var range1 = sourceSheet.Range[cell1Bez, cell2Bez];

            Excel.Range range2 = null;
            if (targetRowIndex1 == -1)
            {
                range2 = targetSheet.Range[cell1Bez, cell2Bez];
            }
            else
            {
                int targetRowIndex2 = targetRowIndex1 + (rowIndex2 - rowIndex1);
                cell1Bez = GetCellBez(targetRowIndex1, colIndex1);
                cell2Bez = GetCellBez(targetRowIndex2, colIndex2);
                range2 = targetSheet.Range[cell1Bez, cell2Bez];
            }

            range1.Copy(Type.Missing);
            //R2.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            range2.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            if (copyColumnWidth)
                range2.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            range2.RowHeight = range1.RowHeight;
        }

        private IProjektInfo GetProjektInfo(int projektId)
        {
            var projInfos = _Database.ListProjInfos();
            return projInfos.FirstOrDefault(x => x.ProjektId == projektId);
        }

        private static string GetCellBez(int rowIndex, int colIndex)
        {
            return TranslateColumnIndexToName(colIndex) + (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static String TranslateColumnIndexToName(int index)
        {
            //assert (index >= 0);

            int quotient = (index) / 26;

            if (quotient > 0)
            {
                return TranslateColumnIndexToName(quotient - 1) + (char)((index % 26) + 65);
            }
            else
            {
                return "" + (char)((index % 26) + 65);
            }
        }

        public void Dispose()
        {
            if (_TargetFile == null) return; // Leave Excel open.

            log.Debug("Dispose");

            if (_WorkBook != null) _WorkBook.Close(false, Missing.Value, Missing.Value);
            if (_MyApp != null) _MyApp.Quit();

            releaseObject(_SheetNW);
            releaseObject(_SheetPari);
            releaseObject(_WorkBook);
            releaseObject(_MyApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                if (obj != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
