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
    internal class ExcelizerNF : IDisposable
    {
        #region log4net Initialization
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(ExcelizerNF));
        static ExcelizerNF()
        {
            if (log4net.LogManager.GetRepository(Assembly.GetExecutingAssembly()).Configured == false)
            {
                log4net.Config.XmlConfigurator.ConfigureAndWatch(
                    new FileInfo(
                        Path.Combine(
                            new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName,
                            "_log4net.config"
                        )
                    )
                );
            }
        }
        #endregion

        private const string TEMPLATE_FILENAME = "Template_NF-Pari.xlsx";
        private Excel.Application _MyApp = null;
        private Excel.Workbook _WorkBook = null;
        private Excel.Worksheet _SheetSumme = null;
        private Excel.Worksheet _SheetAllgemein = null;
        private Excel.Worksheet _SheetTop = null;
        private Excel.Worksheet _SheetAbstell = null;
        IPariDatabase _Database;
        private string _TargetFile;
        private string _TemplateFile = null;
        private GeschossSortComparer _GeschossSortComparer = new GeschossSortComparer();
        private TopAllgComparer _TopAllgComparer = new TopAllgComparer();
        private readonly TextNumSortComparer _TextNumSortComparer = new TextNumSortComparer();
        private readonly Dictionary<string, IWohnungRecord> _WohnungInfos = new Dictionary<string, IWohnungRecord>();

        public ExcelizerNF(IPariDatabase database, string locationHint)
        {
            _Database = database;
            _TargetFile = locationHint;

            _TemplateFile = Path.Combine(Settings.Default.TemplateLocation, TEMPLATE_FILENAME);
            log.Debug(string.Format(CultureInfo.InvariantCulture, "Settings.TemplateLocation: '{0}'", _TemplateFile));
            if (!File.Exists(_TemplateFile))
            {
                if (!File.Exists(_TemplateFile)) throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "File '{0}' doesn't exist!", _TemplateFile));
            }
        }

        internal void ExportNf(int projektId)
        {
            log.Info("ExportNF");
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
                _SheetSumme = GetWorksheet("Template_Summe");
                _SheetAllgemein = GetWorksheet("Template_Allgemein");
                _SheetTop = GetWorksheet("Template_Top");
                var topSheet = GetWorksheet("Top");
                _SheetAbstell = GetWorksheet("Template_Abstellplätze");

                _WohnungInfos.Clear();
                var wohnungen = _Database.GetWohnungen(projektId);
                foreach (var w in wohnungen)
                {
                    _WohnungInfos[w.Top] = w;
                }

                // Here it starts
                var summeSheet = WriteSumme(projektId, pi);
                WriteAllgemein(projektId, pi);
                WriteTops(projektId, pi, summeSheet);
                WriteAbstell(projektId, pi);

                // Cleanup sheets
                log.Debug("Deleting Template-Sheets.");
                _MyApp.DisplayAlerts = false;
                _SheetSumme.Delete();
                _SheetAllgemein.Delete();
                _SheetTop.Delete();
                _SheetAbstell.Delete();
                topSheet.Delete();
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
            catch (Exception)
            {
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, "Unable to get tab {0} in '{1}'!", indexOrName, _TemplateFile));
            }
        }

        private void WriteAllgemein(int projektId, IProjektInfo pi)
        {
            log.Debug("WriteAllgemein");
            var targetSheet = GetWorksheet("ALLGEMEIN");
            const int MAX_COL_INDEX = 4;

            // Überschrift PKW ABSTELLPLÄTZE
            CopyCells(_SheetAllgemein, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);


            // raume
            var raeume = _Database.GetRaeume(projektId);
#if !ALTEVARIANTE
            List<IRaumRecord> tops, allgs, pkws;
            SplitRaeume(raeume, out pkws, out tops, out allgs);
            var allgRaeumePerTop = allgs.GroupBy(x => x.Top).OrderBy(x => x.Key, _TopAllgComparer).ToList();
#else
            // alte variante
            var allgRaeumePerTop = raeume.Where(x => IsAllgForSumme(x)).GroupBy(x => x.Top).OrderBy(x => x.Key, _TopAllgComparer).ToList();
#endif

            int targetRowIndex = 4;
            var matrix = new ExcelMatrix(startRowIndex: targetRowIndex, nrOfCols: MAX_COL_INDEX + 1);
            double sumGes = 0.0;
            foreach (var raumKvp in allgRaeumePerTop)
            {
                var art = raumKvp.Key;
                var topBez = art.Replace("ALLG", "");
                if (string.IsNullOrEmpty(topBez)) topBez = "AUSSENANLAGEN";
                double sumArt = 0.0;

                CopyCells(_SheetAllgemein, targetSheet, rowIndex1: 4, colIndex1: 0, rowIndex2: 5, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, topBez);
                targetRowIndex += 2;
                var perGesch = raumKvp.GroupBy(x => x.Lage).OrderBy(x => x.Key, _GeschossSortComparer);
                bool firstGesch = true;
                foreach (var rpg in perGesch)
                {
                    if (firstGesch)
                    {
                        firstGesch = false;
                    }
                    else
                    {
                        targetRowIndex++;
                    }
                    var geschoss = rpg.Key;

                    var sumProGesch = rpg.Sum(x => x.Flaeche);
                    sumGes += sumProGesch;
                    sumArt += sumProGesch;

                    CopyCells(_SheetAllgemein, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 0, geschoss);

                    var sortedR = rpg.OrderBy(x => x.Raum);
                    foreach (var r in sortedR)
                    {
                        matrix.Add(targetRowIndex, 2, r.Raum);
                        matrix.Add(targetRowIndex, 3, r.Flaeche);
                        matrix.Add(targetRowIndex, 4, "m²");
                        targetRowIndex++;
                        CopyCells(_SheetAllgemein, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    }
                }
                CopyCells(_SheetAllgemein, targetSheet, rowIndex1: 12, colIndex1: 0, rowIndex2: 14, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                targetRowIndex++;

                matrix.Add(targetRowIndex, 2, "Summe");
                matrix.Add(targetRowIndex, 3, sumArt);
                matrix.Add(targetRowIndex, 4, "m²");

                targetRowIndex += 2;
            }
            targetRowIndex += 2;
            CopyCells(_SheetAllgemein, targetSheet, rowIndex1: 39, colIndex1: 0, rowIndex2: 40, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "SUMME ALLER ALLGEMEINFLÄCHEN");
            matrix.Add(targetRowIndex, 3, sumGes);
            matrix.Add(targetRowIndex, 4, "m²");
            targetRowIndex++;
            matrix.Add(targetRowIndex, 0, "nicht für Nutzwert relevant");

            matrix.Write(targetSheet);
        }

        private class AbstellAndZuschlag
        {
            public IRaumRecord Abstellplatz { get; set; }
            public readonly List<IRaumRecord> Zuschlaege;

            public AbstellAndZuschlag()
            {
                Zuschlaege = new List<IRaumRecord>();
            }
        }

        private void WriteAbstell(int projektId, IProjektInfo pi)
        {
            log.Debug("WriteAbstell");
            var targetSheet = GetWorksheet("Abstellplätze");
            const int MAX_COL_INDEX = 6;

            // Überschrift PKW ABSTELLPLÄTZE
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);


            // Abstellplätze
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 4, colIndex1: 0, rowIndex2: 5, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);

            // Gesamtfläche
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 5, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);

            // raume
            var raeume = _Database.GetRaeume(projektId);
#if !ALTEVARIANTE
            List<IRaumRecord> tops, allgs, pkws;
            SplitRaeume(raeume, out pkws, out tops, out allgs);
            List<AbstellAndZuschlag> abAndZu = SplitPkws(pkws);
            var pkwRaeumePerWidmung = abAndZu.GroupBy(x => x.Abstellplatz.Widmung);
#else
            // alte variante
            var pkwRaeumePerWidmung = raeume.Where(x => IsPkwForSumme(x)).GroupBy(x => x.Widmung);
#endif
            int targetRowIndex = 6;
            var matrix = new ExcelMatrix(startRowIndex: targetRowIndex, nrOfCols: MAX_COL_INDEX + 1);
            double sumGes = 0.0;
            foreach (var raumKvp in pkwRaeumePerWidmung)
            {
                var art = raumKvp.Key;
                double sumArt = 0.0;

                var perGesch = raumKvp.GroupBy(x => x.Abstellplatz.Lage).OrderBy(x => x.Key, _GeschossSortComparer);
                foreach (var rpg in perGesch)
                {
                    var geschoss = rpg.Key;

                    var sumProGesch = rpg.Sum(x => x.Abstellplatz.Flaeche);
                    sumGes += sumProGesch;
                    sumArt += sumProGesch;

                    CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 0, geschoss);

                    var rpgOrderByTops = rpg.OrderBy(x => x.Abstellplatz.Top, _TextNumSortComparer).ToList();
                    foreach (var r in rpgOrderByTops)
                    {
                        if (r.Zuschlaege.Count == 0)
                        {
                            matrix.Add(targetRowIndex, 1, r.Abstellplatz.Top);
                            matrix.Add(targetRowIndex, 2, r.Abstellplatz.Widmung);
                            matrix.Add(targetRowIndex, 5, r.Abstellplatz.Flaeche);
                            matrix.Add(targetRowIndex, 6, "m²");
                            targetRowIndex++;
                            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);

                        }
                        else
                        {
                            // todo: add sum to sumGes and sumArt
                            matrix.Add(targetRowIndex, 1, r.Abstellplatz.Top);
                            matrix.Add(targetRowIndex, 2, r.Abstellplatz.Widmung);
                            matrix.Add(targetRowIndex, 3, r.Abstellplatz.Flaeche);
                            matrix.Add(targetRowIndex, 4, "m²");
                            targetRowIndex++;
                            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                            var sumZusch = 0.0;
                            foreach (var zuschl in r.Zuschlaege)
                            {
                                var wid = zuschl.Widmung ?? "";
                                var begrundung = zuschl.Begrundung ?? "";
                                matrix.Add(targetRowIndex, 2, wid + " " + begrundung);
                                matrix.Add(targetRowIndex, 3, zuschl.Flaeche);
                                matrix.Add(targetRowIndex, 4, "m²");
                                sumZusch += zuschl.Flaeche;
                                targetRowIndex++;
                                CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                            }
                            sumArt += sumZusch;
                            sumGes += sumZusch;
                            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 10, colIndex1: 0, rowIndex2: 10, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                            matrix.Add(targetRowIndex, 1, "Summe");
                            matrix.Add(targetRowIndex, 2, r.Abstellplatz.Top);
                            matrix.Add(targetRowIndex, 5, r.Abstellplatz.Flaeche + sumZusch);
                            matrix.Add(targetRowIndex, 6, "m²");
                            targetRowIndex++;
                            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                        }

                    }
                }
                CopyCells(_SheetAbstell, targetSheet, rowIndex1: 14, colIndex1: 0, rowIndex2: 16, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                targetRowIndex++;

                matrix.Add(targetRowIndex, 1, "Summe");
                matrix.Add(targetRowIndex, 2, art);
                matrix.Add(targetRowIndex, 5, sumArt);
                matrix.Add(targetRowIndex, 6, "m²");

                targetRowIndex += 2;
            }
            targetRowIndex += 2;
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 35, colIndex1: 0, rowIndex2: 35, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "SUMME ALLE PKW ABSTELLPLÄTZE");
            matrix.Add(targetRowIndex, 5, sumGes);
            matrix.Add(targetRowIndex, 6, "m²");

            matrix.Write(targetSheet);
        }

        private List<AbstellAndZuschlag> SplitPkws(List<IRaumRecord> all)
        {
            var pkws = all.Where(x => Regex.IsMatch(x.Widmung, "PKW", RegexOptions.IgnoreCase));
            var zuschl = all.Where(x => !Regex.IsMatch(x.Widmung, "PKW", RegexOptions.IgnoreCase)).ToList();

            var abAndZus = new List<AbstellAndZuschlag>();
            foreach (var pkw in pkws)
            {
                var abAndZu = new AbstellAndZuschlag();
                abAndZu.Abstellplatz = pkw;
                abAndZus.Add(abAndZu);

                var toRemove = new List<IRaumRecord>();
                foreach (var z in zuschl)
                {
                    if (string.Compare(z.Top ,pkw.Top, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        abAndZu.Zuschlaege.Add(z);
                        toRemove.Add(z);
                    }
                }

                foreach (var rem in toRemove)
                {
                    zuschl.Remove(rem);
                }
            }

            if (zuschl.Count > 0)
            {
                var msg = string.Join(", ", zuschl.Select(x => (x.Widmung ?? "") + " " + (x.Begrundung ?? "")));
                log.WarnFormat(CultureInfo.CurrentCulture, "Es gibt Zuschläge ohne PKWs! {0}",msg);
            }

            return abAndZus;
        }

        private void WriteAbstellAlt(int projektId, IProjektInfo pi)
        {
            log.Debug("WriteAbstell");
            var targetSheet = GetWorksheet("Abstellplätze");
            const int MAX_COL_INDEX = 4;

            // Überschrift PKW ABSTELLPLÄTZE
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);


            // Abstellplätze
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 4, colIndex1: 0, rowIndex2: 5, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);

            // Gesamtfläche
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 5, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);

            // raume
            var raeume = _Database.GetRaeume(projektId);
#if !ALTEVARIANTE
            List<IRaumRecord> tops, allgs, pkws;
            SplitRaeume(raeume, out pkws, out tops, out allgs);
            var pkwRaeumePerWidmung = pkws.GroupBy(x => x.Widmung);
#else
            // alte variante
            var pkwRaeumePerWidmung = raeume.Where(x => IsPkwForSumme(x)).GroupBy(x => x.Widmung);
#endif
            int targetRowIndex = 6;
            var matrix = new ExcelMatrix(startRowIndex: targetRowIndex, nrOfCols: MAX_COL_INDEX + 1);
            double sumGes = 0.0;
            foreach (var raumKvp in pkwRaeumePerWidmung)
            {
                var art = raumKvp.Key;
                double sumArt = 0.0;

                var perGesch = raumKvp.GroupBy(x => x.Lage).OrderBy(x => x.Key, _GeschossSortComparer);
                foreach (var rpg in perGesch)
                {
                    var geschoss = rpg.Key;

                    var sumProGesch = rpg.Sum(x => x.Flaeche);
                    sumGes += sumProGesch;
                    sumArt += sumProGesch;

                    CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 0, geschoss);

                    foreach (var r in rpg)
                    {
                        matrix.Add(targetRowIndex, 1, r.Top);
                        matrix.Add(targetRowIndex, 2, r.Widmung);
                        matrix.Add(targetRowIndex, 3, r.Flaeche);
                        matrix.Add(targetRowIndex, 4, "m²");
                        targetRowIndex++;
                        CopyCells(_SheetAbstell, targetSheet, rowIndex1: 6, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    }
                }
                CopyCells(_SheetAbstell, targetSheet, rowIndex1: 22, colIndex1: 0, rowIndex2: 24, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                targetRowIndex++;

                matrix.Add(targetRowIndex, 1, "Summe");
                matrix.Add(targetRowIndex, 2, art);
                matrix.Add(targetRowIndex, 3, sumArt);
                matrix.Add(targetRowIndex, 4, "m²");

                targetRowIndex += 2;
            }
            targetRowIndex += 2;
            CopyCells(_SheetAbstell, targetSheet, rowIndex1: 46, colIndex1: 0, rowIndex2: 46, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "SUMME ALLE PKW ABSTELLPLÄTZE");
            matrix.Add(targetRowIndex, 3, sumGes);
            matrix.Add(targetRowIndex, 4, "m²");

            matrix.Write(targetSheet);
        }

        private void WriteTops(int projektId, IProjektInfo pi, Excel.Worksheet summeSheet)
        {
            log.Debug("WriteTops");
            var topSheet = GetWorksheet("Top");
            const int MAX_COL_INDEX = 4;

            // raume
            var raeume = _Database.GetRaeume(projektId);

#if !ALTEVARIANTE
            List<IRaumRecord> tops, allgs, pkws;
            SplitRaeume(raeume,out pkws, out tops, out allgs);
#else
            // alte variante
            var tops = raeume.Where(x => IsTopForSumme(x)).ToList();
#endif
            var raeumePerTop = tops.GroupBy(x => x.Top).OrderBy(x => x.Key, _TextNumSortComparer).ToList();
            foreach (var raumKvp in raeumePerTop)
            {
                var top = raumKvp.Key;
                topSheet.Copy(topSheet);
                var targetSheet = GetWorksheet("Top (2)");
                targetSheet.Name = top;

                // Überschrift Nutzflächenanteile
                CopyCells(_SheetTop, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);

                // bauvorhaben
                CopyCells(_SheetTop, targetSheet, rowIndex1: 1, colIndex1: 0, rowIndex2: 1, colIndex2: 0, copyColumnWidth: false);
                targetSheet.Cells[2, 1] = pi.Bauvorhaben;

                // Topname
                CopyCells(_SheetTop, targetSheet, rowIndex1: 5, colIndex1: 0, rowIndex2: 5, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);
                targetSheet.Cells[6, 1] = top;

                // Start Matrix
                int targetRowIndex = 9;
                var matrix = new ExcelMatrix(startRowIndex: targetRowIndex, nrOfCols: MAX_COL_INDEX + 1);

                // Wohnung
                CopyCells(_SheetTop, targetSheet, rowIndex1: 9, colIndex1: 0, rowIndex2: 10, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);
                IWohnungRecord wohnungRec = null;
                if (_WohnungInfos.TryGetValue(top, out wohnungRec))
                {
                    var wohnTyp = wohnungRec.Typ ?? "";
                    matrix.Add(targetRowIndex, 0, wohnTyp);
                }
                else
                {
                    matrix.Add(targetRowIndex, 0, "Wohnungstyp unbekannt!");
                }

                // Per Geschoß ohne Zuschlag und Zubehör
                targetRowIndex += 2;
                double sumGesamt = 0.0;
                var perGesch = raumKvp.GroupBy(x => x.Lage).OrderBy(x => x.Key, _GeschossSortComparer).ToList();
                foreach (var rpg in perGesch)
                {
                    var raeumeOhneZuSchlagUndZuBehoer = rpg.Where(x => IsTopWithoutZuschlagAndZubehoer(x)).OrderBy(x => x.Raum).ToList();
                    if (raeumeOhneZuSchlagUndZuBehoer.Count == 0) continue;

                    var geschoss = rpg.Key;
                    var sum = raeumeOhneZuSchlagUndZuBehoer.Sum(x => x.Flaeche);

                    CopyCells(_SheetTop, targetSheet, rowIndex1: 11, colIndex1: 0, rowIndex2: 11, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 0, geschoss);

                    foreach (var r in raeumeOhneZuSchlagUndZuBehoer)
                    {
                        matrix.Add(targetRowIndex, 2, r.Raum);
                        matrix.Add(targetRowIndex, 3, r.Flaeche);
                        matrix.Add(targetRowIndex, 4, "m²");
                        targetRowIndex++;
                        CopyCells(_SheetTop, targetSheet, rowIndex1: 12, colIndex1: 0, rowIndex2: 12, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    }

                    targetRowIndex++;
                    CopyCells(_SheetTop, targetSheet, rowIndex1: 20, colIndex1: 0, rowIndex2: 21, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 2, "Summe");
                    matrix.Add(targetRowIndex, 3, sum);
                    matrix.Add(targetRowIndex, 4, "m²");

                    targetRowIndex += 2;

                    sumGesamt += sum;
                }

                // SUMME Fläche TOP ...
                CopyCells(_SheetTop, targetSheet, rowIndex1: 22, colIndex1: 0, rowIndex2: 22, colIndex2: MAX_COL_INDEX, copyColumnWidth: true, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, "SUMME Fläche " + top);
                matrix.Add(targetRowIndex, 3, sumGesamt);
                matrix.Add(targetRowIndex, 4, "m²");

                targetRowIndex += 4;

                // ZUSCHLAG FLÄCHEN
                var sumZuschlag = 0.0;
                CopyCells(_SheetTop, targetSheet, rowIndex1: 9, colIndex1: 0, rowIndex2: 10, colIndex2: MAX_COL_INDEX, copyColumnWidth: true, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, "ZUSCHLAG FLÄCHEN");
                targetRowIndex += 2;
                foreach (var rpg in perGesch)
                {
                    var raeumeZuschlag = rpg.Where(x => IsZuschlag(x)).OrderBy(x => x.Raum).ToList();
                    if (raeumeZuschlag.Count == 0) continue;

                    var sum = raeumeZuschlag.Sum(x => x.Flaeche);

                    var geschoss = rpg.Key;
                    CopyCells(_SheetTop, targetSheet, rowIndex1: 11, colIndex1: 0, rowIndex2: 11, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 0, geschoss);

                    foreach (var r in raeumeZuschlag)
                    {
                        matrix.Add(targetRowIndex, 2, r.Raum);
                        matrix.Add(targetRowIndex, 3, r.Flaeche);
                        matrix.Add(targetRowIndex, 4, "m²");
                        targetRowIndex++;
                        CopyCells(_SheetTop, targetSheet, rowIndex1: 12, colIndex1: 0, rowIndex2: 12, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    }
                    sumGesamt += sum;
                    sumZuschlag += sum;
                }
                targetRowIndex++;
                CopyCells(_SheetTop, targetSheet, rowIndex1: 20, colIndex1: 0, rowIndex2: 21, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 2, "Summe Zuschlag");
                matrix.Add(targetRowIndex, 3, sumZuschlag);
                matrix.Add(targetRowIndex, 4, "m²");

                targetRowIndex += 4;

                // ZUBEHÖR FLÄCHEN
                var sumZubehoer = 0.0;
                CopyCells(_SheetTop, targetSheet, rowIndex1: 9, colIndex1: 0, rowIndex2: 10, colIndex2: MAX_COL_INDEX, copyColumnWidth: true, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, "ZUBEHÖR FLÄCHEN");
                targetRowIndex += 2;
                foreach (var rpg in perGesch)
                {
                    var raeumeZuschlag = rpg.Where(x => IsZubehoer(x)).OrderBy(x => x.Raum).ToList();
                    if (raeumeZuschlag.Count == 0) continue;

                    var geschoss = rpg.Key;
                    var sum = raeumeZuschlag.Sum(x => x.Flaeche);
                    CopyCells(_SheetTop, targetSheet, rowIndex1: 11, colIndex1: 0, rowIndex2: 11, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    matrix.Add(targetRowIndex, 0, geschoss);

                    foreach (var r in raeumeZuschlag)
                    {
                        matrix.Add(targetRowIndex, 2, r.Raum);
                        matrix.Add(targetRowIndex, 3, r.Flaeche);
                        matrix.Add(targetRowIndex, 4, "m²");
                        targetRowIndex++;
                        CopyCells(_SheetTop, targetSheet, rowIndex1: 12, colIndex1: 0, rowIndex2: 12, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                    }
                    sumGesamt += sum;
                    sumZubehoer += sum;
                }

                targetRowIndex++;
                CopyCells(_SheetTop, targetSheet, rowIndex1: 20, colIndex1: 0, rowIndex2: 21, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 2, "Summe Zubehör");
                matrix.Add(targetRowIndex, 3, sumZubehoer);
                matrix.Add(targetRowIndex, 4, "m²");

                targetRowIndex += 2;
                // SUMME Flächen Zuschlag/Zubehör TOP ...
                CopyCells(_SheetTop, targetSheet, rowIndex1: 22, colIndex1: 0, rowIndex2: 22, colIndex2: MAX_COL_INDEX, copyColumnWidth: true, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, "SUMME Flächen Zuschlag/Zubehör " + top);
                matrix.Add(targetRowIndex, 3, sumZubehoer + sumZuschlag);
                matrix.Add(targetRowIndex, 4, "m²");

                // SUMME ALLER FLÄCHE TOP ...
                targetRowIndex += 4;
                CopyCells(_SheetTop, targetSheet, rowIndex1: 45, colIndex1: 0, rowIndex2: 45, colIndex2: MAX_COL_INDEX, copyColumnWidth: true, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, "SUMME ALLER FLÄCHEN " + top);
                matrix.Add(targetRowIndex, 3, sumGesamt);
                matrix.Add(targetRowIndex, 4, "m²");

                matrix.Write(targetSheet);
            }
        }

        private Excel.Worksheet WriteSumme(int projektId, IProjektInfo pi)
        {
            log.Debug("WriteSumme");
            var targetSheet = GetWorksheet("Summe");
            const int MAX_COL_INDEX = 4;

            // Überschrift Nutzflächenanteile
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 0, colIndex1: 0, rowIndex2: 0, colIndex2: MAX_COL_INDEX, copyColumnWidth: true);

            // bauvorhaben
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 1, colIndex1: 0, rowIndex2: 1, colIndex2: 0, copyColumnWidth: false);
            targetSheet.Cells[2, 1] = pi.Bauvorhaben;

            // Gesamtfläche
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 5, colIndex1: 0, rowIndex2: 6, colIndex2: MAX_COL_INDEX, copyColumnWidth: false);


            // raume
            var raeume = _Database.GetRaeume(projektId);
            var wohnungen = _Database.GetWohnungen(projektId);
#if !ALTEVARIANTE
            List<IRaumRecord> tops, allgs, pkws;
            SplitRaeume(raeume, out pkws, out tops, out allgs);
#else
            // alte variante
            var tops = raeume.Where(x => IsTopForSumme(x)).ToList();
            var pkws = raeume.Where(x => IsPkwForSumme(x)).ToList();
            var allgs = raeume.Where(x => IsAllgForSumme(x)).ToList();
#endif
            // Liste 1: Tops und PKW
            double sumTopAndPKW = 0.0;
            int targetRowIndex = 7;
            var matrix = new ExcelMatrix(startRowIndex: targetRowIndex, nrOfCols: MAX_COL_INDEX + 1);
            // tops
            var raeumePerTop = tops.GroupBy(x => x.Top).OrderBy(x => x.Key, _TextNumSortComparer).ToList();
            foreach (var raumKvp in raeumePerTop)
            {
                var top = raumKvp.Key;
                var summe = raumKvp.Sum(x => x.Flaeche);
                var wohnung = wohnungen.FirstOrDefault(x => x.Top == top);
                var widmung = (wohnung != null) ? (wohnung.Widmung ?? "Keine Widmung in Top-Block!") : "Keine Widmung in Top-Block!";
                sumTopAndPKW += summe;
                CopyCells(_SheetSumme, targetSheet, rowIndex1: 7, colIndex1: 0, rowIndex2: 7, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
                matrix.Add(targetRowIndex, 0, top);
                matrix.Add(targetRowIndex, 1, widmung);
                matrix.Add(targetRowIndex, 3, summe);
                matrix.Add(targetRowIndex, 4, "m²");
                targetRowIndex++;
            }
            // PKV-Abstellplätze
            var sumPkwAbst = pkws.Sum(x => x.Flaeche);
            sumTopAndPKW += sumPkwAbst;
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 7, colIndex1: 0, rowIndex2: 7, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "PKW - ABSTELLPLÄTZE");
            matrix.Add(targetRowIndex, 3, sumPkwAbst);
            matrix.Add(targetRowIndex, 4, "m²");

            targetRowIndex++;

            // Summe
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 20, colIndex1: 0, rowIndex2: 22, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            targetRowIndex++;
            matrix.Add(targetRowIndex, 0, "Summe");
            matrix.Add(targetRowIndex, 3, sumTopAndPKW);
            matrix.Add(targetRowIndex, 4, "m²");
            targetRowIndex += 2;
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 23, colIndex1: 0, rowIndex2: 23, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "Gesamtfläche aller TOP's ohne Allgemeinflächen");
            matrix.Add(targetRowIndex, 3, sumTopAndPKW);
            matrix.Add(targetRowIndex, 4, "m²");

            targetRowIndex += 4;
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 27, colIndex1: 0, rowIndex2: 27, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "Gesamtfläche aller TOP's ohne Allgemeinflächen");
            matrix.Add(targetRowIndex, 3, sumTopAndPKW);
            matrix.Add(targetRowIndex, 4, "m²");

            // allgemeinflächen
            var sumAllgs = allgs.Sum(x => x.Flaeche);
            targetRowIndex++;
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 27, colIndex1: 0, rowIndex2: 27, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            matrix.Add(targetRowIndex, 0, "Allgemeinflächen");
            matrix.Add(targetRowIndex, 3, sumAllgs);
            matrix.Add(targetRowIndex, 4, "m²");

            // gesamtsumme
            targetRowIndex++;
            CopyCells(_SheetSumme, targetSheet, rowIndex1: 29, colIndex1: 0, rowIndex2: 31, colIndex2: MAX_COL_INDEX, copyColumnWidth: false, targetRowIndex1: targetRowIndex);
            targetRowIndex += 2;
            matrix.Add(targetRowIndex, 0, "Gesamtfläche aller TOP's mit Allgemeinflächen");
            matrix.Add(targetRowIndex, 3, sumTopAndPKW + sumAllgs);
            matrix.Add(targetRowIndex, 4, "m²");
            matrix.Write(targetSheet);

            return targetSheet;
        }

        private IProjektInfo GetProjektInfo(int projektId)
        {
            var projInfos = _Database.ListProjInfos();
            return projInfos.FirstOrDefault(x => x.ProjektId == projektId);
        }

        #region ExcelHelper
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
        #endregion

        #region Kriterien
#if !ALTEVARIANTE
        private void SplitRaeume(List<IRaumRecord> raeume, out List<IRaumRecord> pkws, out List<IRaumRecord> tops, out List<IRaumRecord> allgs)
        {
            tops = new List<IRaumRecord>();
            allgs = new List<IRaumRecord>();

            var notEmptyRaeume = raeume.Where(IsNotEmpty).ToList();
            pkws = notEmptyRaeume.Where(IsPkwForSumme).ToList();
            var pkwTops = pkws.Select(x => x.Top).ToList();

            foreach (var raumRecord in notEmptyRaeume)
            {
                if (pkws.Contains(raumRecord)) continue;
                if (IsAllgForSumme(raumRecord))
                {
                    allgs.Add(raumRecord);
                    continue;
                }

                if (pkwTops.Contains(raumRecord.Top))
                {
                    pkws.Add(raumRecord);
                }
                else
                {
                    tops.Add(raumRecord);
                }
            }
        }

        private bool IsNotEmpty(IRaumRecord r)
        {
            if (string.IsNullOrEmpty(r.Top) || r.Top.Trim() == string.Empty) return false;
            return true;
        }
#else
        // alte variante
        private bool IsTopForSumme(IRaumRecord r)
        {
            if (string.IsNullOrEmpty(r.Top) || r.Top.Trim() == string.Empty) return false;
            if (IsAllgForSumme(r)) return false;
            if (IsPkwForSumme(r)) return false;
            return true;
        }
#endif
        private bool IsPkwForSumme(IRaumRecord r)
        {
            if (string.IsNullOrEmpty(r.Top) || r.Top.Trim() == string.Empty) return false;
            if (Regex.IsMatch(r.Widmung, "PKW", RegexOptions.IgnoreCase)) return true;
            else return false;
        }

        private bool IsAllgForSumme(IRaumRecord r)
        {
            if (string.IsNullOrEmpty(r.Top) || r.Top.Trim() == string.Empty) return false;
            if (Regex.IsMatch(r.Top, "ALLG", RegexOptions.IgnoreCase)) return true;
            else return false;
        }

        private bool IsTopWithoutZuschlagAndZubehoer(IRaumRecord r)
        {
            if (IsZubehoer(r)) return false;
            if (IsZuschlag(r)) return false;
            return true;
        }

        private bool IsZuschlag(IRaumRecord r)
        {
            return Regex.IsMatch(r.Begrundung, "Zuschlag", RegexOptions.IgnoreCase);
        }

        private bool IsZubehoer(IRaumRecord r)
        {
            return Regex.IsMatch(r.Begrundung, "Zubehör", RegexOptions.IgnoreCase);
        }


        #endregion

        #region Dispose
        public void Dispose()
        {
            if (_TargetFile == null) return; // Leave Excel open.

            log.Debug("Dispose");

            if (_WorkBook != null) _WorkBook.Close(false, Missing.Value, Missing.Value);
            if (_MyApp != null) _MyApp.Quit();

            releaseObject(_SheetSumme);
            releaseObject(_SheetAllgemein);
            releaseObject(_SheetTop);
            releaseObject(_SheetAbstell);
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
        #endregion

    }
}
