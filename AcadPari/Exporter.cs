using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
#if BRX_APP
using _AcAp = Bricscad.ApplicationServices;
//using _AcBr = Teigha.BoundaryRepresentation;
using _AcCm = Teigha.Colors;
using _AcDb = Teigha.DatabaseServices;
using _AcEd = Bricscad.EditorInput;
using _AcGe = Teigha.Geometry;
using _AcGi = Teigha.GraphicsInterface;
using _AcGs = Teigha.GraphicsSystem;
using _AcPl = Bricscad.PlottingServices;
using _AcBrx = Bricscad.Runtime;
using _AcTrx = Teigha.Runtime;
using _AcWnd = Bricscad.Windows;
#elif ARX_APP
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using _AcAp = Autodesk.AutoCAD.ApplicationServices;
using _AcBr = Autodesk.AutoCAD.BoundaryRepresentation;
using _AcCm = Autodesk.AutoCAD.Colors;
using _AcDb = Autodesk.AutoCAD.DatabaseServices;
using _AcEd = Autodesk.AutoCAD.EditorInput;
using _AcGe = Autodesk.AutoCAD.Geometry;
using _AcGi = Autodesk.AutoCAD.GraphicsInterface;
using _AcGs = Autodesk.AutoCAD.GraphicsSystem;
using _AcPl = Autodesk.AutoCAD.PlottingServices;
using _AcBrx = Autodesk.AutoCAD.Runtime;
using _AcTrx = Autodesk.AutoCAD.Runtime;
using _AcWnd = Autodesk.AutoCAD.Windows;
using _AcLm = Autodesk.AutoCAD.LayerManager;
using System.Globalization;
using InterfacesPari;
using FactoryPari;
using System.IO;
#endif

namespace AcadPari
{
    public class Exporter
    {
        #region log4net Initialization
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(Exporter));
        static Exporter()
        {
            if (log4net.LogManager.GetRepository(System.Reflection.Assembly.GetExecutingAssembly()).Configured == false)
            {
                log4net.Config.XmlConfigurator.ConfigureAndWatch(
                    new FileInfo(
                        Path.Combine(
                            // ReSharper disable once AssignNullToNotNullAttribute
                            new FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).DirectoryName,
                            "_log4net.config"
                        )
                    )
                );
            }
        }
        #endregion

        [_AcTrx.CommandMethod("PariAutoFill")]
        public void PariAutoFill()
        {
            log.Info("PariAutoFill");
            try
            {

                var filler = new AttributeFiller();
                filler.Fill();
                var doc = Application.DocumentManager.MdiActiveDocument;
                var editor = doc.Editor;

                editor.WriteMessage("\n\nFill beendet.");
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }


        [_AcTrx.CommandMethod("ConfigPari")]
        public void ConfigPari()
        {
            log.Info("ConfigPari");
            try
            {
                var factory = new Factory();
                var database = factory.CreatePariDatabase();
                if (database == null) throw new InvalidOperationException("Database is null!");

                _AcWnd.OpenFileDialog ofd = new _AcWnd.OpenFileDialog("Datenbank wählen", "", "accdb", "AccessFile", _AcWnd.OpenFileDialog.OpenFileDialogFlags.AllowAnyExtension);
                var dr = ofd.ShowDialog();
                if (dr != DialogResult.OK) return;
                database.SetDatabase(ofd.Filename);

                ofd = new _AcWnd.OpenFileDialog("Excel-Template wählen", "", "xlsx", "ExcelFile", _AcWnd.OpenFileDialog.OpenFileDialogFlags.AllowAnyExtension);
                dr = ofd.ShowDialog();
                if (dr != DialogResult.OK) return;
                var excel = factory.CreateVisualOutputHandler();
                excel.SetTemplates(Path.GetDirectoryName(ofd.Filename));
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        [_AcTrx.CommandMethod("ExportPari")]
        public void ExportPari()
        {
            log.Info("ExportPari");
            try
            {

                var factory = new Factory();
                IPariDatabase database = factory.CreatePariDatabase();

                if (!CheckTableValidity(database)) return;

                var blockReader = new BlockReader();
                blockReader.ReadBlocksFromModelspace();
                var blockInfos = blockReader.BlockInfos;
                var wohnungInfos = blockReader.WohnungInfos;
                var projektInfo = blockReader.ProjektInfo;
                if (projektInfo == null)
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "Der ProjektInfo-Block existiert nicht in der Zeichnung!");
                    log.Error(msg);
                    Application.ShowAlertDialog(msg);
                    return;
                }
                var dwgName = Application.GetSystemVariable("DwgName").ToString();
                var dwgPrefix = Application.GetSystemVariable("DwgPrefix").ToString();
                projektInfo.DwgName = dwgName;
                projektInfo.DwgPrefix = dwgPrefix;
                var projektId = database.GetProjektId(projektInfo);
                if (projektId >= 0)
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "Das Projekt '{0}' existiert bereits! ProjektId = {1}.", projektInfo.Bauvorhaben, projektId);
                    log.Error(msg);
                    Application.ShowAlertDialog(msg);
                    return;
                }

                var tableBuilder = new TableBuilder();
                tableBuilder.Build(blockInfos, wohnungInfos);

                if (CheckInvalidCategoriesAskUser(tableBuilder)) return;


                var doc = Application.DocumentManager.MdiActiveDocument;
                var editor = doc.Editor;

                database.SaveToDatabase(tableBuilder, projektInfo);

                var raumWithoutTop = tableBuilder.RaumTable.Where(x => string.IsNullOrEmpty(x.Top)).ToList();
                if (raumWithoutTop.Count > 0)
                {
                    editor.WriteMessage("\nFolgende Räume haben keine TOP-Information: ");
                    editor.WriteMessage("\n-----------------------------------------------------------------");
                    foreach (var ri in raumWithoutTop)
                    {
                        var msg = string.Format(CultureInfo.CurrentCulture, "\nRaum: {0}\tGeschoss: {1}\tNutzwert: {2}\tHandle: {3}", ri.Raum, ri.Lage, ri.RNW, ri.AcadHandle);
                        log.Warn(msg);
                        editor.WriteMessage(msg);
                    }
                }

                editor.WriteMessage("\n\nExport beendet.");
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        private static bool CheckTableValidity(IPariDatabase database)
        {
            List<string> tableNames;
            try
            {
                tableNames = database.GetTableNames();
            }
            catch (Exception ex)
            {
                var msg = string.Format(CultureInfo.CurrentCulture, "Fehler beim Lesen der Datenbank! {0}", ex.Message);
                log.Error(msg);
                Application.ShowAlertDialog(msg);
                return false;
            }

            if (!tableNames.Contains("Projekt") || !tableNames.Contains("Raum") || !tableNames.Contains("Kategorie") ||
                !tableNames.Contains("GstInfo") || !tableNames.Contains("Wohnung"))
            {
                var msg = string.Format(CultureInfo.CurrentCulture, "Die Datenbank ist ungültig!");
                log.Error(msg);
                Application.ShowAlertDialog(msg);
                return false;
            }

            return true;
        }

        [_AcTrx.CommandMethod("UpdatePariProj")]
        public void UpdatePariProj()
        {
            log.Info("\nUpdatePariProj");
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                var editor = doc.Editor;

                var lg = new Logger(editor, log);
                lg.Info("Starte Update...");

                var factory = new Factory();
                IPariDatabase database = factory.CreatePariDatabase();

                if (!CheckTableValidity(database)) return;

                var blockReader = new BlockReader();
                blockReader.ReadBlocksFromModelspace();
                var blockInfos = blockReader.BlockInfos;
                var wohnungInfos = blockReader.WohnungInfos;
                if (CheckBlockWohnungConsistencyAskUser(blockInfos,wohnungInfos)) return;
                var projektInfo = blockReader.ProjektInfo;
                if (projektInfo == null)
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "Der ProjektInfo-Block existiert nicht in der Zeichnung!");
                    log.Error(msg);
                    Application.ShowAlertDialog(msg);
                    return;
                }

                lg.Info(string.Format(CultureInfo.CurrentCulture, "Projekt-ID: {0}", database.GetProjektId(projektInfo)));

                var tableUpdater = new TableUpdater();
                tableUpdater.Update(blockInfos, wohnungInfos, projektInfo, database);
                tableUpdater.LogStatus(lg);

                if (CheckInvalidCategoriesAskUser(tableUpdater)) return;

                database.UpdateDatabase(tableUpdater, projektInfo);

                lg.Info("\n\nUpdatebeendet.");
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        private bool CheckBlockWohnungConsistencyAskUser(List<IBlockInfo> blockInfos, List<IWohnungInfo> wohnungInfos)
        {
            var raumTops = blockInfos.GroupBy(x => x.Top).Select(x => x.First().Top).ToArray();
            var wohnungTops = wohnungInfos.GroupBy(x => x.Top).Select(x => x.First().Top).ToArray();
            var msgLst = new List<string>();

            //var raumTopsNotInWohnungTops = raumTops.Where(x => !wohnungTops.Contains(x)).ToArray();
            //if (raumTopsNotInWohnungTops.Any())
            //{
            //    var missingTops = string.Join(", ", raumTopsNotInWohnungTops);
            //    var msg = string.Format(CultureInfo.CurrentCulture,
            //        "Folgende Tops in den Raumblöcken existieren nicht in den Wohnungsblöcken: {0}", missingTops);
            //    msgLst.Add(msg);
            //}

            var wohnungTopsNotInRaumblocks = wohnungTops.Where(x => !raumTops.Contains(x)).ToArray();
            if (wohnungTopsNotInRaumblocks.Any())
            {
                var missingTops = string.Join(", ", wohnungTopsNotInRaumblocks);
                var msg = string.Format(CultureInfo.CurrentCulture,
                    "Folgende Tops in den Wohnungsblöcken existieren nicht in den Raumblöcken: {0}", missingTops);
                msgLst.Add(msg);
            }

            var theMsg = string.Join("\n", msgLst.ToArray());
            if (string.IsNullOrEmpty(theMsg)) return false;

            var res = MessageBox.Show(theMsg + "\nSoll trotzdem exportiert werden?", "UpdatePariProj", MessageBoxButtons.YesNo);
            if (res != DialogResult.Yes)
            {
                log.Warn("Abbruch durch Benutzer.");
                return true;
            }

            return false;
        }

        [_AcTrx.CommandMethod("PariLog")]
        public void PariLog()
        {
            log.Info("PariLog");
            try
            {
                var pariLogFileName = Path.Combine(Path.GetTempPath(), Globs.LogfileName);
                if (File.Exists(pariLogFileName))
                {
                    System.Diagnostics.Process.Start(pariLogFileName);
                }
                else
                {
                    var doc = Application.DocumentManager.MdiActiveDocument;
                    var editor = doc.Editor;
                    editor.WriteMessage(string.Format(CultureInfo.CurrentCulture, "\nDatei '{0}' nicht gefunden!", pariLogFileName));
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        private static bool CheckInvalidCategoriesAskUser(TableHandler tableUpdater)
        {
            if (tableUpdater.HasInvalidCategories)
            {
                var msg = string.Format(
                    "Es gibt Kategorien mit abweichenden Nutzwerten. {0}.\nSoll trotzdem exportiert werden?",
                    tableUpdater.JoinInvalidCatNames());
                var res = MessageBox.Show(msg, "UpdatePariProj", MessageBoxButtons.YesNo);
                if (res != DialogResult.Yes)
                {
                    log.Warn("Abbruch durch Benutzer.");
                    return true;
                }
            }

            return false;
        }

        [_AcTrx.CommandMethod("DelPariProj")]
        public void DelPariProj()
        {
            log.Info("DelPariProj");
            try
            {

                var doc = Application.DocumentManager.MdiActiveDocument;

                _AcEd.PromptIntegerOptions inOpts = new _AcEd.PromptIntegerOptions("\nId des Projekts, das gelöscht werden soll: ");
                inOpts.AllowNegative = false;
                inOpts.AllowNone = false;
                var intRes = doc.Editor.GetInteger(inOpts);
                if (intRes.Status != _AcEd.PromptStatus.OK) return;
                int projektId = intRes.Value;

                var factory = new Factory();
                var database = factory.CreatePariDatabase();
                int nrOfDeletedRows = database.DeleteProjekt(projektId);
                if (nrOfDeletedRows == 0)
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "\nEs wurde kein Projekt mit der Id {0} gefunden.", projektId);
                    log.Warn(msg);
                    doc.Editor.WriteMessage(msg);
                }
                else
                {
                    var msg = string.Format(CultureInfo.CurrentCulture, "Anzahl gelöschter Datensätze für ProjektId {1}: {0}", nrOfDeletedRows, projektId);
                    log.Info(msg);
                    doc.Editor.WriteMessage(msg);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        [_AcTrx.CommandMethod("KonsistenzPari")]
        public void KonsistenzPari()
        {
            log.Info("\n\nKonsistenzPari: Konsistenzprüfung wird gestartet:");

            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;

                var factory = new Factory();
                var database = factory.CreatePariDatabase();
                var nrOfAddedFiels = database.CheckExistingFields();
                if (nrOfAddedFiels > 0)
                {
                    doc.Editor.WriteMessage(string.Format("\nAnzahl hinzugefügter Felder: {0}", nrOfAddedFiels));
                }
                if (database.CheckConsistency())
                {
                    doc.Editor.WriteMessage("\nKeine Konsistenzfehler in der Datenkbank.");
                }
                else
                {
                    doc.Editor.WriteMessage("\nKonsistenzfehler in der Datenkbank! Siehe " +  Globs.LogfileName + " (Befehl Parilog)");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        [_AcTrx.CommandMethod("ListPariProj")]
        public void ListPariProj()
        {
            log.Info("ListPariProj");
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                var editor = doc.Editor;

                var factory = new Factory();
                var database = factory.CreatePariDatabase();
                var projInfos = database.ListProjInfos();

                editor.WriteMessage(string.Format(CultureInfo.CurrentCulture, "Datenbank: {0}", database));
                editor.WriteMessage("\nProjekte: ");
                editor.WriteMessage("\n-----------------------------------------------------------------");
                foreach (var pi in projInfos)
                {
                    editor.WriteMessage(string.Format(CultureInfo.CurrentCulture, "\nID: {0}\tName: {1}\tDwgName: {2}\tEZ: {3}", pi.ProjektId, pi.Bauvorhaben, pi.DwgName, pi.EZ));
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }

        [_AcTrx.CommandMethod("ExcelPari")]
        public void ExcelPari()
        {
            log.Info("ExcelPari");
            try
            {
                var fact = new Factory();
                var excelExporter = fact.CreateVisualOutputHandler();
                var database = fact.CreatePariDatabase();

                var dwgPrefix = Application.GetSystemVariable("DwgPrefix").ToString();
                var dwgName = Application.GetSystemVariable("DwgName").ToString();

                int projektId = -1;
                var doc = Application.DocumentManager.MdiActiveDocument;
                _AcEd.PromptIntegerOptions inOpts = new _AcEd.PromptIntegerOptions("\nId des Projekts, für das eine Exceldatei erstellt werden soll <Return für aktuelles>: ");
                inOpts.AllowNegative = false;
                inOpts.AllowNone = true;
                var intRes = doc.Editor.GetInteger(inOpts);
                if (intRes.Status == _AcEd.PromptStatus.OK)
                {
                    projektId = intRes.Value;
                    var pi = database.ListProjInfos().FirstOrDefault(x => x.ProjektId == projektId);
                    if (pi == null)
                    {
                        var msg = string.Format(CultureInfo.CurrentCulture, "Das Projekt mit Id {0} existiert nicht in der Datenbank!", projektId);
                        log.Error(msg);
                        Application.ShowAlertDialog(msg);
                        return;
                    }
                }
                else if (intRes.Status == _AcEd.PromptStatus.None)
                {
                    var blockReader = new BlockReader();
                    blockReader.ReadBlocksFromModelspace();
                    var projektInfo = blockReader.ProjektInfo;
                    if (projektInfo == null)
                    {
                        var msg = string.Format(CultureInfo.CurrentCulture, "Der ProjektInfo-Block existiert nicht in der Zeichnung!");
                        log.Error(msg);
                        Application.ShowAlertDialog(msg);
                        return;
                    }

                    projektInfo.DwgName = dwgName;
                    projektInfo.DwgPrefix = dwgPrefix;

                    projektId = database.GetProjektId(projektInfo);
                    if (projektId < 0)
                    {
                        var msg = string.Format(CultureInfo.CurrentCulture, "Das Projekt wurde nicht gefunden in der Datenbank!");
                        log.Error(msg);
                        Application.ShowAlertDialog(msg);
                        return;
                    }
                }

                doc.Editor.WriteMessage("\nExport wird gestartet...");
                //var targetFile = Path.Combine(dwgPrefix, Path.GetFileNameWithoutExtension(dwgName) + "_NW" + ".xlsx");
                excelExporter.ExportNW(database, null, projektId);
                excelExporter.ExportNF(database, null, projektId);
                doc.Editor.WriteMessage("\nExport fertiggestellt.");
            }
            catch (Exception ex)
            {
                log.Error(ex.Message, ex);
                Application.ShowAlertDialog(ex.Message);
            }
        }
    }
}
