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
using InterfacesPari;
#endif

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices.Core;


namespace AcadPari
{
    public class BlockReader
    {
        private readonly FactoryPari.Factory _factory;
        public BlockReader()
        {
            _factory = new FactoryPari.Factory();
        }

        private readonly List<IBlockInfo> _blockInfos = new List<IBlockInfo>();
        public List<IBlockInfo> BlockInfos
        {
            get { return _blockInfos; }
        }

        private readonly List<IWohnungInfo> _wohnungInfos = new List<IWohnungInfo>();
        public List<IWohnungInfo> WohnungInfos
        {
            get { return _wohnungInfos; }
        }

        private IProjektInfo _projektInfo;
        public IProjektInfo ProjektInfo
        {
            get { return _projektInfo; }
        }

        public List<_AcDb.ObjectId> GetAllBlocksInModelSpaceWith(Func<_AcDb.BlockReference, bool> checkFunc, string blockName)
        {
            var oids = new List<_AcDb.ObjectId>();
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (_AcDb.Transaction tr = doc.TransactionManager.StartTransaction())
            {
                _AcDb.BlockTable bt = (_AcDb.BlockTable)tr.GetObject(db.BlockTableId, _AcDb.OpenMode.ForRead);
                _AcDb.BlockTableRecord btr = (_AcDb.BlockTableRecord)tr.GetObject(bt[_AcDb.BlockTableRecord.ModelSpace], _AcDb.OpenMode.ForRead);

                foreach (var oid in btr)
                {
                    var blockRef = tr.GetObject(oid, _AcDb.OpenMode.ForRead) as _AcDb.BlockReference;
                    if (blockRef != null)
                    {
                        if (string.Compare(blockRef.Name, blockName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            if (checkFunc(blockRef))
                            {
                                oids.Add(oid);
                            }
                        }
                    }
                }
                tr.Commit();
            }

            return oids;
        }

        public void ReadBlocksFromModelspace()
        {
            _blockInfos.Clear();
            _wohnungInfos.Clear();
            _projektInfo = null;

            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (_AcDb.Transaction tr = doc.TransactionManager.StartTransaction())
            {
                _AcDb.BlockTable bt = (_AcDb.BlockTable)tr.GetObject(db.BlockTableId, _AcDb.OpenMode.ForRead);
                _AcDb.BlockTableRecord btr = (_AcDb.BlockTableRecord)tr.GetObject(bt[_AcDb.BlockTableRecord.ModelSpace], _AcDb.OpenMode.ForRead);

                foreach (var oid in btr)
                {
                    var blockRef = tr.GetObject(oid, _AcDb.OpenMode.ForRead) as _AcDb.BlockReference;
                    if (blockRef != null)
                    {
                        if (string.Compare(blockRef.Name, Globs.RaumBlockName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            var attDict = Globs.GetAttributesUcKey(blockRef);
                            var blockInfo = new BlockInfo
                            {
                                Raum = Globs.GetValOrEmpty("RAUM", attDict),
                                Flaeche = Globs.GetValOrEmpty("FLÄCHE", attDict),
                                Zusatz = Globs.GetValOrEmpty("ZUSATZ", attDict),
                                Top = Globs.GetValOrEmpty("TOP", attDict),
                                Geschoss = Globs.GetValOrEmpty("GESCHOSS", attDict),
                                Nutzwert = Globs.GetValOrEmpty("NUTZWERT", attDict),
                                Begrundung = Globs.GetValOrEmpty("BEGRUNDUNG", attDict),
                                Handle = blockRef.Handle.ToString()
                            };
                            _blockInfos.Add(blockInfo);
                        }
                        else if (string.Compare(blockRef.Name, Globs.GstinfoBlockName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            if (_projektInfo == null) _projektInfo = _factory.CreateProjectInfo(); // new AcadPari.ProjektInfo();
                            var attDict = Globs.GetAttributesUcKey(blockRef);
                            _projektInfo.Bauvorhaben = Globs.GetValOrEmpty("BAUVORHABEN", attDict);
                            _projektInfo.Katastralgemeinde = Globs.GetValOrEmpty("KATASTRALGEMEINDE", attDict);
                            _projektInfo.EZ = Globs.GetValOrEmpty("EZ", attDict);

                            var subInfo = _factory.CreateSubInfo(); // new AcadPari.ProjektInfo.SubInfo();
                            subInfo.Gstnr  = Globs.GetValOrEmpty("GRUNDSTÜCKSNUMMER", attDict);
                            subInfo.Flaeche = Globs.GetValOrEmpty("FLACHE", attDict);
                            subInfo.AcadHandle = blockRef.Handle.ToString();
                            _projektInfo.SubInfos.Add(subInfo);
                        }
                        else if (string.Compare(blockRef.Name, Globs.WohnunginfoBlockName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            var attDict = Globs.GetAttributesUcKey(blockRef);
                            var wohnungInfo = new WohnungInfo
                            {
                                Top = Globs.GetValOrEmpty("TOP", attDict),
                                Typ = Globs.GetValOrEmpty("TYP", attDict),
                                Widmung = Globs.GetValOrEmpty("WIDMUNG", attDict),
                                Nutzwert = Globs.GetValOrEmpty("NUTZWERT", attDict),
                            };
                            _wohnungInfos.Add(wohnungInfo);
                        }
                    }
                }
                tr.Commit();
            }

            WidmungNutzwertBegrundungCorrection();


        }

        private void FixBegrundungValue(IBlockInfo blockInfo)
        {
            if (blockInfo.Begrundung == null || string.IsNullOrEmpty(blockInfo.Begrundung.Trim()) || blockInfo.Begrundung.ToUpper().Contains("PKW"))
            {
                blockInfo.Begrundung = "als Wohnungseigentumsobjekt";
            }
        }

        private void WidmungNutzwertBegrundungCorrection()
        {
            var pkwTops = _blockInfos.Where(x => x.Begrundung.ToUpper().Contains("PKW")).Select(x => x.Top).ToList();

            foreach (var blockInfo in _blockInfos)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(blockInfo.Top, "ALLG", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    blockInfo.Widmung = "";
                    blockInfo.Nutzwert = "1,00";
                    blockInfo.Begrundung = "als Wohnungseigentumsobjekt";
                    continue;
                }

                if (blockInfo.Begrundung.ToUpper().Contains("PKW"))
                {
                    // pkw abstellflächen
                    blockInfo.Widmung = blockInfo.Begrundung;
                    continue;
                }

                if (pkwTops.Contains(blockInfo.Top))
                {
                    // zuschlag zu pkw abstellflächen
                    blockInfo.Widmung = blockInfo.Raum;
                    continue;
                }

                var wohnung = _wohnungInfos.FirstOrDefault(x => x.Top == blockInfo.Top);
                if (wohnung == null)
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Es wurde keine Wohnungsinfo für Top '{0}' gefunden!", blockInfo.Top));
                }

                if (!string.IsNullOrEmpty(blockInfo.Nutzwert))
                {
                    blockInfo.Widmung = blockInfo.Raum;
                }
                else
                {
                    blockInfo.Widmung = wohnung.Widmung;
                    blockInfo.Nutzwert = wohnung.Nutzwert;
                    if (string.IsNullOrEmpty(blockInfo.Widmung))
                    {
                        
                        throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Keine Widmung aus Wohnungsinfo für Raum '{0}' in Top '{1}'!",blockInfo.Raum, blockInfo.Top));
                    }
                }

                FixBegrundungValue(blockInfo);
            }
        }
    }
}
