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
    public class AttributeFiller
    {
        private class BlockInfo
        {
            public BlockInfo()
            {
                Attributes = new Dictionary<string, _AcDb.AttributeReference>();
            }

            public _AcDb.BlockReference BlockRef { get; set; }

            public Dictionary<string, _AcDb.AttributeReference> Attributes { get; set; }
        }

        private List<BlockInfo> _raumBlockInfos = new List<BlockInfo>();
        private List<BlockInfo> _wohnungBlockInfos = new List<BlockInfo>();
        private BlockInfo _projektBlockInfo;

        const string attTop = "TOP";
        const string attBegrundung = "BEGRUNDUNG";
        const string attRaum = "RAUM";
        private const string attWohnWidmung = "WIDMUNG";

        const string attExcelBegrundung = "EXCEL_BEGRUNDUNG";
        const string attNW_PariAls = "NW_PARIALS";
        const string attNF_SummeBez = "NF_SUMMEBEZ";
        const string attNW_Widmung = "NW_WIDMUNG";
        const string attNF_SummeWidmung = "NF_SUMMEWIDMUNG";
        const string attNF_AllgemeinGruppe = "NF_ALLGEMEINGRUPPE";
        const string attIsZuschlag = "ISZUSCHLAG";
        const string attIsZubehoer = "ISZUBEHÖR";
        const string attNF_SpecialHandling = "NF_SPECIALHANDLING";
        const string attNF_SpecialZuschlagBez = "NF_SPECIALZUSCHLAGBEZ";


        public void Fill()
        {
            _raumBlockInfos.Clear();
            _wohnungBlockInfos.Clear();
            _projektBlockInfo = null;

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
                            var attDict = Globs.GetAttributes(blockRef, tr);
                            var blockInfo = new BlockInfo
                            {
                                BlockRef = blockRef,
                                Attributes = attDict,
                            };
                            _raumBlockInfos.Add(blockInfo);
                        }
                        else if (string.Compare(blockRef.Name, Globs.GstinfoBlockName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            //if (_projektInfo == null) _projektInfo = _factory.CreateProjectInfo(); // new AcadPari.ProjektInfo();
                            //var attDict = Globs.GetAttributesUcKey(blockRef);
                            //_projektInfo.Bauvorhaben = Globs.GetValOrEmpty("BAUVORHABEN", attDict);
                            //_projektInfo.Katastralgemeinde = Globs.GetValOrEmpty("KATASTRALGEMEINDE", attDict);
                            //_projektInfo.EZ = Globs.GetValOrEmpty("EZ", attDict);

                            //var subInfo = _factory.CreateSubInfo(); // new AcadPari.ProjektInfo.SubInfo();
                            //subInfo.Gstnr = Globs.GetValOrEmpty("GRUNDSTÜCKSNUMMER", attDict);
                            //subInfo.Flaeche = Globs.GetValOrEmpty("FLACHE", attDict);
                            //subInfo.AcadHandle = blockRef.Handle.ToString();
                            //_projektInfo.SubInfos.Add(subInfo);
                        }
                        else if (string.Compare(blockRef.Name, Globs.WohnunginfoBlockName, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            var attDict = Globs.GetAttributes(blockRef, tr);
                            var wohnungInfo = new BlockInfo()
                            {
                                BlockRef = blockRef,
                                Attributes = attDict,
                            };
                            _wohnungBlockInfos.Add(wohnungInfo);
                        }
                    }
                }

                // PKV Wohnungen
                var pkwWohnungsInfos = new List<BlockInfo>();
                foreach (var raumBlockInfo in _raumBlockInfos)
                {
                    var begrundung = raumBlockInfo.Attributes[attBegrundung];
                    var valBegrundung = begrundung.TextString;
                    var valBegrundungUc = valBegrundung.ToUpperInvariant();

                    var top = raumBlockInfo.Attributes[attTop];
                    var valTop = top.TextString;
                    var valTopUc = valTop.ToUpperInvariant();

                    var raum = raumBlockInfo.Attributes[attRaum];
                    var valRaum = raum.TextString;
                    var valRaumUc = valRaum.ToUpperInvariant();

                    if (valBegrundungUc.Contains("PKW"))
                    {
                        var wohnungsInfoBlock = GetWohnungsInfoBlock(valTopUc);
                        if (wohnungsInfoBlock == null)
                        {
                            throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Kein Wohnungsinfoblock für Raumblock '{2}'('{1}','{0}')!", raumBlockInfo.BlockRef.Handle.ToString(), valTop, valRaum));
                        }
                        if (!pkwWohnungsInfos.Contains(wohnungsInfoBlock)) pkwWohnungsInfos.Add(wohnungsInfoBlock);
                    }
                }


                foreach (var raumBlockInfo in _raumBlockInfos)
                {
                    var bIsZuschlag = false;
                    var bIsZubehör = false;

                    var begrundung = raumBlockInfo.Attributes[attBegrundung];
                    var valBegrundung = begrundung.TextString;
                    var valBegrundungUc = valBegrundung.ToUpperInvariant();

                    var top = raumBlockInfo.Attributes[attTop];
                    var valTop = top.TextString;
                    var valTopUc = valTop.ToUpperInvariant();

                    var raum = raumBlockInfo.Attributes[attRaum];
                    var valRaum = raum.TextString;
                    var valRaumUc = valRaum.ToUpperInvariant();

                    var wohnungsInfoBlock = GetWohnungsInfoBlock(valTopUc);
                    if (wohnungsInfoBlock == null)
                    {
                        throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Kein Wohnungsinfoblock für Raumblock '{2}'('{1}','{0}')!", raumBlockInfo.BlockRef.Handle.ToString(), valTop, valRaum));
                    }


                    // Excel_Begrundung
                    var excelBegrundung = raumBlockInfo.Attributes[attExcelBegrundung];
                    excelBegrundung.UpgradeOpen();
                    if (string.IsNullOrEmpty(valBegrundung.Trim()) || valBegrundungUc.Contains("ALLG") ||
                        valBegrundungUc.Contains("PKW"))
                    {
                        excelBegrundung.TextString = "als Wohnungseigentumsobjekt";
                    }
                    else
                    {
                        excelBegrundung.TextString = valBegrundung;
                    }
                    excelBegrundung.DowngradeOpen();


                    // NW_PariAls
                    var nwPariAls = raumBlockInfo.Attributes[attNW_PariAls];
                    nwPariAls.UpgradeOpen();
                    if (valBegrundungUc.Contains("ALS ZUBEHÖR"))
                    {
                        nwPariAls.TextString = "Als Wohnungseigentumszubehör";
                    }
                    else if (valBegrundungUc.Contains("ALS ZUSCHLAG"))
                    {
                        nwPariAls.TextString = "Als Wohnungseigentumszuschlag";
                    }
                    else
                    {
                        nwPariAls.TextString = "Als Wohnungseigentumsobjekt";
                    }
                    nwPariAls.DowngradeOpen();

                    // IsZuschlag
                    var isZuschlag = raumBlockInfo.Attributes[attIsZuschlag];
                    isZuschlag.UpgradeOpen();
                    if (valBegrundungUc.Contains("ALS ZUSCHLAG"))
                    {
                        bIsZuschlag = true;
                        isZuschlag.TextString = "x";
                    }
                    else
                    {
                        isZuschlag.TextString = "";
                    }

                    // IsZubehör
                    var isZubehoer = raumBlockInfo.Attributes[attIsZubehoer];
                    isZubehoer.UpgradeOpen();
                    if (valBegrundungUc.Contains("ALS ZUBEHÖR"))
                    {
                        bIsZubehör = true;
                        isZubehoer.TextString = "x";
                    }
                    else
                    {
                        isZubehoer.TextString = "";
                    }
                    isZubehoer.DowngradeOpen();

                    // NF_SummeBez, NF_SummeWidmung
                    var nfSummeBez = raumBlockInfo.Attributes[attNF_SummeBez];
                    nfSummeBez.UpgradeOpen();
                    var nfSummeWidmung = raumBlockInfo.Attributes[attNF_SummeWidmung];
                    nfSummeWidmung.UpgradeOpen();
                    if (pkwWohnungsInfos.Contains(wohnungsInfoBlock))
                    {
                        nfSummeBez.TextString = "PKW - ABSTELLPLÄTZE";
                        nfSummeWidmung.TextString = "";
                    }
                    else if (valTopUc.Contains("ALLG"))
                    {
                        nfSummeBez.TextString = "Allgemeinflächen";
                        nfSummeWidmung.TextString = "";
                    }
                    else
                    {
                        nfSummeBez.TextString = valTop;
                        nfSummeWidmung.TextString = wohnungsInfoBlock.Attributes[attWohnWidmung].TextString;
                    }
                    nfSummeWidmung.DowngradeOpen();
                    nfSummeBez.DowngradeOpen();

                    // NF_AllgemeinGruppe
                    var nfAllgemeinGruppe = raumBlockInfo.Attributes[attNF_AllgemeinGruppe];
                    nfAllgemeinGruppe.UpgradeOpen();
                    if (valTopUc.Trim() == "TOP ALLG")
                    {
                        nfAllgemeinGruppe.TextString = "TOP";
                    }
                    else if (valTopUc.Trim() == "ALLG")
                    {
                        nfAllgemeinGruppe.TextString = "AUSSENANLAGEN";
                    }
                    nfAllgemeinGruppe.DowngradeOpen();

                    // NF_SpecialHandling
                    // NF_SpecialZuschlagBez
                    var specialHandling = raumBlockInfo.Attributes[attNF_SpecialHandling];
                    specialHandling.UpgradeOpen();
                    var specZuschlagBez = raumBlockInfo.Attributes[attNF_SpecialZuschlagBez];
                    specZuschlagBez.UpgradeOpen();
                    if (pkwWohnungsInfos.Contains(wohnungsInfoBlock))
                    {
                        specialHandling.TextString = "Abstellplätze";
                        if (bIsZuschlag)
                        {
                            specZuschlagBez.TextString = valRaum + " " + valBegrundung;
                        }
                        else
                        {
                            specZuschlagBez.TextString = "";
                        }
                    }
                    else
                    {
                        //specialHandling.TextString = "";
                        //specZuschlagBez.TextString = "";
                    }
                    specialHandling.DowngradeOpen();
                    specZuschlagBez.DowngradeOpen();


                    // NW_Widmung
                    var nwWidmung = raumBlockInfo.Attributes[attNW_Widmung];
                    nwWidmung.UpgradeOpen();
                    if (valTopUc.Contains("ALLG"))
                    {
                        nwWidmung.TextString = "";
                    }
                    else if (bIsZuschlag || bIsZubehör)
                    {
                        nwWidmung.TextString = valRaum;
                    }
                    else
                    {
                        nwWidmung.TextString = wohnungsInfoBlock.Attributes[attWohnWidmung].TextString;
                    }
                    nwWidmung.DowngradeOpen();
                }

                tr.Commit();
            }

            _raumBlockInfos.Clear();
            _wohnungBlockInfos.Clear();
            _projektBlockInfo = null;
        }

        private BlockInfo GetWohnungsInfoBlock(string valTopUc)
        {
            return _wohnungBlockInfos.FirstOrDefault(x => x.Attributes[attTop].TextString.ToUpper() == valTopUc);
        }
    }
}
