
using System;
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
#endif
using System.Collections.Generic;
using System.Globalization;
using Autodesk.AutoCAD.ApplicationServices.Core;

namespace AcadPari
{
    internal class Globs
    {
        public const string GstinfoBlockName = "Grundstücksinfo";
        public const string RaumBlockName = "Raumblock";
        public const string WohnunginfoBlockName = "Wohnungsinfo";
        public const string LogfileName = "Parifizierung.log";

        public static Dictionary<string, string> GetAttributesUcKey(_AcDb.BlockReference blockRef)
        {
            Dictionary<string, string> valuePerTag = new Dictionary<string, string>();

            _AcAp.Document doc = Application.DocumentManager.MdiActiveDocument;
            _AcDb.Database db = doc.Database;
            using (var trans = db.TransactionManager.StartTransaction())
            {

                foreach (_AcDb.ObjectId attId in blockRef.AttributeCollection)
                {
                    var anyAttRef = trans.GetObject(attId, _AcDb.OpenMode.ForRead) as _AcDb.AttributeReference;
                    if (anyAttRef != null)
                    {
                        valuePerTag[anyAttRef.Tag.ToUpperInvariant()] = anyAttRef.TextString;
                    }
                }
                trans.Commit();
            }
            return valuePerTag;
        }

        public static Dictionary<string, _AcDb.AttributeReference> GetAttributes(_AcDb.BlockReference blockRef, _AcDb.Transaction trans)
        {
            var valuePerTag = new Dictionary<string, _AcDb.AttributeReference>();

            foreach (_AcDb.ObjectId attId in blockRef.AttributeCollection)
            {
                var anyAttRef = trans.GetObject(attId, _AcDb.OpenMode.ForRead) as _AcDb.AttributeReference;
                if (anyAttRef != null)
                {
                    valuePerTag[anyAttRef.Tag.ToUpperInvariant()] = anyAttRef;
                }
            }
            return valuePerTag;
        }

        public static string GetValOrEmpty(string attName, Dictionary<string, string> attDict)
        {
            string val;
            if (attDict.TryGetValue(attName, out val)) return val;
            return "";
        }

    }
}
