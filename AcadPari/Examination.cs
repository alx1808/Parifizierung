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
#endif

using System;
using System.Globalization;


namespace AcadPari
{
    public class Examination
    {
        [_AcTrx.CommandMethod("PariKommaPruef", _AcTrx.CommandFlags.UsePickSet | _AcTrx.CommandFlags.Redraw | _AcTrx.CommandFlags.Modal)]
        public static void PariKommaPruef()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            try
            {
                ClearImpliedSelection(ed);

                var blockReader = new BlockReader();
                if (!CheckInvalidKomma(Globs.RaumBlockName, "Raumblöcke", HasRaumblockInvalidKomma, blockReader, ed))
                {
                    CheckInvalidKomma(Globs.WohnunginfoBlockName, "Wohnungsblöcke", HasWohnungblockInvalidKomma, blockReader, ed);
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage(ex.Message);
            }
        }

        private static void ClearImpliedSelection(_AcEd.Editor ed)
        {
            ed.SetImpliedSelection(new _AcDb.ObjectId[0]);
        }

        private static bool CheckInvalidKomma(string blockName,
            string blockNameStringPlural, Func<_AcDb.BlockReference, bool> checkFunc, BlockReader blockReader, _AcEd.Editor ed)
        {
            var invalidRaumblocks = blockReader.GetAllBlocksInModelSpaceWith(checkFunc, blockName);
            if (invalidRaumblocks.Count > 0)
            {
                ed.WriteMessage(string.Format(CultureInfo.CurrentCulture, "\nAnzahl gefundener {1} mit Komma: {0}",
                    invalidRaumblocks.Count, blockNameStringPlural));
                ed.SetImpliedSelection(invalidRaumblocks.ToArray());
                return true;
            }
            ed.WriteMessage(string.Format(CultureInfo.CurrentCulture, "\nEs wurden keine {0} mit Komma gefunden.",
                blockNameStringPlural));
            return false;
        }

        private static bool HasRaumblockInvalidKomma(_AcDb.BlockReference blockRef)
        {
            var attDict = Globs.GetAttributesUcKey(blockRef);
            var flaeche = Globs.GetValOrEmpty("FLÄCHE", attDict);
            if (flaeche.IndexOf(".", StringComparison.Ordinal) >= 0) return true;
            var nutzwert = Globs.GetValOrEmpty("NUTZWERT", attDict);
            if (nutzwert.IndexOf(".", StringComparison.Ordinal) >= 0) return true;
            return false;
        }

        private static bool HasWohnungblockInvalidKomma(_AcDb.BlockReference blockRef)
        {
            var attDict = Globs.GetAttributesUcKey(blockRef);
            var nutzwert = Globs.GetValOrEmpty("NUTZWERT", attDict);
            if (nutzwert.IndexOf(".", StringComparison.Ordinal) >= 0) return true;
            return false;
        }
    }
}
