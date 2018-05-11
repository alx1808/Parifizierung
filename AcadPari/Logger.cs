using Autodesk.AutoCAD.ApplicationServices.Core;
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
using AcadPari.Properties;
using System.Globalization;
using InterfacesPari;
using FactoryPari;
using System.IO;
#endif

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AcadPari
{
    internal class Logger
    {
        private _AcEd.Editor _editor;
        private log4net.ILog _log;

        public Logger(_AcEd.Editor editor,log4net.ILog log)
        {
            _editor = editor;
            _log = log;
        }
        public void Info(string msg)
        {
            _log.Info(msg);
            _editor.WriteMessage("\n" + msg);
        }
        public void Warn(string msg)
        {
            _log.Warn(msg);
            _editor.WriteMessage("\nWarnung:" + msg);
        }
        public void Error(string msg)
        {
            _log.Error(msg);
            _editor.WriteMessage("\nFehler:" + msg);
        }
    }
}
