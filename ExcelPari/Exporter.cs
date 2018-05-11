using ExcelPari.Properties;
using InterfacesPari;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelPari
{
    public class Exporter : IVisualOutput
    {
        public void ExportNW(IPariDatabase database, string locationHint, int projektId)
        {
            if (database == null) throw new ArgumentNullException("database");
            if (locationHint != null)
            {
                if (File.Exists(locationHint)) throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "File '{0}' already exist!", locationHint), "locationHint");
            }

            using (var excelizer = new ExcelizerNW(database, locationHint))
            {
                excelizer.ExportNW(projektId);
            }
        }

        public void ExportNF(IPariDatabase database, string locationHint, int projektId)
        {
            if (database == null) throw new ArgumentNullException("database");
            if (locationHint != null)
            {
                if (File.Exists(locationHint)) throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "File '{0}' already exist!", locationHint), "locationHint");
            }

            using (var excelizer = new ExcelizerNF(database, locationHint))
            {
                excelizer.ExportNf(projektId);
            }
        }

        public void SetTemplates(object oTemplate)
        {
            if (oTemplate == null) throw new ArgumentNullException(paramName: "oTemplate");
            var templateLocation = oTemplate.ToString();
            if (!System.IO.Directory.Exists(templateLocation))
            {
                throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Verzeichnis '{0}' nicht gefunden!", templateLocation));
            }

            Settings.Default.TemplateLocation = templateLocation;
            Settings.Default.Save();
        }
    }
}
