using System;
using System.Globalization;
using InterfacesPari;

namespace FactoryPari
{
    public class RaumRecord : IRaumRecord
    {
        public int RaumId { get; set; }
        public string Top { get; set; }
        public string Lage { get; set; }
        public string Raum { get; set; }
        public string Widmung { get; set; }
        public string RNW { get; set; }
        public string Begrundung { get; set; }
        public double Nutzwert { get; set; }
        public double Flaeche { get; set; }
        public int ProjektId { get; set; }
        public int KategorieId { get; set; }
        public string AcadHandle { get; set; }
        public IKategorieRecord Kategorie { get; set; }

        public bool IsEqualTo(IRaumRecord otherRaumRecord)
        {
            if (Top != otherRaumRecord.Top) return false;
            if (Lage != otherRaumRecord.Lage) return false;
            if (Raum != otherRaumRecord.Raum) return false;
            if (Widmung != otherRaumRecord.Widmung) return false;
            if (RNW != otherRaumRecord.RNW) return false;
            if (Begrundung != otherRaumRecord.Begrundung) return false;
            if (!DblEquals(Nutzwert, otherRaumRecord.Nutzwert)) return false;
            if (!DblEquals(Flaeche, otherRaumRecord.Flaeche)) return false;

            return true;
        }

        private bool DblEquals(double d1, double d2)
        {
            return Math.Abs(d1 - d2) < 0.00001;
        }

        public string KatIdentification
        {
            get
            {
                return string.Join("|", new string[] { Top, Lage, Widmung, Begrundung });
            }
        }

        public void UpdateValuesFrom(IBlockInfo acadBlockInfo)
        {
            Top = acadBlockInfo.Top;
            Lage = acadBlockInfo.Geschoss;
            Raum = acadBlockInfo.Raum;
            RNW = acadBlockInfo.Nutzwert.Trim();
            Begrundung = acadBlockInfo.Begrundung;
            Widmung = acadBlockInfo.Widmung;
            if (string.IsNullOrEmpty(RNW))
            {
                // todo: rnw darf nicht mehr leer sein. wenn rnw nicht vorhanden ist, dann kommt er aus dem top-block der wohnung
                // ALLG – Flächen haben keinen Wohnungsblock. Woher kommt der Nutzwert? Bisher wurde dieser immer auf 1.0 gesetzt.
                //throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Kein Nutzwert {2} in {0}, Top {1}.", acadBlockInfo.Raum, acadBlockInfo.Top, acadBlockInfo.Nutzwert));
                Nutzwert = 1.0;
            }
            else
            {
                double nutzwert;
                var rnw = RNW.Replace(',', '.');
                if (!double.TryParse(rnw, NumberStyles.Any, CultureInfo.CurrentCulture, out nutzwert))
                {
                    throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Ungültiger Nutzwert {2} in {0}, Top {1}.", acadBlockInfo.Raum, acadBlockInfo.Top, acadBlockInfo.Nutzwert));
                }
                Nutzwert = nutzwert;
            }
            var m2S = GetM2(acadBlockInfo);
            double m2;
            if (!double.TryParse(m2S, out m2))
            {
                throw new InvalidOperationException(string.Format(CultureInfo.CurrentCulture, "Ungültige Fläche {2} in {0}, Top {1}.", acadBlockInfo.Raum, acadBlockInfo.Top, acadBlockInfo.Flaeche));
            }

            //WidmungAndBegrundungCorrection();

            Flaeche = m2;
            AcadHandle = acadBlockInfo.Handle;
        }

        private static string GetM2(IBlockInfo bi)
        {
            var m2S = bi.Flaeche.Trim().Replace(".", "").Replace(',', '.');
            int iom = m2S.IndexOf('m');
            if (iom >= 0)
            {
                m2S = m2S.Substring(0, iom);
            }
            m2S = m2S.Trim();
            return m2S;
        }

        #region Widmung and Begrundung Correction

        /// <summary>
        /// Nach Einlesen der Information aus den AutoCAD-Blöcken werden hier die Werte für Widmung und Begrundung festgelegt.
        /// </summary>
        /// <remarks>
        /// wichtig: Zuerst SetWidmungFromOtherValues, da danach Begrundung in FixBegrundungValue geändert wird, das von SetWidmungFromOtherValues verwendet wird.
        /// </remarks>
        [Obsolete("Wird nun in BlockReader.WidmungNutzwertBegrundungCorrection gemacht")]
        private void WidmungAndBegrundungCorrection()
        {
            
            SetWidmungFromOtherValues();
            FixBegrundungValue();
        }

        private void FixBegrundungValue()
        {
            if (Begrundung == null || string.IsNullOrEmpty(Begrundung.Trim()) || Begrundung.ToUpper().Contains("PKW"))
            {
                Begrundung = "als Wohnungseigentumsobjekt";
            }
        }

        private void SetWidmungFromOtherValues()
        {
            if (string.IsNullOrEmpty(RNW))
            {
                Widmung = "Wohnung";
            }
            else if (Begrundung.ToUpper().Contains("PKW"))
            {
                Widmung = Begrundung;
            }
            else
            {
                Widmung = Raum;
            }
        }

        #endregion

        public override string ToString()
        {
            return string.Format(CultureInfo.CurrentCulture,
                "ID: {0}, Top: {1}, Lage: {2}, Widmung: {3}, Begrundung: {4}, Nutzwert: {5}, ProjektID: {6}", RaumId, Top, Lage,
                Widmung, Begrundung, Nutzwert, ProjektId);
        }

        public IRaumRecord ShallowCopy()
        {
            return (IRaumRecord) MemberwiseClone();
        }
    }
}

