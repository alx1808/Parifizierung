using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPari
{
    internal class ExcelMatrix
    {
        private int _NrOfCols;
        private int _StartRowIndex;
        private List<ExRow> _Rows = new List<ExRow>();
        public ExcelMatrix(int startRowIndex, int nrOfCols)
        {
            if (nrOfCols <= 0) throw new ArgumentNullException(string.Format(CultureInfo.InvariantCulture, "Invalid nrOrCols: {0}", nrOfCols), "nrOfCols");
            if (startRowIndex <= 0) throw new ArgumentNullException(string.Format(CultureInfo.InvariantCulture, "Invalid startRowIndex: {0}", nrOfCols), "startRowIndex");
            _NrOfCols = nrOfCols;
            _StartRowIndex = startRowIndex;
        }
        public void Add(int rowIndex, int colIndex, object o)
        {
            AddMissingRows(rowIndex);
            _Rows[rowIndex-_StartRowIndex].Arr[colIndex] = o;
        }

        private void AddMissingRows(int rowIndex)
        {
            for (int i = _Rows.Count+_StartRowIndex; i <= rowIndex; i++)
            {
                _Rows.Add(new ExRow(_NrOfCols));
            }
        }
        internal void Write(Microsoft.Office.Interop.Excel.Worksheet targetSheet)
        {
            object[,] indexMatrix = BuildMatrix();
            var b1 = GetCellBez(_StartRowIndex, 0);
            var b2 = GetCellBez(_StartRowIndex + _Rows.Count - 1, _NrOfCols - 1);
            var range = targetSheet.Range[b1, b2];
            range.set_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault, indexMatrix);
        }

        private object[,] BuildMatrix()
        {
            object[,] indexMatrix = new object[_Rows.Count, _NrOfCols];
            for (int rowCnt = 0; rowCnt < _Rows.Count; rowCnt++)
            {
                var r = _Rows[rowCnt];
                for (int colCnt = 0; colCnt < _NrOfCols; colCnt++)
                {
                    indexMatrix[rowCnt, colCnt] = r.Arr[colCnt];
                }
            }
            return indexMatrix;
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

        private class ExRow
        {
            private object[] _Arr;
            public object[] Arr
            {
                get { return _Arr; }
                set { _Arr = value; }
            }
            public ExRow(int cols)
            {
                _Arr = new object[cols];
                for (int i = 0; i < cols; i++)
                {
                    _Arr[i] = Missing.Value; // "";
                }
            }
        }
    }
}
