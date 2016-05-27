using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SpreadsheetLight;

namespace SpreadsheetWriter
{
    /// <summary>
    /// This class contains a cell writer that supports the SpreadsheetLight library.  This 
    /// should contain everything that is needed in order to use writer combinators with 
    /// SpreadsheetLight documents.
    /// </summary>
    public static class SpreadsheetLight
    {
        public class SheetProvider : ISheet
        {
            // SpreadsheetLight doesn't expose worksheets as objects, so we hold a reference to 
            // the SLDocument instead.  In order to write different sheets, the caller must change 
            // the active worksheet in the SLDocument before passing it to a writer.
            public SLDocument Document { get; private set; }

            public SheetProvider(SLDocument doc)
            {
                Document = doc;
            }

            public void WriteCell<T>(CellIndex index, T value)
            {
                if (typeof(T) == typeof(string))
                    Document.SetCellValue(index.Row, index.Col, (string)(object)value);

                if (typeof(T) == typeof(int))
                    Document.SetCellValue(index.Row, index.Col, (int)(object)value);

                if (typeof(T) == typeof(long))
                    Document.SetCellValue(index.Row, index.Col, (long)(object)value);

                if (typeof(T) == typeof(double))
                    Document.SetCellValue(index.Row, index.Col, (double)(object)value);

                if (typeof(T) == typeof(DateTime))
                    Document.SetCellValue(index.Row, index.Col, (DateTime)(object)value);

                throw new NotSupportedException(String.Format(
                    "Value of type {0} cannot be written to a cell",
                    typeof(T).Name));
            }
        }
    }
}
