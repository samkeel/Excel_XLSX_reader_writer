using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Excel_XLSX_reader_writer.Model
{
    public class excelRead
    {
        public static void ReadLargeFile(string fileName)
        {
            // SAX method to read large file, cell by cell.
            // prefrred approach to large files to prevent 'out of memory' errors

            // Open the document in read only
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {

            }
        }
    }
}
