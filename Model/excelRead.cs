using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;


namespace Excel_XLSX_reader_writer.Model
{
    public class excelRead
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);

        public static void ReadExcelFile(string fileName, string sheetName)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            List<IDataList> excelDataList = new List<IDataList>();

            try
            {
                // Start Excel and get Application object.
                oXL = new Excel.Application();
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(fileName, UpdateLinks: false, ReadOnly: true));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                //Count of rows used
                int lastRow = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int i = 6; i <= lastRow; i++)
                {
                    System.Array cellValues = (System.Array)oSheet.get_Range("A" + i.ToString(),
                           "T" + i.ToString()).Cells.Value2;
                    excelDataList.Add(new dataList
                    {
                        stringValue1 = cellValues.GetValue(1, 2).ToString(),
                        stringValue2 = cellValues.GetValue(1, 3).ToString(),
                        stringValue3 = cellValues.GetValue(1, 5).ToString(),
                        stringValue4 = cellValues.GetValue(1, 6).ToString()
                    });
                }
            }
            catch
            {

            }
            string monkey = "fish";

        }

    }
}
