using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_XLSX_reader_writer.Model
{
    public class dataList : IDataList
    {
        public string stringValue1 { get; set; }
        public string stringValue2 { get; set; }
        public string stringValue3 { get; set; }
        public string stringValue4 { get; set; }
    }
    public interface IDataList
    {
        string stringValue1 { get; set; }
        string stringValue2 { get; set; }
        string stringValue3 { get; set; }
        string stringValue4 { get; set; }
    }
}
