using System.Collections.Generic;

namespace ExcelToMySql.Excel
{
    /// <summary>
    /// Data read from an .xlsx file
    /// </summary>
    public class ExcelMetaData
    {
        public List<string> ColumnName = new List<string>();
        public List<List<object>> Datas = new List<List<object>>();
    }
}
