using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToMySql.Excel
{
    /// <summary>
    /// 엑셀 파일에서 읽어들인 메타 데이터
    /// </summary>
    public class ExcelMetaData
    {
        public List<string> ColumnName = new List<string>();
        public List<List<object>> Datas = new List<List<object>>();
    }
}
