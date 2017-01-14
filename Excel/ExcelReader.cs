using System;
using System.Collections.Generic;
using System.IO;
using Excel;

namespace ExcelToMySql.Excel
{
    /// <summary>
    /// 엑셀 데이터를 읽어들이는 클래스
    /// </summary>
    class ExcelReader
    {
        /// <summary>
        /// 엑셀에서 컬럼을 읽을 때 무시 할 데이터 포맷
        /// </summary>
        public readonly string[] IgnoreTypes = new string[] { "text", "ref" };

        private static bool ReadColumn(IExcelDataReader reader, ExcelMetaData metaData)
        {
            if (!reader.Read())
            {
                return false;
            }

            for (int i = 0; i < reader.FieldCount; i++)
            {
                metaData.ColumnName.Add(reader.GetString(i));
            }

            return true;
        }

        /// <summary>
        /// 엑셀로부터 데이터와 포맷을 읽어와 메타 데이터화 시킵니다.
        /// </summary>
        /// <param name="absoluteFilePath"></param>
        /// <param name="outMetaData"></param>
        /// <returns></returns>
        public static bool ReadExcel(string absoluteFilePath, out ExcelMetaData outMetaData)
        {
            outMetaData = new ExcelMetaData();

            try
            {
                using (var stream = File.Open(absoluteFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        // 컬럼 정보를 읽음
                        excelReader.IsFirstRowAsColumnNames = true;
                        if (!ReadColumn(excelReader, outMetaData))
                        {
                            throw new Exception("Read failed.");
                        }

                        // 이후 데이터 읽기
                        while (excelReader.Read())
                        {
                            // 컬럼 정보는 0부터 있고 데이터는 그 뒤에 있다.
                            var row = new List<object>();
                            for (int i = 0; i < excelReader.FieldCount; i++)
                            {
                                row.Add(excelReader.GetValue(i));
                            }

                            outMetaData.Datas.Add(row);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return true;
        }
    }
}
