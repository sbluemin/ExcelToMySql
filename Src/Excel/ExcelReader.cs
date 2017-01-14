using System;
using System.Collections.Generic;
using System.IO;
using Excel;

namespace ExcelToMySql.Excel
{
    /// <summary>
    /// A .xlsx to ExcelMetaData converter.
    /// </summary>
    public class ExcelReader
    {
        [Obsolete("Not used.")]
        public readonly string[] IgnoreTypes = new string[] { "text", "ref" };

        /// <summary>
        /// Read column name from .xlsx file.
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="metaData"></param>
        /// <returns></returns>
        private static bool ReadColumnName(IExcelDataReader reader, ExcelMetaData metaData)
        {
            if (!reader.Read())
            {
                return false;
            }

            for (int i = 0; i < reader.FieldCount; i++)
            {
                metaData.ColumnNames.Add(reader.GetString(i));
            }

            return true;
        }

        /// <summary>
        /// Read .xlsx file and convert to ExcelMetaData.
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
                        if (!ReadColumnName(excelReader, outMetaData))
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
