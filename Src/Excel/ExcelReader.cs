using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Excel;

namespace ExcelToMySql.Excel
{
    public static class KeyValuePairExtensions
    {
        public static bool IsNull<T, TU>(this KeyValuePair<T, TU> pair)
        {
            return pair.Equals(new KeyValuePair<T, TU>());
        }
    }

    /// <summary>
    /// A .xlsx to ExcelMetaData converter.
    /// </summary>
    public class ExcelReader
    {
        /// <summary>
        /// Read column name from .xlsx file.
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="metaData"></param>
        /// <param name="config"></param>
        /// <param name="outIgnoreFields"></param>
        /// <returns></returns>
        private static bool ReadColumnName(IExcelDataReader reader, ExcelMetaData metaData, Configuration config, out List<int> outIgnoreFields)
        {
            outIgnoreFields = new List<int>();
            if (!reader.Read())
            {
                return false;
            }

            for (int i = config.DataEntryPointColumnIndex; i < reader.FieldCount; i++)
            {
                var name = reader.GetString(i);

                // 컬럼이 null일 경우 무시
                if (name == null)
                {
                    outIgnoreFields.Add(i);
                    continue;
                }

                // 특정 문자열이 포함 된 필드 무시
                foreach (var j in config.IgnoreIfIncludeString)
                {
                    if (name.Contains(j))
                    {
                        outIgnoreFields.Add(i);
                        continue;
                    }
                }

                // 타입이 존재하지 않는 필드 무시
                var isNotFoundType = false;
                foreach (var j in config.SqlTypeMap)
                {
                    if (name.Contains(j.Key))
                    {
                        continue;
                    }
                }

                if(isNotFoundType)
                {
                    if(!config.IsIgnoreNotFoundTypeColumn)
                    {
                        throw new NotFoundTypeException(name);
                    }
                    else
                    {
                        outIgnoreFields.Add(i);
                        continue;
                    }
                }

                // 컬럼 공백 제거
                name = name.Trim();

                // 컬럼 추가
                metaData.ColumnNames.Add(name);
            }

            return true;
        }

        /// <summary>
        /// Read .xlsx file and convert to ExcelMetaData.
        /// </summary>
        /// <param name="absoluteFilePath"></param>
        /// <param name="outMetaData"></param>
        /// <returns></returns>
        public static void ReadExcel(string absoluteFilePath, out ExcelMetaData outMetaData)
        {
            var config = new Configuration();
            ReadExcel(absoluteFilePath, config, out outMetaData);
        }

        /// <summary>
        /// Read .xlsx file and convert to ExcelMetaData.
        /// </summary>
        /// <param name="absoluteFilePath"></param>
        /// <param name="config"></param>
        /// <param name="outMetaData"></param>
        /// <returns></returns>
        public static void ReadExcel(string absoluteFilePath, Configuration config, out ExcelMetaData outMetaData)
        {
            outMetaData = new ExcelMetaData();

            using (var stream = File.Open(absoluteFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    // Move data entry point.
                    for (var i = 0; i < config.DataEntryPointRowIndex; i++)
                    {
                        if (!excelReader.Read())
                        {
                            throw new Exception("Read failed.");
                        }
                    }

                    // Read column info
                    List<int> ignoreFields;
                    if (!ReadColumnName(excelReader, outMetaData, config, out ignoreFields))
                    {
                        throw new Exception("Read failed.");
                    }

                    // Read datas
                    while (excelReader.Read())
                    {
                        var row = new List<object>();
                        for (int i = config.DataEntryPointColumnIndex; i < excelReader.FieldCount; i++)
                        {
                            // 무시되는 필드는 건너뛴다.
                            if(ignoreFields.Contains(i))
                            {
                                continue;
                            }

                            var value = excelReader.GetValue(i);
                            if(value == null)
                            {
                                // Set default data if value is null
                                foreach(var j in config.YourStringType)
                                {
                                    if(outMetaData.ColumnNames[i].Contains(j))
                                    {
                                        value = "null";
                                        break;
                                    }
                                    else
                                    {
                                        value = 0;
                                        break;
                                    }
                                }
                            }

                            // Add data
                            row.Add(value);
                        }

                        outMetaData.Datas.Add(row);
                    }
                }
            }
        }
    }
}
