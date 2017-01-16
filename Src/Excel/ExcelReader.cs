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
                    // 데이터를 읽을때 필드수가 모자라서 인덱스 오버플로우 에러가 나므로 임시로 넣고 뒤에서 지운다.
                    name = "null";
                    outIgnoreFields.Add(i);
                    goto Next;
                }

                // 특정 문자열이 포함 된 필드 무시
                foreach (var j in config.IgnoreIfIncludeString)
                {
                    if (name.Contains(j))
                    {
                        name = "null";
                        outIgnoreFields.Add(i);
                        goto Next;
                    }
                }

                // 타입이 존재하지 않는 필드 무시
                var isNotFoundType = true;
                foreach (var j in config.SqlTypeMap)
                {
                    if (name.Contains(j.Key))
                    {
                        isNotFoundType = false;
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
                        name = "null";
                        outIgnoreFields.Add(i);
                        goto Next;
                    }
                }

                // 컬럼 공백 제거
                name = name.Trim();

                Next:

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
                                var isString = false;
                                foreach(var j in config.YourStringType)
                                {
                                    if(outMetaData.ColumnNames[i].Contains(j))
                                    {
                                        isString = true;
                                        break;
                                    }
                                }

                                if(isString)
                                {
                                    value = "null";
                                }
                                else
                                {
                                    value = 0;
                                }
                            }

                            // Add data
                            row.Add(value);
                        }

                        outMetaData.Datas.Add(row);
                    }

                    // 임시로 넣어진 null 컬럼을 제거 한다.
                    outMetaData.ColumnNames.RemoveAll(e => e == "null");

                    // 중복 키를 찾는다.
                    var duplicateKeys = outMetaData.ColumnNames.GroupBy(x => x)
                        .Where(group => group.Count() > 1)
                        .Select(group => group.Key);

                    if(duplicateKeys.ToList().Count > 0)
                    {
                        throw new DuplicateColumnException(duplicateKeys.ToList());
                    }
                }
            }
        }
    }
}
