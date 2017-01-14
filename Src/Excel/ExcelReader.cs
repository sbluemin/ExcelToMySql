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
        /// <param name="config"></param>
        /// <returns></returns>
        private static bool ReadColumnName(IExcelDataReader reader, ExcelMetaData metaData, ExcelReaderConfiguration config)
        {
            if (!reader.Read())
            {
                return false;
            }

            for (int i = config.DataEntryPointColumnIndex; i < reader.FieldCount; i++)
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
        public static void ReadExcel(string absoluteFilePath, out ExcelMetaData outMetaData)
        {
            var config = new ExcelReaderConfiguration();
            ReadExcel(absoluteFilePath, config, out outMetaData);
        }

        /// <summary>
        /// Read .xlsx file and convert to ExcelMetaData.
        /// </summary>
        /// <param name="absoluteFilePath"></param>
        /// <param name="config"></param>
        /// <param name="outMetaData"></param>
        /// <returns></returns>
        public static void ReadExcel(string absoluteFilePath, ExcelReaderConfiguration config, out ExcelMetaData outMetaData)
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
                    if (!ReadColumnName(excelReader, outMetaData, config))
                    {
                        throw new Exception("Read failed.");
                    }

                    // Read datas
                    while (excelReader.Read())
                    {
                        var row = new List<object>();
                        for (int i = config.DataEntryPointColumnIndex; i < excelReader.FieldCount; i++)
                        {
                            row.Add(excelReader.GetValue(i));
                        }

                        outMetaData.Datas.Add(row);
                    }
                }
            }
        }
    }
}
