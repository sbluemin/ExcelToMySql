using System;
using ExcelToMySql;
using ExcelToMySql.Excel;
using ExcelToMySql.MySql;

namespace HelloExcelToMySql
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Read excel data and convert meta data.
            ExcelMetaData metaData;
            ExcelReader.ReadExcel(@".\Excel\Sample.xlsx", out metaData);

            // 2. Generate sql(like .sql file) from SqlTable.
            var config = new Configuration
            {
                TableName = "Sample"
            };

            var table = new SqlTable(metaData, config);
            var query = table.GenerateSql();

            // Enjoy! :) 
            Console.WriteLine(query);
        }
    }
}