using System;
using ExcelToMySql.Excel;
using ExcelToMySql.MySql;

namespace ExcelToMySql
{

    class Program
    {
        static void Main(string[] args)
        {
            // 1. Read excel data and convert meta data.
            ExcelMetaData metaData;
            ExcelReader.ReadExcel(@"C:\Temp\aa.xlsx", out metaData);

            // 2. Generate sql(like .sql file) from SqlTable.
            var config = new SqlTableConfiguration
            {
                TableName = "actor_data"
            };

            var table = new SqlTable(metaData, config);
            var query = table.GenerateSql();

            // ex
            System.IO.File.WriteAllText(@".\aa.sql", query);
            Console.WriteLine(query);
        }
    }
}