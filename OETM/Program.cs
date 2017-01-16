using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using ExcelToMySql.Excel;
using ExcelToMySql.MySql;

namespace OETM
{
    class Program
    {
        static StringBuilder _sql = new StringBuilder();
        static object _sqlLockObject = new object();

        static void RunGenerateSql_Task(string absoluteFilePath)
        {
            if(Path.GetExtension(absoluteFilePath).CompareTo(@".xlsx") != 0)
            {
                return;
            }
            
            ExcelMetaData metaData;
            ExcelReader.ReadExcel(absoluteFilePath, out metaData);

            var config = new SqlTableConfiguration
            {
                TableName = Path.GetFileNameWithoutExtension(absoluteFilePath)
            };

            var table = new SqlTable(metaData, config);

            try
            {
                var query = table.GenerateSql();

                lock (_sqlLockObject)
                {
                    _sql.Append("\n");
                    _sql.Append("\n");
                    _sql.Append(query);
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Complete generate! \"{0}\"", absoluteFilePath);
                Console.ResetColor();
            }
            catch(NotFoundTypeException e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed! \"{0}\" -> \"{1}\"", absoluteFilePath, e.Message);
                Console.ResetColor();
            }
            catch(NotMappedTypeException e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed! \"{0}\" -> \"{1}\"", absoluteFilePath, e.Message);
                Console.ResetColor();
            }
        }

        static void WriteInvalidOption()
        {
            Console.WriteLine("TODO");
        }

        static int Main(string[] args)
        {
            if(args.Length <= 0 || args.Length > 1)
            {
                WriteInvalidOption();
                return -1;
            }

            if(Path.GetFileName(args[0]).CompareTo(@"*") == 0)
            {
                var files = Directory.GetFiles(Path.GetDirectoryName(args[0]));
                var tasks = new List<Task>();
                foreach(var i in files)
                {
                    tasks.Add(Task.Run(() => RunGenerateSql_Task(i)));
                }

                foreach(var i in tasks)
                {
                    i.Wait();
                }
            }
            else
            {
                RunGenerateSql_Task(args[0]);
            }

            File.WriteAllText(@"oetm.sql", _sql.ToString());

            return 0;
        }
    }
}
