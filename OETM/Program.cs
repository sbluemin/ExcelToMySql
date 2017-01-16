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
        static object _consoleLockObject = new object();

        static void WriteConsole(bool isSuccess, string message)
        {
            lock(_consoleLockObject)
            {
                if (isSuccess)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(message);
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(message);
                    Console.ResetColor();
                }
            }
        }

        static void RunGenerateSql_Task(string absoluteFilePath)
        {
            if(Path.GetExtension(absoluteFilePath).CompareTo(@".xlsx") != 0)
            {
                return;
            }

            try
            {
                ExcelMetaData metaData;
                var excelConfig = new ExcelReaderConfiguration
                {
                    IgnoreIfIncludeString = new string[] { "ref", "text" },
                };
                ExcelReader.ReadExcel(absoluteFilePath, excelConfig, out metaData);

                var config = new SqlTableConfiguration
                {
                    TableName = Path.GetFileNameWithoutExtension(absoluteFilePath),
                    IsIgnoreNotFoundTypeColumn = true,
                };

                var table = new SqlTable(metaData, config);

                var query = table.GenerateSql();

                lock (_sqlLockObject)
                {
                    _sql.Append("\n");
                    _sql.Append("\n");
                    _sql.Append(query);
                }

                WriteConsole(true, string.Format("Success! \"{0}\"", absoluteFilePath));
            }
            catch(NotFoundTypeException e)
            {
                WriteConsole(false, string.Format("Failed! \"{0}\" -> \"{1}\"", absoluteFilePath, e.Message));
            }
            catch(NotMappedTypeException e)
            {
                WriteConsole(false, string.Format("Failed! \"{0}\" -> \"{1}\"", absoluteFilePath, e.Message));
            }
            catch (Exception e)
            {
                WriteConsole(false, string.Format("Failed! \"{0}\" \"{1}\"", absoluteFilePath, e.Message));
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
