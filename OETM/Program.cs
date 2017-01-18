using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using ExcelToMySql;
using ExcelToMySql.Excel;
using ExcelToMySql.MySql;

namespace OETM
{
    class Program
    {
        class Options
        {
            [Option(shortName:'p', HelpText = @"Excel file location. (ex. C:\\* or C:\\temp.xlsx or ..\*)", Required = true)]
            public string Path { get; set; }

            [Option(shortName:'f', DefaultValue = "", HelpText = "Prefix to be appended when tables are created in MySql.")]
            public string TablePrefix { get; set; }
        }

        static StringBuilder _sql = new StringBuilder();
        static object _sqlLockObject = new object();
        static object _consoleLockObject = new object();
        static Options _options = new Options();

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

        static void WriteToSql(string sql)
        {
            lock (_sqlLockObject)
            {
                _sql.Append("\n");
                _sql.Append("\n");
                _sql.Append(sql);
            }
        }

        static void GenerateSql(string absoluteFilePath)
        {
            if(Path.GetExtension(absoluteFilePath).CompareTo(@".xlsx") != 0)
            {
                return;
            }

            try
            {
                ExcelMetaData metaData;
                var config = new Configuration
                {
                    TableName = _options.TablePrefix + Path.GetFileNameWithoutExtension(absoluteFilePath),
                    IsIgnoreNotFoundTypeColumn = true,
                    IgnoreIfIncludeString = new string[] { "ref", "text" },
                    YourStringType = new string[] { "ref", "text" },
                    MultiKeyTableName = new string[] { _options.TablePrefix + "actor_data" },
                };
                ExcelReader.ReadExcel(absoluteFilePath, config, out metaData);

                var table = new SqlTable(metaData, config);
                var query = table.GenerateSql();

                WriteToSql(query);

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
            catch(DuplicateColumnException e)
            {
                WriteConsole(false, string.Format("Failed! \"{0}\"", e.DupMessage));
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
            Parser.Default.ParseArgumentsStrict(args, _options);

            var files = Directory.GetFiles(Path.GetDirectoryName(_options.Path));
            var tasks = new List<Task>();
            foreach (var i in files)
            {
                tasks.Add(Task.Run(() => GenerateSql(i)));
            }

            foreach (var i in tasks)
            {
                i.Wait();
            }

            File.WriteAllText(@"oetm.sql", _sql.ToString());

            return 0;
        }
    }
}
