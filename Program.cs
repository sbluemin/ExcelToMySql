using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel;

namespace ExcelToMySql
{
    /// <summary>
    /// 엑셀 파일에서 읽어들인 메타 데이터
    /// </summary>
    class ExcelMetaData
    {
        public List<string> ColumnName = new List<string>();
        public List<List<object>> Datas = new List<List<object>>();
    }

    /// <summary>
    /// 엑셀 데이터를 읽어들이는 클래스
    /// </summary>
    class ExcelReader
    {
        /// <summary>
        /// 엑셀에서 컬럼을 읽을 때 무시 할 데이터 포맷
        /// </summary>
        public readonly string[] IgnoreTypes = new string[] { "text", "ref" };

        private static bool ReadColumn(FileStream stream, ExcelMetaData metaData)
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                reader.IsFirstRowAsColumnNames = true;
                var result = reader.AsDataSet();

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
                    // 컬럼 정보를 읽음
                    if (!ReadColumn(stream, outMetaData))
                    {
                        throw new Exception("Read failed.");
                    }

                    using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        // 초기 위치에는 컬럼 정보가 있으므로 한번 Read한다.
                        if (!excelReader.Read())
                        {
                            return true;
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

    class Program
    {
        public class SqlTableConfiguration
        {
            public string TableName;
        }

        public class SqlTable
        {
            /*
                note: mysqldump output data

                DROP TABLE IF EXISTS `tb_data_bot_data`;
                CREATE TABLE `tb_data_bot_data` (
                    `int_bot_data_tid` int(11) NOT NULL,
                    `int_power_min` int(11) NOT NULL,
                    `int_power_max` int(11) NOT NULL,
                    `int_array_stage_team_tid_1` int(11) NOT NULL,
                    `int_array_stage_team_tid_2` int(11) NOT NULL,
                    `int_array_stage_team_tid_3` int(11) NOT NULL,
                    `int_array_stage_team_tid_4` int(11) NOT NULL,
                    `int_array_stage_team_tid_5` int(11) NOT NULL,
                    PRIMARY KEY(`int_bot_data_tid`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8;

                --
                -- Dumping data for table `tb_data_bot_data`
                --

                LOCK TABLES `tb_data_bot_data` WRITE;
                INSERT IGNORE INTO `tb_data_bot_data` VALUES(1,9000,12000,100101,100102,100103,100104,100105),(2,12000,15000,100201,100202,100203,100204,100205);
                UNLOCK TABLES;
             */

            public readonly ExcelMetaData ExcelMetaData;
            public readonly SqlTableConfiguration Configuration;

            public SqlTable(ExcelMetaData metaData, SqlTableConfiguration config)
            {
                ExcelMetaData = metaData;
                Configuration = config;
            }
               
            private void NewQuery_DropTable(StringBuilder builder)
            {
                builder.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", Configuration.TableName);
            }

            private void NewQuery_CreateTable(StringBuilder builder)
            {
                builder.AppendFormat("CREATE TABLE `{0}`\n", Configuration.TableName);
            }

            /// <summary>
            /// 메타 데이터를 토대로 테이블을 생성하는 쿼리문을 만듭니다.
            /// </summary>
            /// <returns></returns>
            public string GenerateSql()
            {
                var builder = new StringBuilder();
                NewQuery_DropTable(builder);
                NewQuery_CreateTable(builder);

                return builder.ToString();
            }
        }


        static void Main(string[] args)
        {
            // 1. Read excel data and convert meta data.
            ExcelMetaData metaData;
            ExcelReader.ReadExcel(@"C:\Temp\actor_data.xlsx", out metaData);


            // 2. Generate sql(like .sql file) from SqlTable.
            var config = new SqlTableConfiguration
            {
                TableName = "actor_data"
            };

            var table = new SqlTable(metaData, config);
            var query = table.GenerateSql();

            // ex
            Console.WriteLine(query);
        }
    }
}