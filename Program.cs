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

        private static bool ReadColumn(IExcelDataReader reader, ExcelMetaData metaData)
        {
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
                    using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        // 컬럼 정보를 읽음
                        excelReader.IsFirstRowAsColumnNames = true;
                        if (!ReadColumn(excelReader, outMetaData))
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

                /*
                 *
                 DROP TABLE IF EXISTS `tb_data_cutscene_base`;
                 CREATE TABLE `tb_data_cutscene_base` (
                    `int_cutscene_tid` int(11) NOT NULL,
                    `int_cutscene_group_tid` int(11) NOT NULL,
                    `text_prefabname` longtext,
                    `int_next_cutscene_tid` int(11) NOT NULL,
                    `bool_start` smallint(6) NOT NULL,
                    `bool_end` smallint(6) NOT NULL,
                    `bool_dontsave` smallint(6) NOT NULL,
                    `bool_timesave` smallint(6) NOT NULL,
                    `int_load_fx_cutscene_tid` int(11) NOT NULL,
                    PRIMARY KEY(`int_cutscene_tid`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8;

                --
                -- Dumping data for table `tb_data_cutscene_base`
                --

                LOCK TABLES `tb_data_cutscene_base` WRITE;
                INSERT IGNORE INTO `tb_data_cutscene_base` VALUES(1,1,'C1D1_Cutscene_1',2,1,0,0,1,1),(2,1,'C1D1_Cutscene_2',3,0,0,0,1,2),(3,1,'C1D1_Cutscene_3',4,0,1,0,1,3),(4,1,'Cutscene_loop',-1,0,0,1,0,4),(5,2,'C1D2_Cutscene_1',6,1,0,0,1,5),(6,2,'C1D2_Cutscene_3',7,0,1,0,1,6),(7,2,'Cutscene_loop',-1,0,0,1,0,7),(8,5,'Cutscene_c1d5_2_2',9,0,0,0,0,8),(9,5,'Cutscene_c1d5_2_3',17,0,0,0,1,9),(10,3,'C1D3_Cutscene_1',11,1,0,0,1,10),(11,3,'C1D3_Cutscene_2',12,0,1,0,1,11),(12,3,'Cutscene_loop',-1,0,0,1,0,12),(13,4,'Cutscene_c1d4c1',14,1,1,0,1,13),(14,4,'Cutscene_loop',-1,0,0,1,0,14),(15,5,'Cutscene_c1d5_1',16,1,0,0,1,15),(16,5,'Cutscene_c1d5_2_1',8,0,0,0,1,16),(17,5,'Cutscene_c1d5_3',-1,0,1,0,1,17),(18,6,'Cutscene_c1d5_lobby',-1,1,0,1,0,18),(19,7,'Cutscene_c1a6',20,1,1,0,1,19),(20,7,'Cutscene_c1a6_lobby',-1,0,0,1,0,20),(21,8,'Cutscene_c1a7_1',22,1,0,0,1,21),(22,8,'Cutscene_c1a7_2',-1,0,1,0,0,22),(24,9,'Cutscene_c2a1_1',26,1,0,0,1,24),(26,9,'Cutscene_c2a1_2',29,0,1,0,1,26),(28,11,'Cutscene_c2a2_2',29,1,1,0,1,28),(29,11,'Cutscene_c2_Lobby',-1,0,0,1,0,29),(30,12,'Cutscene_c2a2_3',29,1,1,0,1,30),(32,13,'Cutscene_c2a2_4',29,1,1,0,1,32),(34,14,'Cutscene_c2a3',35,1,1,0,1,34),(35,14,'Cutscene_loop',-1,0,0,1,0,35),(36,15,'Cutscene_c2a4_2',37,1,1,0,1,36),(37,15,'Cutscene_c2a4_Lobby',-1,0,0,1,0,37),(38,16,'Cutscene_c2a2_0_1',39,1,0,0,1,38),(39,16,'Cutscene_c2a2_0_2',-1,0,1,0,0,39);
                UNLOCK TABLES;

                 */

            public readonly ExcelMetaData MetaData;
            public readonly SqlTableConfiguration Configuration;

            public SqlTable(ExcelMetaData metaData, SqlTableConfiguration config)
            {
                MetaData = metaData;
                Configuration = config;
            }

            /// <summary>
            /// 규약에 맞는 컬럼의 포맷을 MySQL에 맞는 string형으로 반환 합니다.
            /// </summary>
            /// <param name="columnName"></param>
            /// <returns></returns>
            private string GetTypeFromColumnName(string columnName)
            {
                var strings = columnName.Split('_');
                return strings[0];
            }
              
            private void NewQuery_DropTable(StringBuilder builder)
            {
                builder.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", Configuration.TableName);
            }

            private void NewQuery_CreateTable(StringBuilder builder)
            {
                builder.AppendFormat("CREATE TABLE `{0}` (\n", Configuration.TableName);

                foreach(var i in MetaData.ColumnName)
                {
                    builder.AppendFormat("`{0}` {1},\n", i, GetTypeFromColumnName(i));
                }

                builder.AppendFormat("PRIMARY KEY(`{0}`)\n", MetaData.ColumnName[0]);
                builder.AppendFormat(") ENGINE=InnoDB DEFAULT CHARSET=utf8;\n");
            }

            private void NewQuery_AddDatas(StringBuilder builder)
            {
                builder.AppendFormat("LOCK TABLES `{0}` WRITE;\n", Configuration.TableName);

                foreach (var i in MetaData.Datas)
                {
                    foreach(var j in i)
                    {
                        ;
                    }
                }

                builder.AppendFormat("UNLOCK TABLES;\n");
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
                NewQuery_AddDatas(builder);

                return builder.ToString();
            }
        }


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
            Console.WriteLine(query);
        }
    }
}