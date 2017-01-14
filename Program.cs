using System;
using System.Collections.Generic;
using System.IO;
using Excel;

namespace ExcelToMySql
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

    class Program
    {
        public class Table
        {
            // Excel 데이터에서 읽어왔을때의 데이터들
            public List<string> ColumnName = new List<string>();
            public List<List<object>> Datas = new List<List<object>>();

            // 1차 가공 데이터들
        }

        public static bool ReadColumn(FileStream stream, Table table)
        {
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                reader.IsFirstRowAsColumnNames = true;
                var result = reader.AsDataSet();

                if(!reader.Read())
                {
                    return false;
                }

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    table.ColumnName.Add(reader.GetString(i));
                }

                return true;
            }
        }

        public static bool ReadExcel(string absoluteFilePath)
        {
            try
            {
                using (var stream = File.Open(absoluteFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {


                        var table = new Table();



                        // 이후 데이터 읽기
                        while (excelReader.Read())
                        {
                            for (int i = 0; i < excelReader.FieldCount; i++)
                            {
                                
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        static void Main(string[] args)
        {
            ReadExcel(@"C:\Temp\item_grade.xlsx");
        }
    }
}