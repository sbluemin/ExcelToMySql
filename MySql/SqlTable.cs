using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelToMySql.Excel;

namespace ExcelToMySql.MySql
{
    public class SqlTable
    {
        public readonly ExcelMetaData MetaData;
        public readonly SqlTableConfiguration Configuration;

        public SqlTable(ExcelMetaData metaData, SqlTableConfiguration config)
        {
            MetaData = metaData;
            Configuration = config;
        }

        private readonly Dictionary<string, string> _sqlTypeMap = new Dictionary<string, string>()
            {
                 {"int", "int(11)"},
                 {"short", "smallint(6)"},
                 {"char", "char(1)" },
                 {"byte", "char(1)" },
                 {"text", "varchar(255)"},
                 {"ref", "varchar(255)"},
            };

        /// <summary>
        /// 규약에 맞는 컬럼의 포맷을 MySQL에 맞는 string형으로 반환 합니다.
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private string GetTypeFromColumnName(string columnName) => _sqlTypeMap[columnName.Split('_')[0]];

        /// <summary>
        /// 테이블을 DROP 시키는 쿼리를 만든다.
        /// </summary>
        /// <param name="builder"></param>
        private void NewQuery_DropTable(StringBuilder builder)
        {
            builder.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", Configuration.TableName);
        }

        /// <summary>
        /// 테이블을 생성하는 쿼리를 만든다.
        /// </summary>
        /// <param name="builder"></param>
        private void NewQuery_CreateTable(StringBuilder builder)
        {
            builder.AppendFormat("CREATE TABLE `{0}` (\n", Configuration.TableName);

            foreach (var i in MetaData.ColumnName)
            {
                builder.AppendFormat("`{0}` {1} NOT NULL,\n", i, GetTypeFromColumnName(i));
            }

            builder.AppendFormat("PRIMARY KEY(`{0}`)\n", MetaData.ColumnName[0]);
            builder.AppendFormat(") ENGINE=InnoDB DEFAULT CHARSET=utf8;\n");
        }

        /// <summary>
        /// 테이블에 데이터를 추가하는 쿼리를 만든다.
        /// </summary>
        /// <param name="builder"></param>
        private void NewQuery_AddDatas(StringBuilder builder)
        {
            builder.AppendFormat("LOCK TABLES `{0}` WRITE;\n", Configuration.TableName);
            builder.AppendFormat("INSERT IGNORE INTO `{0}` VALUES", Configuration.TableName);

            var fieldCount = MetaData.Datas[0].Count;

            for (var i = 0; i < MetaData.Datas.Count; i++)
            {
                builder.Append("(");
                for (var j = 0; j < MetaData.Datas[i].Count; j++)
                {
                    if (j + 1 >= MetaData.Datas[i].Count)
                    {
                        if (MetaData.Datas[i][j] is string)
                        {
                            builder.AppendFormat("'{0}'", MetaData.Datas[i][j]);
                        }
                        else
                        {
                            builder.AppendFormat("{0}", MetaData.Datas[i][j]);
                        }
                    }
                    else
                    {
                        if (MetaData.Datas[i][j] is string)
                        {
                            builder.AppendFormat("'{0},'", MetaData.Datas[i][j]);
                        }
                        else
                        {
                            builder.AppendFormat("{0},", MetaData.Datas[i][j]);
                        }
                    }
                }

                if (i + 1 >= MetaData.Datas.Count)
                {
                    builder.Append(");\n");
                }
                else
                {
                    builder.Append("),\n");
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
}
