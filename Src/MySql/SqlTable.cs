using System.Text;
using ExcelToMySql.Excel;

namespace ExcelToMySql.MySql
{
    /// <summary>
    /// A table query generator for MySql
    /// </summary>
    public class SqlTable
    {
        public readonly ExcelMetaData MetaData;
        public readonly SqlTableConfiguration Configuration;

        public SqlTable(ExcelMetaData metaData, SqlTableConfiguration config)
        {
            MetaData = metaData;
            Configuration = config;
        }

        /// <summary>
        /// Generate 'DROP TABLE' query.
        /// </summary>
        /// <param name="builder"></param>
        public void NewQuery_DropTable(StringBuilder builder)
        {
            builder.AppendFormat("DROP TABLE IF EXISTS `{0}`;\n", Configuration.TableName);
        }

        /// <summary>
        /// Generate 'CREATE TABLE' query.
        /// </summary>
        /// <param name="builder"></param>
        public void NewQuery_CreateTable(StringBuilder builder)
        {
            builder.AppendFormat("CREATE TABLE `{0}` (\n", Configuration.TableName);

            foreach (var i in MetaData.ColumnNames)
            {
                foreach(var j in Configuration.SqlTypeMap)
                {
                    if (i.Contains(j.Key))
                    {
                        builder.AppendFormat("`{0}` {1} NOT NULL,\n", i, j.Value);
                        break;
                    }
                }
            }

            builder.AppendFormat("PRIMARY KEY(`{0}`)\n", MetaData.ColumnNames[0]);
            builder.AppendFormat(") ENGINE=InnoDB DEFAULT CHARSET=utf8;\n");
        }

        /// <summary>
        /// Generate 'INSERT' query.
        /// </summary>
        /// <param name="builder"></param>
        public void NewQuery_AddDatas(StringBuilder builder)
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
        /// Generate new table sql by ExcelMetaData
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
