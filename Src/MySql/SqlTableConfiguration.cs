using System.Collections.Generic;

namespace ExcelToMySql.MySql
{
    public class SqlTableConfiguration
    {
        private static readonly Dictionary<string, string> _defaultSqlTypeMap = new Dictionary<string, string>()
            {
                 {"int", "int(11)"},
                 {"short", "smallint(6)"},
                 {"char", "char(1)" },
                 {"bool", "int(11)" },
                 {"byte", "char(1)" },
                 {"float", "double" },
                 {"double", "double" },
                 {"text", "varchar(255)"},
                 {"ref", "varchar(255)"},
            };

        /// <summary>
        /// Mapping your custom field type(.xlsx) to sql type.
        /// </summary>
        public readonly Dictionary<string, string> SqlTypeMap = _defaultSqlTypeMap;

        /// <summary>
        /// The name of the table to be created in mysql.
        /// </summary>
        public string TableName;

        public bool IsIgnoreNotFoundTypeColumn = false;
    }
}
