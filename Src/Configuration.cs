using System.Collections.Generic;

namespace ExcelToMySql
{
    public class Configuration
    {
        public int DataEntryPointColumnIndex = 0;
        public int DataEntryPointRowIndex = 0;

        private static readonly Dictionary<string, string> _defaultSqlTypeMap = new Dictionary<string, string>()
            {
                 {"int", "int(11)"},
                 {"short", "smallint(6)"},
                 {"char", "smallint(6)" },
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
        /// Ignore if include specify string.
        /// </summary>
        public string[] IgnoreIfIncludeString = new string[] { };

        public string[] YourStringType = new string[] { };

        public string[] MultiKeyTableName = new string[] { };

        public bool IsIgnoreNotFoundTypeColumn = false;

        /// <summary>
        /// The name of the table to be created in mysql.
        /// </summary>
        public string TableName;
    }
}
