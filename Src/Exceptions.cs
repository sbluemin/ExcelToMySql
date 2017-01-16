using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToMySql
{
    public class NotMappedTypeException : Exception
    {
        public NotMappedTypeException(string yourTypeName, string expectedTypeName) 
            : base(string.Format("Not mapped type : \"{0}\" -> \"{1}\"", yourTypeName, expectedTypeName))
        { }
    }

    public class NotFoundTypeException : Exception
    {
        public NotFoundTypeException(string yourColumnName)
            : base(string.Format("Not found type : \"{0}\"", yourColumnName))
        { }
    }

    public class DuplicateColumnException : Exception
    {
        public readonly string DupMessage;

        public DuplicateColumnException(List<string> duplicatedColumns)
        {
            var sb = new StringBuilder();
            sb.Append("Duplicate coulmn:");

            foreach (var i in duplicatedColumns)
            {
                sb.AppendFormat(" {0}", i);
            }

            DupMessage = sb.ToString();
        }
    }
}
