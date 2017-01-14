using System;

namespace ExcelToMySql.MySql
{
    public class NotMappedTypeException : Exception
    {
        public NotMappedTypeException(string yourTypeName, string expectedTypeName) 
            : base(string.Format("Not mapped type : \"{0}\" -> \"{1}\"", yourTypeName, expectedTypeName))
        { }
    }
}
