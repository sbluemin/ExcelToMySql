using System;

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
}
