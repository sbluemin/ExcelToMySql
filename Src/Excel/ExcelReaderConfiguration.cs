namespace ExcelToMySql.Excel
{
    public class ExcelReaderConfiguration
    {
        /// <summary>
        /// Ignore if include specify string.
        /// </summary>
        public string[] IgnoreIfIncludeString = new string[] { };

        public int DataEntryPointColumnIndex = 0;
        public int DataEntryPointRowIndex = 0;
    }
}
