using ExcelToMySql.MySql;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Text;
using ExcelToMySql.Excel;
using System.Collections.Generic;

namespace ExcelToMySql.MySql.Tests
{
    [TestClass()]
    public class SqlTableTests
    {
        const string _tableName = "SqlTableTest";

        [TestMethod()]
        public void NewQuery_DropTableTest()
        {
            var table = new SqlTable(new ExcelMetaData(), new SqlTableConfiguration
            {
                TableName = "SqlTableTest"
            });

            var builder = new StringBuilder();
            table.NewQuery_DropTable(builder);

            Assert.AreEqual("DROP TABLE IF EXISTS `SqlTableTest`;\n", builder.ToString());
        }

        [TestMethod()]
        public void NewQuery_CreateTableTest()
        {
            var config = new SqlTableConfiguration
            {
                TableName = "SqlTableTest"
            };

            var metaData = new ExcelMetaData();
            metaData.ColumnNames.Add("text_column_name_1");
            metaData.ColumnNames.Add("int_column_name_2");

            Assert.Fail("Not used.");
        }

        [TestMethod()]
        public void NewQuery_AddDatasTest()
        {
            var config = new SqlTableConfiguration
            {
                TableName = "SqlTableTest"
            };

            var metaData = new ExcelMetaData();
            metaData.ColumnNames.Add("text_column_name_1");
            metaData.ColumnNames.Add("int_column_name_2");
            metaData.Datas.Add(new List<object>
            {
                "test_field",
                100
            });


            Assert.Fail("Not used.");
        }
    }
}