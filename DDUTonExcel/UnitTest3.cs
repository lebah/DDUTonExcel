using System;
using System.Data;
using System.Data.OleDb;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;


namespace DDUTonExcel
{
    [TestClass]
    public class UnitTest3
    {
        const string WORKBOOK_NAME = @"D:\GitHub\DDUTonExcel\DDUTonExcel\Sample.xlsx";
        const string SHEET_NAME = "SampleSheet";
        private TestContext testContextInstance;
        private string connectionString;
        private string fileName;
        private string sheetName;
        List<string> listColumn = new List<string>();


        public TestContext TestContext
        {
            get { return testContextInstance; }
            set { testContextInstance = value; }
        }

        public UnitTest3()
        {
            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\GitHub\DDUTonExcel\DDUTonExcel\Sample.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES"";";
            sheetName = SHEET_NAME;
        }

        [DataSource(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\GitHub\DDUTonExcel\DDUTonExcel\Sample.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES"";",
            "SampleSheet$")]
        [TestMethod()]
        public void CheckColumnsHeader()
        {

            Assert.AreEqual(1, 1);
        }

        [TestMethod()]
        public void CheckColumnsHeader3()
        {
            var dt = GetWorksheet(sheetName);
            Assert.AreEqual("Id", listColumn[0]);
        }

        public DataTable GetWorksheet(string worksheetName)
        {
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbDataAdapter cmd = new OleDbDataAdapter(
                "select * from [" + worksheetName + "$]", con);

            con.Open();
            var excelDataSet = new DataSet();
            cmd.Fill(excelDataSet);
            //con.Close();

            DataTable dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, sheetName + "$", null });

            listColumn = new List<string>();
            foreach (DataRow row in dt.Rows)
            
            {
                listColumn.Add(row["Column_name"].ToString());
            }

            return excelDataSet.Tables[0];
        }

        private void OpenExcel2(bool isOpenXMLFormat)
        {
            //open the excel file using OLEDB
            OleDbConnection con;

            if (isOpenXMLFormat)
                //read a 2007 file
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                    fileName + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            else
                //read a 97-2003 file
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    fileName + ";Extended Properties=Excel 8.0;";

            con = new OleDbConnection(connectionString);
            con.Open();

            //get all the available sheets
            System.Data.DataTable dataSet = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            //get the number of sheets in the file
            var workSheetNames = new String[dataSet.Rows.Count];
            int i = 0;
            foreach (DataRow row in dataSet.Rows)
            {
                //insert the sheet's name in the current element of the array
                //and remove the $ sign at the end
                workSheetNames[i] = row["TABLE_NAME"].ToString().Trim(new[] { '$' });
                i++;
            }

            if (con != null)
            {
                con.Close();
                con.Dispose();
            }

            if (dataSet != null)
                dataSet.Dispose();
        }

        private string ReportErrors(Dictionary<string, List<string>> errors)
        {
            var result = string.Empty;
            if (errors.Count > 0)
            {
                result += "Found the following errors:";
                result += "---------------------------------";
                result += string.Format("{0,-15} | Error", "Cell");
                result += "---------------------------------";
            }

            foreach (KeyValuePair<string, List<string>> kv in errors)
                result += string.Format("{0,-15} | {1}", kv.Key, kv.Value.Aggregate((e, s) => e + ", " + s));

            return result;
        }

        [TestMethod]
        public void TestMethod1()
        {
        }
    }
}
