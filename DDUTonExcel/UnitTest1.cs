﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;


namespace DDUTonExcel
{
    [TestClass]
    public class UnitTest1
    {
        const string WORKBOOK_NAME = @"D:\GitHub\DDUTonExcel\DDUTonExcel\Sample.xlsx";
        const string SHEET_NAME = "SampleSheet";
        private TestContext testContextInstance;

        public TestContext TestContext
        {
            get { return testContextInstance; }
            set { testContextInstance = value; }
        }

        public UnitTest1()
        {

        }

        [DataSource(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\GitHub\DDUTonExcel\DDUTonExcel\Sample.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES"";",
            "SampleSheet$")]
        [TestMethod()]
        public void CheckColumnsHeader()
        {
            
            Assert.AreEqual(1, 1);
        }

        [TestMethod()]
        public void CheckColumnsHeader2()
        {
            OpenExcel(WORKBOOK_NAME, SHEET_NAME);
        }

        public System.Data.DataTable GetWorksheet(string worksheetName)
        {
            OleDbConnection con = new System.Data.OleDb.OleDbConnection(connectionString);
            OleDbDataAdapter cmd = new System.Data.OleDb.OleDbDataAdapter(
                "select * from [" + worksheetName + "$]", con);

            con.Open();
            System.Data.DataSet excelDataSet = new DataSet();
            cmd.Fill(excelDataSet);
            con.Close();

            DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, sheetName, null });

            List<string> listColumn = new List<string>();
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
            workSheetNames = new String[dataSet.Rows.Count];
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

        private void OpenExcel(string workBookName, string sheetName)
        {
            Excel.Application excelapplication = null;
            Excel.Workbook workbook = null;

            try
            {
                excelapplication = new Excel.Application();
                workbook = excelapplication.Workbooks.Open(workBookName);
                var errors = new Dictionary<string, List<string>>();
                
                var sheet = workbook.Sheets[sheetName];

                int rowCount = sheet.UsedRange.Cells.Rows.Count;
                int colCount = sheet.UsedRange.Cells.Columns.Count;
                var usedCells = sheet.UsedRange.Cells;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        Excel.Range range = usedCells[i, j];
                        List<string> cellErrors = new List<string>();
                        if (!IsHeaderCell(workbook, range))
                        {
                            string cellDisplayTitle = String.Format("{0}!{1}", sheet.Name, range.Address);
                            errors[cellDisplayTitle] = cellErrors;
                        }

                    }
                }
                //ReportErrors(errors);
            }
            finally
            {
                if (workbook != null)
                    workbook.Close();
                if (excelapplication != null)
                    excelapplication.Quit();
            }
        }

        private bool IsHeaderCell(Excel.Workbook workbook, Excel.Range range)
        {
            // Look through workbook names:
            foreach (Excel.Name namedRange in workbook.Names)
            {
                if (range.Parent == namedRange.RefersToRange.Parent && range.Application.Intersect(range, namedRange.RefersToRange) != null)
                    return true;
            }

            // Look through worksheet-names.
            foreach (Excel.Name namedRange in range.Worksheet.Names)
            {
                if (range.Parent == namedRange.RefersToRange.Parent && range.Worksheet.Application.Intersect(range, namedRange.RefersToRange) != null)
                    return true;
            }
            return false;
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
