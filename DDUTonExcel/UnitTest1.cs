using System;
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
