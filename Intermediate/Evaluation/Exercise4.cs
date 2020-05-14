using System;
using System.Collections.Generic;
using ClosedXML.Excel;
      
using NUnit.Framework;

namespace Excel.Evaluation.Intermediate
{
    [TestFixture("vgsales.xlsx")]
    public class Exercise4
    {
        public Exercise4(string filename)
        {
            var directory = Environment.CurrentDirectory;
            directory += @"\..\..\..\..\Solution\";
            workbookFilename = directory + filename;
        }

        private IXLWorksheet sheet;
        private XLWorkbook workbook;
        private readonly string workbookFilename;

        [OneTimeSetUp]
        public void Setup()
        {
            workbook = new XLWorkbook(workbookFilename);
            sheet = workbook.Worksheets.Worksheet(1);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            workbook.Dispose();
        }

    }
}