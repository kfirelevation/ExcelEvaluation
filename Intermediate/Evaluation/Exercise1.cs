using System;
using System.IO;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace Excel.Evaluation.Intermediate
{
    [TestFixture("vgsales.xlsx")]
    public class Exercise1
    {
        public Exercise1(string filename)
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

//        [Test]
        public void TestGlobalSales()
        {
            for (var row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var cell = cur_row.Cell((int)VideoGameSalesSheetCols.GlobalSales);

                Assert.IsTrue(cell.TryGetValue<double>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

                continue;

                double expected_val = 0;
                for (var j = (int) VideoGameSalesSheetCols.NaSales; j <= (int) VideoGameSalesSheetCols.OtherSales; j++)
                    expected_val += cur_row.Cell(j).GetDouble();

                Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");

                Assert.That(cell_actual_val, Is.EqualTo(expected_val).Within(0.01),
                    $"Cell {cell.Address} value should be {expected_val} but it is {cell_actual_val}");
            }
        }


        [Test]
        public void TestGlobalSalesNpoi()
        {
            IWorkbook workbook;
            //Write the stream data of workbook to the root directory
            using (var stream = new FileStream(workbookFilename, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = WorkbookFactory.Create(stream);
                stream.Close();
            }

            var npoi_sheet = workbook.GetSheetAt(0);

            int count = 0;
            var max_count = (double)npoi_sheet.LastRowNum;

            // first the range in the first column; 
            for (int row_idx = 1; row_idx <= npoi_sheet.LastRowNum; row_idx++)
            {
                var cur_row = npoi_sheet.GetRow(row_idx);
                var cell = cur_row.Cells[(int)VideoGameSalesSheetCols.GlobalSales - 1];
                count++;

                double expected_val = 0;
                for (var j = (int) VideoGameSalesSheetCols.NaSales; j <= (int) VideoGameSalesSheetCols.OtherSales; j++)
                    expected_val += cur_row.Cells[(j - 1)].NumericCellValue;

                Assert.IsTrue(cell.CellType == CellType.Formula, $"Cell {cell.Address} should be formula");

                Assert.That(cell.NumericCellValue, Is.EqualTo(expected_val).Within(0.01),
                    $"Cell {cell.Address} value should be {expected_val} but it is {expected_val}");
            }
        }
    }
}