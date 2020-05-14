using System;
using System.IO;
using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.Util;
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

        private ISheet sheet;
        IWorkbook workbook;
        private readonly string workbookFilename;

        [OneTimeSetUp]
        public void Setup()
        {
            using (var stream = new FileStream(workbookFilename, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = WorkbookFactory.Create(stream);
                stream.Close();
            }
            sheet = workbook.GetSheetAt(0);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            workbook.Close();
        }

        [Test]
        public void TestYearColumn()
        {
            // first the range in the first column; 
            for (int row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);
                var cell = cur_row.Cells[(int) VideoGameSalesSheetCols.Year - 1];

                Assert.IsTrue(cell.CellType == CellType.Numeric, $"Cell {cell.Address} should be numeric");
            }
        }

        [Test]
        public void TestTableSort()
        {
            // it is a dependency, so we run it before this. 
            TestYearColumn();

            var row_idx = 1;
            var cur_row = sheet.GetRow(row_idx);
            var prev_vgd = new VideoGameDetails()
            {
                Year = (int) cur_row.Cells[(int) VideoGameSalesSheetCols.Year - 1].NumericCellValue,
                Genre = cur_row.Cells[(int) VideoGameSalesSheetCols.Genre - 1].StringCellValue,
                Platform = cur_row.Cells[(int) VideoGameSalesSheetCols.Platform - 1].ToString(),
                Rank = (int) cur_row.Cells[(int) VideoGameSalesSheetCols.Rank - 1].NumericCellValue
            };

            for (row_idx = 2; row_idx <= sheet.LastRowNum; row_idx++)
            {
                cur_row = sheet.GetRow(row_idx);

                var cur_vgd = new VideoGameDetails()
                {
                    Year = (int)cur_row.Cells[(int)VideoGameSalesSheetCols.Year - 1].NumericCellValue,
                    Genre = cur_row.Cells[(int)VideoGameSalesSheetCols.Genre - 1].StringCellValue,
                    Platform = cur_row.Cells[(int)VideoGameSalesSheetCols.Platform - 1].ToString(),
                    Rank = (int)cur_row.Cells[(int)VideoGameSalesSheetCols.Rank - 1].NumericCellValue
                };

                Assert.GreaterOrEqual(cur_vgd, prev_vgd, $"Table is not ordered correctly see row {cur_row.RowNum}");

                prev_vgd = cur_vgd;
            }
        }

        [Test]
        public void TestGlobalSalesValues()
        {
            int count = 0;
            var max_count = (double)sheet.LastRowNum;

            // first the range in the first column; 
            for (int row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);
                var cell = cur_row.Cells[(int)VideoGameSalesSheetCols.GlobalSales - 1];
                count++;

                double expected_val = 0;
                for (var j = (int) VideoGameSalesSheetCols.NaSales; j <= (int) VideoGameSalesSheetCols.OtherSales; j++)
                    expected_val += cur_row.Cells[(j - 1)].NumericCellValue;

                Assert.IsTrue(cell.CellType == CellType.Formula, $"Cell {cell.Address} should be formula");
                Assert.IsTrue(cell.CachedFormulaResultType == CellType.Numeric,
                    $"Cell {cell.Address} formula result should be numeric");

                Assert.That(cell.NumericCellValue, Is.EqualTo(expected_val).Within(0.01),
                    $"Cell {cell.Address} value should be {expected_val} but it is {cell.NumericCellValue}");
            }
        }
    }
}