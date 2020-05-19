using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace Excel.Evaluation.Intermediate
{
    [TestFixture("vgsales.xlsx")]
    public class Exercise2
    {
        public Exercise2(string filename)
        {
            var directory = Environment.CurrentDirectory;
            directory += @"\..\..\..\..\Solution\";
            workbookFilename = directory + filename;
        }

        IWorkbook workbook;
        private readonly string workbookFilename;
        private VideoGameDetailsCollection rawData;

        [OneTimeSetUp]
        public void Setup()
        {
            using (var stream = new FileStream(workbookFilename, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = WorkbookFactory.Create(stream);
                stream.Close();
            }
            rawData = new VideoGameDetailsCollection(workbook.GetSheetAt(0));
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            workbook.Close();
        }

        [Test]
        public void TestStrategyTabTableValues()
        {
            TestTabSortFormula(2);
            var sheet = workbook.GetSheetAt(2);
            int prev_rank = 0;
            int actual_row_count = 0;
            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);

                var genre_cell = cur_row.Cells[(int) VideoGameSalesSheetCols.Genre - 1];
                Assert.AreEqual("Strategy", genre_cell.StringCellValue, $"row {cur_row.RowNum} contains game from Genre {genre_cell.StringCellValue} ");
                actual_row_count++;
            }
            var expect_count = rawData.Count(v => string.Compare(v.Genre, "Strategy", StringComparison.OrdinalIgnoreCase) == 0);
            Assert.AreEqual(expect_count, actual_row_count, $"rows count after filter should be {expect_count} but it is {actual_row_count} ");
        }

        [TestCase(1)]
        [TestCase(2)]
        public void TestTabSortFormula(int sheetIndex)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            var cell = sheet.GetRow(1).Cells[1];
            Assert.IsTrue(cell.CellType == CellType.Formula, $"Cell {cell.Address} should be formula");
            Assert.IsTrue(cell.CellFormula.Contains("SORT"), $"Cell {cell.Address} formula should include SORT");
            Assert.IsTrue(cell.CellFormula.Contains("FILTER"), $"Cell {cell.Address} formula should include SORT");
            Assert.IsTrue(cell.CellFormula.Contains("vgsales!"),
                $"Cell {cell.Address} formula should include a reference to vgsales Tab");
        }

        [Test]
        public void TestEightiesTabTableValues()
        {
            TestTabSortFormula(1);
            var sheet = workbook.GetSheetAt(1);
            int prev_rank = 0;
            int actual_row_count = 0;
            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);

                var rank_cell = cur_row.Cells[(int)VideoGameSalesSheetCols.Rank - 1];
                Assert.IsTrue(rank_cell.CachedFormulaResultType == CellType.Numeric, $" cell {rank_cell.CellType} should be numeric");
                var rank = (int)rank_cell.NumericCellValue;
                Assert.Greater(rank, prev_rank, $"table is not ordered correctly see row: {cur_row.RowNum}");

                var year_cell = cur_row.Cells[(int)VideoGameSalesSheetCols.Year - 1];
                Assert.IsTrue(year_cell.CachedFormulaResultType == CellType.Numeric, $" cell {year_cell.CellType} should be numeric");
                var year = (int)year_cell.NumericCellValue;
                Assert.That(year, Is.InRange(1980, 1989), $"row {cur_row.RowNum} contains game from year {year} which is not from the eighties");
                actual_row_count++;
            }
            var expect_count = rawData.Count(v => v.Year >= 1980 && v.Year <= 1989);
            Assert.AreEqual(expect_count, actual_row_count, $"rows count after filter should be {expect_count} but it is {actual_row_count} ");
        }
    }
}