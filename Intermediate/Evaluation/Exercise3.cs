using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Office.CustomUI;

namespace Excel.Evaluation.Intermediate
{
    [TestFixture("vgsales.xlsx")]
    public class Exercise3
    {
        public Exercise3(string filename)
        {
            var directory = Environment.CurrentDirectory;
            directory += @"\..\..\..\..\Solution\";
            workbookFilename = directory + filename;
        }

        IWorkbook workbook;
        private ISheet sheet;
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
            sheet = workbook.GetSheetAt(3);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            workbook.Close();
        }

        //[Test]
        public void TestStrategyTabTableValues()
        {
//            TestTabSortFormula(2);
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


        [Test]
        public void TestGenresColumn()
        {
            int prev_rank = 0;
            int actual_row_count = 0;
            var genere_values = new List<string>();
            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);
                var genre_cell = cur_row.Cells[0];
                Assert.IsTrue(genre_cell.CellType == CellType.String,
                    $"cell {genre_cell.Address} type should be string");
                genere_values.Add(genre_cell.StringCellValue);

                actual_row_count++;
            }
            var actual_distinct_count = genere_values.Distinct().Count();
            Assert.AreEqual(actual_distinct_count, actual_row_count, "Some Genres occurs more than once");

            var expect_distinct_count = rawData.Select(x => x.Genre).Distinct().Count();
            Assert.AreEqual(expect_distinct_count, actual_distinct_count, $"wrong number of genre found {expect_distinct_count} but there is {actual_distinct_count}");
        }


        private void AssertFormula(ICell cell, string[] vals)
        {
            Assert.IsTrue(cell.CellType == CellType.Formula, $"Cell {cell.Address} should be formula");
            Assert.IsTrue(cell.CellFormula.Contains("vgsales!"),
                $"Cell {cell.Address} formula should include a reference to vgsales Tab");
            foreach (var v in vals)
            {
                StringAssert.Contains(v, cell.CellFormula);
            }
        }

        [Test]
        public void TestGamesCount()
        {
            TestGenresColumn();

            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);

                var genre = cur_row.Cells[0].StringCellValue;
                var games_count = cur_row.Cells[1];
                AssertFormula(games_count, new[]{"COUNTIF"});

                Assert.IsTrue(games_count.CachedFormulaResultType == CellType.Numeric,
                    $"cell {games_count.Address} type should be numeric");

                var expected_value = rawData.Count(d => d.Genre.Equals(genre));
                var actual_val = (int) games_count.NumericCellValue;
                Assert.AreEqual(expected_value, actual_val, $"Genre {genre} count is {games_count} but it should be {expected_value}");
            }
        }

        [Test]
        public void TestTotalIncome()
        {
            TestGenresColumn();

            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);

                var genre = cur_row.Cells[0].StringCellValue;
                var games_count = cur_row.Cells[2];
                AssertFormula(games_count, new[] { "SUMIF" });

                Assert.IsTrue(games_count.CachedFormulaResultType == CellType.Numeric,
                    $"cell {games_count.Address} type should be numeric");

                var expected_value = rawData.Where(d => d.Genre == genre).Sum(d => d.GlobalSales);
                var actual_val = (double)games_count.NumericCellValue;
                Assert.AreEqual(expected_value, actual_val, 0.01, $"Genre {genre} count is {games_count} but it should be {expected_value}");
            }
        }

        public void TestPieChart()
        {
            var w = new ClosedXML.Excel.XLWorkbook();
            var sheet = w.Worksheets.Worksheet(1);

        }
    }
}