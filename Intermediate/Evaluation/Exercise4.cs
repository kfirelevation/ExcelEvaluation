using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using CellType = NPOI.SS.UserModel.CellType;

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
            sheet = workbook.GetSheetAt(4);
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
        public void TestPlatformColumn()
        {
            var actual_row_count = 0;
            var platform_values = new List<string>();

            var cur_row = sheet.GetRow(1);
            var cell = cur_row.Cells[0];

            AssertFormula(cell, new[]
            {
                "UNIQUE"
            });

            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                cur_row = sheet.GetRow(row_idx);
                cell = cur_row.Cells[0];

                switch (cell.CachedFormulaResultType)
                {
                    case CellType.String:
                        platform_values.Add(cell.StringCellValue);
                        break;
                    case CellType.Numeric:
                        platform_values.Add(cell.NumericCellValue.ToString());
                        break;
                    default:
                        Assert.True(false, $"cell {cell.Address} type should be string or numeric");
                        break;
                }
                actual_row_count++;
            }
            var actual_distinct_count = platform_values.Distinct().Count();
            Assert.AreEqual(actual_distinct_count, actual_row_count, "Some Genres occurs more than once");

            var expect_distinct_count = rawData.Select(x => x.Platform).Distinct().Count();
            Assert.AreEqual(expect_distinct_count, actual_distinct_count, $"wrong number of platform found {expect_distinct_count} but there is {actual_distinct_count}");
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
        public void TestYearsHeader()
        {
            var actual_row_count = 0;
            var values = new List<int>();

            var cur_row = sheet.GetRow(0);
            var cell = cur_row.Cells[1];

            AssertFormula(cell, new[]
            {
                "UNIQUE", "TRANSPOSE"
            });

            int prev_year = -1;

            for (var idx = 1; idx < cur_row.Cells.Count; idx++)
            {
                cell = cur_row.Cells[idx];
                Assert.IsTrue(cell.CachedFormulaResultType == CellType.Numeric , $"cell {cell.Address} type should be string or numeric");
                var cur_cell_year = (int)cell.NumericCellValue;
                values.Add(cur_cell_year);
                Assert.IsTrue(cur_cell_year > prev_year, "The years header is not order correctly. It should be order ascending ");
                prev_year = cur_cell_year;
            }
            var actual_distinct_count = values.Distinct().Count();
            Assert.AreEqual(actual_distinct_count, values.Count, "Some years appears more than once");

            var expect_distinct_count = rawData.Select(x => x.Year).Distinct().Count();
            var years_array = rawData.Select(x => x.Year).Distinct().ToArray();
            foreach (var y in years_array)
            {
                Console.WriteLine(y);    
            }
            Assert.AreEqual(expect_distinct_count, actual_distinct_count, $"wrong number of years found {expect_distinct_count} but there is {actual_distinct_count}");
        }

        [Test]
        public void TestTotalIncomes()
        {
            TestYearsHeader();
            TestPlatformColumn();

            var x_values = new List<int>();
            var cur_row = sheet.GetRow(0);

            for (var idx = 1; idx < cur_row.Cells.Count; idx++)
                x_values.Add((int) cur_row.Cells[idx].NumericCellValue);

            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                cur_row = sheet.GetRow(row_idx);
                int cur_col = 1;
                foreach(var x_val in x_values)
                {
                    var cell = cur_row.Cells[0];
                    var y_value = cell.CachedFormulaResultType == CellType.Numeric
                        ? cell.NumericCellValue.ToString()
                        : cell.StringCellValue;
                    var sum_income = cur_row.Cells[cur_col];
                    AssertFormula(sum_income, new[] {"SUMIF"});

                    Assert.IsTrue(sum_income.CachedFormulaResultType == CellType.Numeric,
                        $"cell {sum_income.Address} type should be numeric");

                    var expected_value = rawData.Where(d => d.Platform == y_value && d.Year == x_val).Sum(d => d.GlobalSales);
                    var actual_val = sum_income.NumericCellValue;
                    Assert.AreEqual(expected_value, actual_val, 0.01,
                        $"Platform {y_value} on year {x_val} sum is {sum_income} but it should be {expected_value}");
                    cur_col++;
                }
            }
        }


        [Test]
        public void TestLineChart()
        {
            var document = SpreadsheetDocument.Open(workbookFilename, true);
            var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where
                (s => s.Name == "PlatformByYear");
            if (!sheets.Any())
            {
                // The specified worksheet does not exist.
                return;
            }
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

            var charts = worksheetPart.DrawingsPart.ChartParts;

            // Add a new drawing to the worksheet.
            Assert.AreEqual(charts.Count(), 1, "Worksheet should include two pie charts");
            foreach (var chart in charts )
            {
                Assert.AreEqual(1, chart.ChartSpace.Descendants<PlotArea>().First().Descendants<LineChart>().Count(),
                    "Chart style is not PieChart");
            }
        }
    }
}