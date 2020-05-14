using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;

namespace Excel.Evaluation.Fundamentals
{
    [TestFixture("BikeStoreSample.xlsx")]
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

        [Test]
        public void TestYearsTotals()
        {
            var totals_sheet = workbook.Worksheets.Worksheet(2);

            var expected_vals = new Dictionary<int, double>();
            for (int row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var order_date = cur_row.Cell((int) BikeStoreSheetCols.OrderDate).GetDateTime();
                var amount = cur_row.Cell((int) BikeStoreSheetCols.LineTotalAfterDiscount).GetDouble();

                if (!expected_vals.ContainsKey(order_date.Year))
                    expected_vals.Add(order_date.Year, 0);
                expected_vals[order_date.Year] += amount;
            }

            var first_row = totals_sheet.Row(1);
            var total_row = totals_sheet.Row(2);

            for (int col = 2; ; col++)
            {
                if (!first_row.Cell(col).TryGetValue<int>(out var header_cell_val))
                    break;

                var total_cell = total_row.Cell(col);
                Assert.IsTrue(total_cell.HasFormula, $"Cell {total_cell.Address} should be formula");
                Assert.IsTrue(total_cell.TryGetValue<double>(out var total_actual_value), $"Cell {total_cell.Address} value is not a number");

                // test whether the formula contains IF. 
                StringAssert.Contains("SUMIF", total_cell.FormulaR1C1,
                    $"Cell {total_cell.Address} formula should include SUMIF");

                Assert.That(total_actual_value, Is.EqualTo(expected_vals[header_cell_val]).Within(0.01),
                    $"Cell {total_cell.Address} value should be {expected_vals[header_cell_val]} but it is {total_actual_value}");

            }
        }

        [Test]
        public void TestMonthTotals()
        {
            var totals_sheet = workbook.Worksheets.Worksheet(3);

            var expected_vals = new Dictionary<int, Dictionary<int, double>>();
            for (int row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var order_date = cur_row.Cell((int)BikeStoreSheetCols.OrderDate).GetDateTime();
                var amount = cur_row.Cell((int)BikeStoreSheetCols.LineTotalAfterDiscount).GetDouble();

                if (!expected_vals.ContainsKey(order_date.Year))
                    expected_vals.Add(order_date.Year, new Dictionary<int, double>());
                if (!expected_vals[order_date.Year].ContainsKey(order_date.Month))
                    expected_vals[order_date.Year].Add(order_date.Month, 0);

                expected_vals[order_date.Year][order_date.Month] += amount;
            }

            var first_row = totals_sheet.Row(1);
            for (int row_idx = 2; row_idx < totals_sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var total_row = totals_sheet.Row(2);
                for (int col = 2;; col++)
                {
                    if (!first_row.Cell(col).TryGetValue<int>(out var header_cell_val))
                        break;

                    if (!total_row.Cell(1).TryGetValue<int>(out var month_cell_val))
                        break;

                    var total_cell = total_row.Cell(col);
                    Assert.IsTrue(total_cell.HasFormula, $"Cell {total_cell.Address} should be formula");
                    Assert.IsTrue(total_cell.TryGetValue<double>(out var total_actual_value),
                        $"Cell {total_cell.Address} value is not a number");

                    // test whether the formula contains IF. 
                    StringAssert.Contains("SUMIF", total_cell.FormulaR1C1,
                        $"Cell {total_cell.Address} formula should include SUMIF");

                    Assert.IsTrue(total_cell.Style.NumberFormat.Format.Contains("$"), $"cell format {total_cell.Address} should be $");

                    Assert.That(total_actual_value, Is.EqualTo(expected_vals[header_cell_val][month_cell_val]).Within(0.01),
                        $"Cell {total_cell.Address} value should be {expected_vals[header_cell_val][month_cell_val]} but it is {total_actual_value}");
                }
            }
        }
    }
}