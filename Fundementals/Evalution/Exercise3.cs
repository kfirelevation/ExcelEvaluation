using System;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using NUnit.Framework;


namespace ExcelFundenmentalsEvalution
{
    [TestFixture("BikeStoreSample.xlsx")]
    public class Exercise3
    {
        public Exercise3(string filename)
        {
            var directory = Environment.CurrentDirectory;
            directory += @"\..\..\..\..\Solution\";
            var file_path = directory + filename;
            workbook = new XLWorkbook(file_path);
            sheet = workbook.Worksheets.Worksheet(1);
        }

        private readonly IXLWorksheet sheet;
        private readonly XLWorkbook workbook;

        enum BikeStoreSheetCols
        {
            OrderId = 1,
            ItemId = 2,
            ProductName = 3,
            ModelYear = 4,
            ListPrice = 6,
            Quantity = 7, 
            LineTotal = 8, 
            Discount = 9,
            DiscountPerUnit = 10, 
            LineTotalAfterDiscount = 11,
            OrderDate = 12, 
            SalesMan = 13, 
            OldModel = 16
        };

        [TearDown]
        public void ClosedWorkbook()
        {
            workbook.Dispose();
        }

        [Test]
        public void TestColumnOldModel()
        {
            for (int row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var cell = cur_row.Cell((int)BikeStoreSheetCols.OldModel);

                // Format is not required for this column.

                // test if the value is string. 
                Assert.IsTrue(cell.TryGetValue<string>(out var cell_actual_val));

                var cell_expected_val =
                    cur_row.Cell((int) BikeStoreSheetCols.OrderDate).GetValue<DateTime>().Year >
                    cur_row.Cell((int) BikeStoreSheetCols.ModelYear).GetValue<int>()
                        ? "YES"
                        : "NO";

                Assert.That(cell_expected_val, Is.EqualTo(cell_actual_val).IgnoreCase,  $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");

                // test whether this cell is actually a formula.
                Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");

                // test whether the formula contains IF. 
                StringAssert.Contains("IF", cell.FormulaR1C1,
                    $"Cell {cell.Address} formula should include conditional");

                // test whether the formula referencing correct rows and columns. 
                StringAssert.Contains("-4", cell.FormulaR1C1,
                    $"Cell {cell.Address} formula should reference column order_date ");
                StringAssert.Contains("-12", cell.FormulaR1C1,
                    $"Cell {cell.Address} formula should reference column model_year");
                Assert.IsTrue(Regex.Matches(cell.FormulaA1, cell.Address.RowNumber.ToString()).Count() == 2, $"Cell {cell.Address} is not referencing correct rows.  ");
            }
        }

        [Test]
        public void TestColumnOldModelTotal()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.OldModel);
            int cell_expected_val = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.OldModel);

                // Format is not required for this column.


                cell_expected_val +=
                    cur_row.Cell((int)BikeStoreSheetCols.OrderDate).GetValue<DateTime>().Year >
                    cur_row.Cell((int)BikeStoreSheetCols.ModelYear).GetValue<int>()
                        ? 1
                        : 0;
            }

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.OldModel);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<int>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val), $"Cell {cell.Address} value should be {cell_actual_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("COUNTIF", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNTIF function");
            string[] expected_reference = 
            {
                    $"R[-{row_idx - 2}]C:R[-1]C",
                    $"R[-1]C:R[-{row_idx - 2}]C"
            };
            Assert.IsTrue(expected_reference.Any(s => cell.FormulaR1C1.Contains(s)), 
                $"Cell {cell.Address} formula is not referencing the correct range");
        }
    }
}