using System;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using NUnit.Framework;


namespace ExcelFundamentalsEvalution
{
    [TestFixture("BikeStoreSample.xlsx")]
    public class Exercise2
    {
        public Exercise2(string filename)
        {
            var directory = Environment.CurrentDirectory;
            directory += @"\..\..\..\..\Solution\";
            var file_path = directory + filename;
            workbook = new XLWorkbook(file_path);
            sheet = workbook.Worksheets.Worksheet(1);
        }

        private readonly IXLWorksheet sheet;
        private readonly XLWorkbook workbook;

        [TearDown]
        public void TearDown()
        {
//            workbook.Dispose();
        }

        [Test]
        public void TestPotentialIncome()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotal);
            double cell_expected_val = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotal);
                cell_expected_val += cell.GetValue<double>();
            }

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotal);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<double>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is a number 
            Assert.IsTrue(cell.Style.Font.Bold, $"Font in cell {cell.Address} should be bold");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val), $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("SUM", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNT function");
            string[] expected_reference =
            {
                $"R[-{row_idx - 2}]C:R[-1]C",
                $"R[-1]C:R[-{row_idx - 2}]C"
            };
            Assert.IsTrue(expected_reference.Any(s => cell.FormulaR1C1.Contains(s)),
                $"Cell {cell.Address} formula is not referencing the correct range");
        }

        [Test]
        public void TestOrderDetailsCount()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.Quantity);
            int cell_expected_val = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.Quantity);
                cell_expected_val += cell.GetValue<int>();
            }

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.Quantity);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<int>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is a number 
            Assert.IsTrue(cell.Style.Font.Bold, $"Font in cell {cell.Address} should be bold");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val), $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("SUM", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNT function");
            string[] expected_reference = 
            {
                    $"R[-{row_idx - 2}]C:R[-1]C",
                    $"R[-1]C:R[-{row_idx - 2}]C"
            };
            Assert.IsTrue(expected_reference.Any(s => cell.FormulaR1C1.Contains(s)), 
                $"Cell {cell.Address} formula is not referencing the correct range");
        }

        [Test]
        public void TestOrderedItems()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.ItemId);
            int cell_expected_val = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.ItemId);

                cell_expected_val++;
            }

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.ItemId);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<int>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is a number 
            Assert.IsTrue(cell.Style.Font.Bold, $"Font in cell {cell.Address} should be bold");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val), $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("COUNT", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNT function");
            string[] expected_reference =
            {
                $"R[-{row_idx - 2}]C:R[-1]C",
                $"R[-1]C:R[-{row_idx - 2}]C"
            };
            Assert.IsTrue(expected_reference.Any(s => cell.FormulaR1C1.Contains(s)),
                $"Cell {cell.Address} formula is not referencing the correct range");
        }

        [Test]
        public void TestAverageCostForOrderedItem()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.ListPrice);
            double cell_expected_val = 0;
            double sum = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.ListPrice);
                sum += cell.GetValue<double>();
            }

            cell_expected_val = sum / (row_idx - 2);

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.ListPrice);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<double>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is a number 
            Assert.IsTrue(cell.Style.Font.Bold, $"Font in cell {cell.Address} should be bold");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val).Within(0.001), $"Cell {cell.Address} value should be {cell_actual_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("AVERAGE", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNT function");
            string[] expected_reference =
            {
                $"R[-{row_idx - 2}]C:R[-1]C",
                $"R[-1]C:R[-{row_idx - 2}]C"
            };
            Assert.IsTrue(expected_reference.Any(s => cell.FormulaR1C1.Contains(s)),
                $"Cell {cell.Address} formula is not referencing the correct range");
        }

        [Test]
        public void TestAverageDiscountPerUnit()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.DiscountPerUnit);
            double cell_expected_val = 0;
            double sum = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.DiscountPerUnit);
                sum += cell.GetValue<double>();
            }

            cell_expected_val = sum / (row_idx - 2);

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.DiscountPerUnit);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<double>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is a number 
            Assert.IsTrue(cell.Style.Font.Bold, $"Font in cell {cell.Address} should be bold");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val).Within(0.001), $"Cell {cell.Address} value should be {cell_actual_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("AVERAGE", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNT function");
            string[] expected_reference =
            {
                $"R[-{row_idx - 2}]C:R[-1]C",
                $"R[-1]C:R[-{row_idx - 2}]C"
            };
            Assert.IsTrue(expected_reference.Any(s => cell.FormulaR1C1.Contains(s)),
                $"Cell {cell.Address} formula is not referencing the correct range");
        }

        [Test]
        public void TestTotalIncome()
        {
            var last_row = sheet.LastRowUsed().RowNumber();
            var row_idx = 2;
            var cur_row = sheet.Row(row_idx);
            var cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotalAfterDiscount);
            double cell_expected_val = 0;

            for (; row_idx < last_row; row_idx++)
            {
                cur_row = sheet.Row(row_idx);
                cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotalAfterDiscount);
                cell_expected_val += cell.GetValue<double>();
            }

            cur_row = sheet.Row(row_idx);
            cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotalAfterDiscount);

            // test if the value is a number 
            Assert.IsTrue(cell.TryGetValue<double>(out var cell_actual_val), $"Cell {cell.Address} value is not a number");

            // test if the value is a number 
            Assert.IsTrue(cell.Style.Font.Bold, $"Font in cell {cell.Address} should be bold");

            // test if the value is correct. 
            Assert.That(cell_actual_val, Is.EqualTo(cell_expected_val), $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");

            // because there could be many variations of this formula, we just test if it's a real formula. 
            Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should be formula");
            StringAssert.Contains("SUM", cell.FormulaR1C1,
                $"Cell {cell.Address} formula should include COUNT function");
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