using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace Excel.Evaluation.Fundamentals
{
    [TestFixture("BikeStoreSample.xlsx")]
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

        [Test]
        public void TestColumnListPrice()
        {
            for (var row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var cell = cur_row.Cell((int) BikeStoreSheetCols.ListPrice);
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("$"), $"cell format {cell.Address} should be $" );
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("0.00"), $"cell format {cell.Address} should be 2 digit accurate");
            }
        }

        [Test]
        public void TestColumnDiscount()
        {
            for (var row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var cell = cur_row.Cell((int)BikeStoreSheetCols.Discount);

                // cell number format 9 is %0  
                // see : https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
                Assert.IsTrue(cell.Style.NumberFormat.NumberFormatId == 9 || cell.Style.NumberFormat.Format.Contains("%"), $"cell format {cell.Address} should be %");
            }
        }

        [Test]
        public void TestLineTotal()
        {
            for (var row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var cell = cur_row.Cell((int) BikeStoreSheetCols.LineTotal);

                // test the number format
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("$"), $"cell format {cell.Address} should be $");
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("0.00"),
                    $"cell format {cell.Address} should be 2 digit accurate");

                // test the value
                var cell_actual_val = cell.GetValue<double>();
                var cell_expected_val = cur_row.Cell((int)BikeStoreSheetCols.Quantity).GetValue<double>() *
                                        cur_row.Cell((int)BikeStoreSheetCols.ListPrice).GetValue<double>();
                Assert.AreEqual(cell_expected_val, cell_actual_val, $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");


                // test the formula
                Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should contain formula");
                Assert.IsTrue(cell.FormulaR1C1.Equals("RC[-2]*RC[-1]") || cell.FormulaR1C1.Equals("RC[-1]*RC[-2]"), $"Cell {cell.Address} formula is wrong");
            }
        }

        [Test]
        public void TestLineTotalAfterDiscount()
        {
            for (var row_idx = 2; row_idx < sheet.LastRowUsed().RowNumber(); row_idx++)
            {
                var cur_row = sheet.Row(row_idx);
                var cell = cur_row.Cell((int)BikeStoreSheetCols.LineTotalAfterDiscount);

                // test the number format
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("$"), $"cell format {cell.Address} should be $");
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("0.00"),
                    $"cell format {cell.Address} should be 2 digit accurate");

                // test the value
                var cell_actual_val = cell.GetValue<double>();
                var cell_expected_val = cur_row.Cell((int) BikeStoreSheetCols.Quantity).GetValue<double>() *
                                        cur_row.Cell((int) BikeStoreSheetCols.ListPrice).GetValue<double>() *
                                        (1 - cur_row.Cell((int) BikeStoreSheetCols.Discount).GetValue<double>());
                Assert.That(cell_expected_val, Is.EqualTo(cell_actual_val).Within(0.01),  $"Cell {cell.Address} value should be {cell_expected_val} but it is {cell_actual_val}");

                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("$"), $"cell format {cell.Address} should be $");
                Assert.IsTrue(cell.Style.NumberFormat.Format.Contains("0.00"),
                    $"cell format {cell.Address} should be 2 digit accurate");

                // because there could be many variations of this formula, we just test if it's a real formula. 
                Assert.IsTrue(cell.HasFormula, $"Cell {cell.Address} should contain formula");
            }
        }
    }
}