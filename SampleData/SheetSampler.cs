using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel.Fundemenetals
{
    public abstract class SheetSampler
    {
        protected readonly string MasterFilename;
        protected readonly int SampleColumn;

        private readonly Dictionary<int, int> valuesCount;

        public SheetSampler(string masterFile, int sampleColumn = 0)
        {
            MasterFilename = masterFile;
            this.SampleColumn = sampleColumn;
            valuesCount = new Dictionary<int, int>();
            BuildColIndex();
        }

        Dictionary<int, int> CreateRowsSample(int outputSize, int randomSeed)
        {
            var d = new Random(randomSeed);
            var selected_count = 0;
            return valuesCount.OrderBy(e => d.Next()).ToArray()
                .TakeWhile(e => (selected_count += e.Value) <= outputSize).ToDictionary(x => x.Key, x => x.Value);
        }

        public void CreateSample(string optionsOutputFile, int linesCount = 200, int optionsSeed  = 0, ProgressBar progressBar = null)
        {
            var samples = CreateRowsSample(linesCount, optionsSeed);
            ExecuteSampler(optionsOutputFile, samples, progressBar);
        }

        public abstract void ExecuteSampler(string outputFilename, Dictionary<int, int> selectedRow, IProgress<double> progressBar = null);

        public void CreateSampleOld(string outputFilename, int outputSize = 200, int randomSeed = 0)
        {
            //Write the stream data of workbook to the root directory
            var source_file_stream = new FileStream(MasterFilename, FileMode.Open);
            var workbook = WorkbookFactory.Create(source_file_stream);
            var destination_workbook = new XSSFWorkbook();

            for (short i = 0; i < workbook.NumberOfFonts; i++)
            {
                var dst_font = destination_workbook.NumberOfFonts > i ? destination_workbook.GetFontAt(i) : destination_workbook.CreateFont();
                var src_font = workbook.GetFontAt(i);
                dst_font.IsBold = src_font.IsBold;
                dst_font.FontName = src_font.FontName;
                dst_font.FontHeight = src_font.FontHeight;
                Console.WriteLine(workbook.GetFontAt(i).ToString());
            }

            Console.WriteLine();

            for (short i = 0; i < destination_workbook.NumberOfFonts; i++)
            {
                Console.WriteLine(destination_workbook.GetFontAt(i).ToString());
            }

            var sheet = workbook.GetSheetAt(0);
            var d = new Random(randomSeed);
            var selected_count = 0;
            var shuffled = valuesCount.OrderBy(e => d.Next()).ToArray()
                .TakeWhile(e => (selected_count += e.Value) <= outputSize).ToDictionary(x => x.Key, x => x.Value);

            //here, we must insert at least one sheet to the workbook. otherwise, Excel will say 'data lost in file'
            //So we insert three sheet just like what Excel does
            var dest_sheet = destination_workbook.CreateSheet(sheet.SheetName);
            var dest_row = 1;

            CopyHeaderRow(dest_sheet, sheet);

            // first the range in the first column; 
            for (int src_row = 1; sheet.GetRow(src_row)?.Cells[SampleColumn].CellType == CellType.Numeric ; src_row++)
            {
                var cell_val = (int) sheet.GetRow(src_row).Cells[SampleColumn].NumericCellValue;
                if (!shuffled.ContainsKey(cell_val))
                    continue;
                CopyRow(dest_sheet, dest_row, sheet, src_row);
                dest_row++;
            }

            // get last row on destination and autosize the worksheet. 
            var _row = dest_sheet.GetRow(dest_row - 1); // 

            for (var i_col = 0; i_col < _row.Cells.Count; i_col++)
                dest_sheet.AutoSizeColumn(i_col);

            source_file_stream.Close();

            //Write the stream data of workbook to the root directory
            var dst_file = new FileStream(outputFilename, FileMode.Create);
            destination_workbook.Write(dst_file);
            dst_file.Close();
        }

        private void CopyHeaderRow(ISheet dst, ISheet src)
        {
            CopyRow(dst, 0, src, 0);

            var row = dst.GetRow(0);
            for (var i_col = 0; i_col < row.Cells.Count; i_col++)
            {
                var style = row.Cells[i_col].CellStyle;
                Console.WriteLine(style.FontIndex);
            }
        }

        private void CopyRow(ISheet dst, int dstRowIndex, ISheet src, int srcRowIndex)
        {
            var src_row = src.GetRow(srcRowIndex);
            var dst_row = dst.CreateRow(dstRowIndex);

            //             dst.Workbook.FindFont(src_font.IsBold, src_font.FontHeight, src_font.IsItalic, src_font.FontName, src_font

            for (var i_col = 0; i_col < src_row.Cells.Count; i_col++)
            {
                var src_cell = src_row.Cells[i_col];
                var dst_cell = dst_row.CreateCell(i_col);

                switch (src_row.Cells[i_col].CellType)
                {
                    case CellType.Numeric:
                        dst_cell.SetCellValue(src_cell.NumericCellValue);
                        break;
                    case CellType.String:
                        dst_cell.SetCellValue(src_cell.StringCellValue);
                        break;
                }

                var src_font = src_cell.CellStyle.GetFont(src.Workbook);
                dst_cell.CellStyle.SetFont(dst.Workbook.GetFontAt(src_cell.CellStyle.FontIndex));
            }
        }

        private void BuildColIndex()
        {
            //Write the stream data of workbook to the root directory
            var source_file_stream = new FileStream(MasterFilename, FileMode.Open);
            var workbook = WorkbookFactory.Create(source_file_stream, ImportOption.All);
            source_file_stream.Close();
            var sheet = workbook.GetSheetAt(0);

            var row = 1;
            // first the range in the first column; 
            while (sheet.GetRow(row)?.Cells[SampleColumn].CellType == CellType.Numeric)
            {
                var cell_val = (int)sheet.GetRow(row).Cells[SampleColumn].NumericCellValue;
                if (!valuesCount.ContainsKey(cell_val))
                    valuesCount.Add(cell_val, 1);
                else
                    valuesCount[cell_val]++;
                row++;
            }
        }

    }
}