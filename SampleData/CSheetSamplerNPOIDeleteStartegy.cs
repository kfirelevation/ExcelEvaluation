using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel.Fundemenetals;
using NPOI.SS.UserModel;

namespace SampleData
{
    internal class CSheetSamplerNpoiDeleteStrategy : SheetSampler
    {

        public CSheetSamplerNpoiDeleteStrategy(string masterFile, int sampleColumn = 0) : base(masterFile, sampleColumn)
        {
            
        }

        public override void ExecuteSampler(string outputFilename, Dictionary<int, int> selectedRow, IProgress<double> progressBar)
        {
            IWorkbook workbook;
            //Write the stream data of workbook to the root directory
            using (var stream = new FileStream(MasterFilename, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = WorkbookFactory.Create(stream);
                stream.Close();
            }

            var sheet = workbook.GetSheetAt(0);

            int count = 0;
            var max_count = (double)sheet.LastRowNum;

            // first the range in the first column; 
            for (int row_idx = sheet.LastRowNum; row_idx >= 1; row_idx--)
            {
                var cur_row = sheet.GetRow(row_idx);
                var cell_val = (int)cur_row.Cells[SampleColumn].NumericCellValue;
                if (selectedRow.ContainsKey(cell_val))
                    continue;
                RemoveRow(sheet, row_idx);
                count++;
                progressBar?.Report(count / max_count);
            }

            //Write the stream data of workbook to the root directory
            File.Delete(outputFilename);
            using (var stream = new FileStream(outputFilename, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                workbook.Write(stream);
                stream.Close();
            }
        }

        private static void RemoveRow(ISheet sheet, int rowIndex)
        {
            var last_row_num = sheet.LastRowNum;
            if (rowIndex >= 0 && rowIndex < last_row_num)
            {
                sheet.ShiftRows(rowIndex + 1, last_row_num, -1);
            }

            if (rowIndex != last_row_num)
                return;

            var removing_row = sheet.GetRow(rowIndex);
            if (removing_row != null)
            {
                sheet.RemoveRow(removing_row);
            }
        }
    }
}
