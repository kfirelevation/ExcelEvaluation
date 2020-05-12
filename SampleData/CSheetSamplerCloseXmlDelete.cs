using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Excel.Fundemenetals;


namespace SampleData
{
    internal class CSheetSamplerCloseXmlDelete : SheetSampler
    {
        public CSheetSamplerCloseXmlDelete(string masterFile, int sampleColumn = 0) : base(masterFile, sampleColumn)
        {
            
        }

        public override void ExecuteSampler(string outputFilename, Dictionary<int, int> selectedRow, IProgress<double> progressBar)
        {

            File.Copy(MasterFilename, outputFilename, true);

            var workbook = new XLWorkbook(outputFilename);
            var sheet = workbook.Worksheets.Worksheet(1);

            int count = 0;
            var max_count = sheet.LastRowUsed().RowNumber();

            for (int row_idx = max_count; row_idx >= 2; row_idx--)
            {
                var cur_row = sheet.Row(row_idx);
                var cell_val = cur_row.Cell(SampleColumn + 1).GetValue<int>();
                count++;
                progressBar?.Report(count / (double) max_count);
                if (selectedRow.ContainsKey(cell_val))
                    continue;
                cur_row.Delete();
            }

            workbook.Save();
        }
    }
}
