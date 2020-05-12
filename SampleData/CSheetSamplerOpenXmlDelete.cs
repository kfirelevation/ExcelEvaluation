using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel.Fundemenetals;


namespace SampleData
{
    internal class CSheetSamplerOpenXmlDelete : SheetSampler
    {
        public CSheetSamplerOpenXmlDelete(string masterFile, int sampleColumn = 0) : base(masterFile, sampleColumn)
        {
            
        }

        public override void ExecuteSampler(string outputFilename, Dictionary<int, int> selectedRow, IProgress<double> progressBar)
        {

            File.Copy(MasterFilename, outputFilename, true);

            var spreadsheet = SpreadsheetDocument.Open(outputFilename, true);

            var sheet = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

            int count = 0;
            var max_count = sheet.Count();


            for (int row_idx = max_count-1; row_idx >= 1; row_idx--)
            {
                var cur_row = sheet.ElementAt(row_idx);
                var cell_val = int.Parse(cur_row.ElementAt(SampleColumn).InnerText);
                count++;
                progressBar?.Report(count / (double)max_count);
                if (selectedRow.ContainsKey(cell_val))
                    continue;
                cur_row.RemoveAllChildren();
                cur_row.Remove();
            }

            spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
            spreadsheet.Save();
        }
    }
}
