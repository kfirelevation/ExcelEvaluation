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

        protected SheetSampler(string masterFile, int sampleColumn = 0)
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