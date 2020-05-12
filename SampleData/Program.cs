using CommandLine;
using System;
using System.Collections.Generic;
using System.IO;
using SampleData;

namespace Excel.Fundemenetals
{
    class Program
    {
        public class Options
        {
            [Option('s', "source", Required = true, HelpText = "Master file to sample from ")]
            public string MasterFile { get; set; }

            [Option('k', "keycol", Required = false, HelpText = "The column that will be used as key for deleting", Default = 0)]
            public int KeyColumn { get; set; }

            [Option('o', "output", Required = true, HelpText = "File to Write to")]
            public string OutputFile { get; set; }

            [Option('l', "lines", Required = false, HelpText = "Number of lines for each sample from master", Default = 100)]
            public int LinesCount { get; set; }

            [Option('s', "seed", Required = false, HelpText = "The Random seed from which the sample will be created", Default = 0)]
            public int Seed { get; set; }

            [Option('n', "no_of_samples", Required = false, HelpText = "How many samples files to create", Default = 1)]
            public int SamplesDatasetsCount { get; set; }
        }

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(RunOptions).WithNotParsed(HandleParseError);
        }

        private static void HandleParseError(IEnumerable<Error> errors)
        {
            foreach (var e in errors)
                Console.WriteLine(e.StopsProcessing);
        }

        private static void RunOptions(Options options)
        {
            SheetSampler sampler = new CSheetSamplerCloseXmlDelete(options.MasterFile, options.KeyColumn);

            var filename = Path.GetFileNameWithoutExtension(options.OutputFile);
            var extension = Path.GetExtension(options.OutputFile);
            for (int i = 1; i <= options.SamplesDatasetsCount; i++)
            {
                Console.Write($"Sample {i}:");
                sampler.CreateSample($"{filename}{i:00}{extension}", options.LinesCount, options.Seed + i,
                    new ProgressBar());
                Console.WriteLine();
            }
        }
    }
}
