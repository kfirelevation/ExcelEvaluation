using System;
using System.Collections.Generic;
using System.Text;

namespace SampleData
{
    interface ISampleStrategy
    {
        void CreateSample(string outputFilename, int outputSize = 200, int randomSeed = 0, IProgress<double> progressBar = null);
    }
}
