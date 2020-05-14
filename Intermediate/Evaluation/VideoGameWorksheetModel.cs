
using System;
using NPOI.OpenXmlFormats.Dml.Diagram;

namespace Excel.Evaluation.Intermediate
{
    enum VideoGameSalesSheetCols
    {
        Rank = 1,
        Name = 2,
        Platform = 3,
        Year = 4,
        Genre = 5,
        Publisher = 6,
        NaSales = 7,
        EuSales = 8,
        JpSales = 9,
        OtherSales = 10,
        GlobalSales = 11,
    };

    class VideoGameDetails : IComparable
    {
        public int Year { get; set; }
        public string Genre { get; set; }
        public string Platform { get; set; }
        public int Rank { get; set; }

        public int CompareTo(object? obj)
        {
            var other = obj as VideoGameDetails;
            if (other == null)
                return -1;

            var year_comparison = Year.CompareTo(other.Year);
            if (year_comparison != 0)
                return year_comparison;
            var genre_comparison = string.Compare(Genre, other.Genre, StringComparison.Ordinal);
            if (genre_comparison != 0)
                return genre_comparison;
            var platform_comparison = string.Compare(Platform, other.Platform, StringComparison.Ordinal);
            if (platform_comparison != 0)
                return platform_comparison;
            return Rank.CompareTo(other.Rank);
        }
    }
}
