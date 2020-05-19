
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Presentation;
using NPOI.SS.UserModel;
using NUnit.Framework;

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
        public string Publisher { get; set; }
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

    class VideoGameDetailsCollection : IEnumerable<VideoGameDetails>
    {
        private readonly IDictionary<int, VideoGameDetails> collection; 
        public VideoGameDetailsCollection(ISheet sheet)
        {
            collection = new Dictionary<int, VideoGameDetails>();
            for (var row_idx = 1; row_idx <= sheet.LastRowNum; row_idx++)
            {
                var cur_row = sheet.GetRow(row_idx);
                var cur_vgd = new VideoGameDetails()
                {
                    Year = (int)cur_row.Cells[(int)VideoGameSalesSheetCols.Year - 1].NumericCellValue,
                    Genre = cur_row.Cells[(int)VideoGameSalesSheetCols.Genre - 1].StringCellValue,
                    Platform = cur_row.Cells[(int)VideoGameSalesSheetCols.Platform - 1].ToString(),
                    Rank = (int)cur_row.Cells[(int)VideoGameSalesSheetCols.Rank - 1].NumericCellValue,
                    Publisher = cur_row.Cells[(int)VideoGameSalesSheetCols.Publisher - 1].StringCellValue
                };
                collection.Add(row_idx, cur_vgd);
            }
        }

        public VideoGameDetails this[int index] => collection[index];
        public IEnumerator<VideoGameDetails> GetEnumerator()
        {
            return collection.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return collection.Values.GetEnumerator();
        }
    }
}
