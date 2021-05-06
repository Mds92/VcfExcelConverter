using System.Collections.Generic;

namespace VcfConverter.Classes
{
    public class ExcelFileModel
    {
        public List<ExcelCellModel> Cells { get; set; }
        public List<List<dynamic>> Rows { get; set; }
    }

    public class ExcelCellModel
    {
        public string Title { get; set; }
        public int Width { get; set; }

        /// <summary>
        /// تعداد ستون هایی که باید با هم مرج شوند
        /// </summary>
        public int MergeColumnCount { get; set; }

        public string NumberFormat { get; set; }
    }
}
