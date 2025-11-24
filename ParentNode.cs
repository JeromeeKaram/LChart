using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LChart_Comparison_Tool
{
    public class ParentNode
    {
        public Microsoft.Office.Interop.Excel.Range ParentCell;
        public Microsoft.Office.Interop.Excel.Range MergedArea;
    }

    public class UpResult
    {
        public Microsoft.Office.Interop.Excel.Range ParentMergedCell { get; set; } = null;
        public List<Microsoft.Office.Interop.Excel.Range> LeftCells { get; set; } = new();
        public List<Microsoft.Office.Interop.Excel.Range> RightCells { get; set; } = new();
    }

}
