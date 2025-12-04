using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LChart_Comparison_Tool;

public class ItemGroup
{
    public string ModuleName { get; set; }      // Combined Column2 + Column3
    public List<string> BlockNumbers { get; set; } = new();
    public string Direction { get; set; }
}

