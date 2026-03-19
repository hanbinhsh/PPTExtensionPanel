using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTExtensionPanel
{
    public class PptTool
    {
        public string Name { get; set; }  // 显示在按钮上的名字，比如"左对齐"
        public string MsoId { get; set; } // PPT 原生的命令 ID，比如"ObjectsAlignLeft"

        // 重写 ToString 是为了在后面的 CheckedListBox 中直接显示名字
        public override string ToString()
        {
            return Name;
        }
    }
}
