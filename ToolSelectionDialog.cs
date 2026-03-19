using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPTExtensionPanel
{
    public partial class ToolSelectionDialog : Form
    {
        // 定义我们支持的所有 PPT 原生命令字典
        private List<PptTool> allTools = new List<PptTool>
        {
            new PptTool { Name = "左对齐", MsoId = "ObjectsAlignLeft" },
            new PptTool { Name = "水平居中", MsoId = "ObjectsAlignCenterHorizontal" },
            new PptTool { Name = "右对齐", MsoId = "ObjectsAlignRight" },
            new PptTool { Name = "顶端对齐", MsoId = "ObjectsAlignTop" },
            new PptTool { Name = "垂直居中", MsoId = "ObjectsAlignMiddleVertical" },
            new PptTool { Name = "底端对齐", MsoId = "ObjectsAlignBottom" },
            new PptTool { Name = "横向分布", MsoId = "AlignDistributeHorizontally" },
            new PptTool { Name = "纵向分布", MsoId = "AlignDistributeVertically" }
        };

        // 用来存放用户最终勾选的工具
        public List<PptTool> SelectedTools { get; private set; } = new List<PptTool>();

        public ToolSelectionDialog(List<string> alreadySelectedIds)
        {
            InitializeComponent();

            chkListTools.CheckOnClick = true;
            btnSave.Click += btnSave_Click;

            foreach (var tool in allTools)
            {
                bool isChecked = alreadySelectedIds != null && alreadySelectedIds.Contains(tool.MsoId);
                chkListTools.Items.Add(tool, isChecked);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SelectedTools.Clear();
            // 遍历所有打钩的项目，存入 SelectedTools 列表中
            foreach (object itemChecked in chkListTools.CheckedItems)
            {
                SelectedTools.Add((PptTool)itemChecked);
            }

            this.DialogResult = DialogResult.OK; // 告诉主界面，用户点击了保存
            this.Close();
        }
    }
}
