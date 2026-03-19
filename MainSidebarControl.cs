using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Linq;

namespace PPTExtensionPanel
{
    public partial class MainSidebarControl : UserControl
    {
        private string configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PptExtensionPanelConfig.txt");
        private List<string> currentSelectedIds = new List<string>();
        public MainSidebarControl()
        {
            InitializeComponent();
            btnAddTool.FlatStyle = FlatStyle.Flat;
            btnAddTool.BackColor = Color.FromArgb(64, 64, 64);
            btnAddTool.ForeColor = Color.White;
            btnAddTool.Click += btnAddTool_Click;
            this.Resize += MainSidebarControl_Resize;
            LoadSavedConfig();
        }
        private void MainSidebarControl_Resize(object sender, EventArgs e)
        {
            // 确保 flowLayoutPanel1 不为空
            if (flowLayoutPanel1 != null)
            {
                foreach (Control ctrl in flowLayoutPanel1.Controls)
                {
                    // 只调整按钮的宽度
                    if (ctrl is Button)
                    {
                        ctrl.Width = flowLayoutPanel1.Width - 10;
                    }
                }
            }
        }

        private void LoadSavedConfig()
        {
            if (File.Exists(configPath))
            {
                string savedData = File.ReadAllText(configPath);
                if (!string.IsNullOrWhiteSpace(savedData))
                {
                    currentSelectedIds = savedData.Split(',').ToList();

                    // 根据保存的 ID 重建按钮列表
                    List<PptTool> restoredTools = new List<PptTool>();
                    foreach (string id in currentSelectedIds)
                    {
                        restoredTools.Add(new PptTool { Name = GetToolNameById(id), MsoId = id });
                    }
                    RefreshPanel(restoredTools);
                }
            }
        }

        private string GetToolNameById(string id)
        {
            switch (id)
            {
                case "ObjectsAlignLeft": return "左对齐";
                case "ObjectsAlignCenterHorizontal": return "水平居中";
                case "ObjectsAlignRight": return "右对齐";
                case "ObjectsAlignTop": return "顶端对齐";
                case "ObjectsAlignMiddleVertical": return "垂直居中";
                case "ObjectsAlignBottom": return "底端对齐";
                case "AlignDistributeHorizontally": return "横向分布";
                case "AlignDistributeVertically": return "纵向分布";
                default: return "快捷工具";
            }
        }

        // 点击底部的“添加”按钮，弹出我们刚才写的设置窗口
        private void btnAddTool_Click(object sender, EventArgs e)
        {
            using (ToolSelectionDialog dialog = new ToolSelectionDialog(currentSelectedIds))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    currentSelectedIds.Clear();
                    foreach (var tool in dialog.SelectedTools)
                    {
                        currentSelectedIds.Add(tool.MsoId);
                    }

                    File.WriteAllText(configPath, string.Join(",", currentSelectedIds));

                    RefreshPanel(dialog.SelectedTools);
                }
            }
        }

        // 动态生成按钮的魔法在这里
        private void RefreshPanel(List<PptTool> tools)
        {
            flowLayoutPanel1.Controls.Clear(); // 先清空之前的按钮

            foreach (var tool in tools)
            {
                Button btn = new Button();
                btn.Text = tool.Name;
                btn.Tag = tool.MsoId; // 把 PPT 的原生命令 ID 藏在 Tag 属性里

                // 解决你“横着容易点错”的痛点：把按钮做大，铺满宽度
                btn.Width = flowLayoutPanel1.Width - 10;
                btn.Height = 40;
                btn.Margin = new Padding(3, 3, 3, 5); // 留点间距更好看

                // 绑定点击事件
                btn.Click += DynamicButton_Click;

                // 把生成的按钮塞进面板里
                flowLayoutPanel1.Controls.Add(btn);
            }
        }

        // 动态按钮被点击时，触发 PPT 原生对齐命令
        private void DynamicButton_Click(object sender, EventArgs e)
        {
            Button clickedBtn = sender as Button;
            if (clickedBtn != null && clickedBtn.Tag != null)
            {
                string msoId = clickedBtn.Tag.ToString();
                try
                {
                    // 1. 获取当前用户选中的对象
                    var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                    // 2. 检查用户是否真的选中了“形状”
                    if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        var shapes = selection.ShapeRange;

                        // 对齐操作至少需要选中 2 个形状
                        if (shapes.Count >= 2)
                        {
                            var msoFalse = Microsoft.Office.Core.MsoTriState.msoFalse;

                            switch (msoId)
                            {
                                case "ObjectsAlignLeft": shapes.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignLefts, msoFalse); return;
                                case "ObjectsAlignCenterHorizontal": shapes.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, msoFalse); return;
                                case "ObjectsAlignRight": shapes.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignRights, msoFalse); return;
                                case "ObjectsAlignTop": shapes.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignTops, msoFalse); return;
                                case "ObjectsAlignMiddleVertical": shapes.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles, msoFalse); return;
                                case "ObjectsAlignBottom": shapes.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignBottoms, msoFalse); return;
                                case "AlignDistributeHorizontally": shapes.Distribute(Microsoft.Office.Core.MsoDistributeCmd.msoDistributeHorizontally, msoFalse); return;
                                case "AlignDistributeVertically": shapes.Distribute(Microsoft.Office.Core.MsoDistributeCmd.msoDistributeVertically, msoFalse); return;
                            }
                        }
                    }

                    // 3. 兜底方案：如果你以后在字典里加了其他原生功能（比如组合、置于顶层），依然走原生调用
                    Globals.ThisAddIn.Application.ActiveWindow.Activate();
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(msoId);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("执行失败，请确保你已经在幻灯片上选中了至少两个形状！");
                }
            }
        }

        private void MainSidebarControl_Load(object sender, EventArgs e)
        {
            
        }

        private void btnAddTool_Click_1(object sender, EventArgs e)
        {

        }
    }
}
