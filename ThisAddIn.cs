using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PPTExtensionPanel
{
    public partial class ThisAddIn
    {

        private MainSidebarControl sidebarControl;
        public Microsoft.Office.Tools.CustomTaskPane CustomTaskPane { get; private set; }

        private string paneStatePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PptPanelState.txt");

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            sidebarControl = new MainSidebarControl();
            CustomTaskPane = this.CustomTaskPanes.Add(sidebarControl, "Ice菜单");
            CustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            CustomTaskPane.Width = 200;
            bool isVisible = true;
            if (System.IO.File.Exists(paneStatePath))
            {
                string savedState = System.IO.File.ReadAllText(paneStatePath).Trim();
                if (savedState == "False")
                {
                    isVisible = false;
                }
            }
            CustomTaskPane.Visible = isVisible;
            CustomTaskPane.VisibleChanged += CustomTaskPane_VisibleChanged;
        }

        private void CustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            try
            {
                // 侧边栏状态一旦改变，立刻把 True 或 False 写进本地文件
                System.IO.File.WriteAllText(paneStatePath, CustomTaskPane.Visible.ToString());
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
