using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTExtensionPanel
{
    public partial class ThisAddIn
    {
        private const string PaneTitle = "Ice菜单";
        private const int PaneWidth = 260;

        private readonly Dictionary<int, CustomTaskPane> panesByWindowHandle = new Dictionary<int, CustomTaskPane>();
        private readonly string paneStatePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PptPanelState.txt");

        private bool defaultPaneVisible = true;

        public Microsoft.Office.Tools.CustomTaskPane CustomTaskPane
        {
            get { return GetActiveWindowPane(); }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.EnableVisualStyles();

            defaultPaneVisible = LoadPaneVisibleState();

            Application.WindowActivate += Application_WindowActivate;

            EnsurePaneForWindow(Application.ActiveWindow);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.WindowActivate -= Application_WindowActivate;

            foreach (CustomTaskPane pane in panesByWindowHandle.Values)
            {
                pane.VisibleChanged -= CustomTaskPane_VisibleChanged;
            }

            panesByWindowHandle.Clear();
        }

        public void ToggleActiveWindowPane()
        {
            CleanupClosedWindowPanes();

            CustomTaskPane pane = EnsurePaneForWindow(Application.ActiveWindow);
            if (pane == null)
            {
                return;
            }

            pane.Visible = !pane.Visible;
        }

        public CustomTaskPane GetActiveWindowPane()
        {
            return EnsurePaneForWindow(Application.ActiveWindow);
        }

        private void Application_WindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            CleanupClosedWindowPanes();
            EnsurePaneForWindow(wn);
        }

        private CustomTaskPane EnsurePaneForWindow(PowerPoint.DocumentWindow window)
        {
            if (window == null)
            {
                return null;
            }

            int windowHandle = window.HWND;
            if (panesByWindowHandle.TryGetValue(windowHandle, out CustomTaskPane existingPane))
            {
                return existingPane;
            }

            MainSidebarControl sidebarControl = new MainSidebarControl();
            CustomTaskPane pane = CustomTaskPanes.Add(sidebarControl, PaneTitle, window);
            pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Width = PaneWidth;
            pane.Visible = defaultPaneVisible;
            pane.VisibleChanged += CustomTaskPane_VisibleChanged;

            panesByWindowHandle[windowHandle] = pane;
            return pane;
        }

        private void CleanupClosedWindowPanes()
        {
            HashSet<int> liveHandles = new HashSet<int>();

            foreach (PowerPoint.DocumentWindow window in Application.Windows)
            {
                liveHandles.Add(window.HWND);
            }

            List<int> closedHandles = new List<int>();
            foreach (int handle in panesByWindowHandle.Keys)
            {
                if (!liveHandles.Contains(handle))
                {
                    closedHandles.Add(handle);
                }
            }

            foreach (int handle in closedHandles)
            {
                CustomTaskPane pane = panesByWindowHandle[handle];
                pane.VisibleChanged -= CustomTaskPane_VisibleChanged;
                CustomTaskPanes.Remove(pane);
                panesByWindowHandle.Remove(handle);
            }
        }

        private void CustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            CustomTaskPane pane = sender as CustomTaskPane;
            if (pane == null)
            {
                return;
            }

            defaultPaneVisible = pane.Visible;

            try
            {
                File.WriteAllText(paneStatePath, defaultPaneVisible.ToString());
            }
            catch
            {
            }
        }

        private bool LoadPaneVisibleState()
        {
            try
            {
                if (!File.Exists(paneStatePath))
                {
                    return true;
                }

                string savedState = File.ReadAllText(paneStatePath).Trim();
                return !string.Equals(savedState, "False", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return true;
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
