using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
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
        private readonly string logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PPTExtensionPanel.log");

        private bool defaultPaneVisible = true;
        private Timer deferredStartupTimer;

        public CustomTaskPane CustomTaskPane
        {
            get { return GetActiveWindowPane(); }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                defaultPaneVisible = LoadPaneVisibleState();

                this.Application.WindowActivate += Application_WindowActivate;

                if (defaultPaneVisible)
                {
                    // Defer pane creation until Office finishes its own startup work.
                    deferredStartupTimer = new Timer();
                    deferredStartupTimer.Interval = 800;
                    deferredStartupTimer.Tick += DeferredStartupTimer_Tick;
                    deferredStartupTimer.Start();
                }
            }
            catch (Exception ex)
            {
                WriteLog("ThisAddIn_Startup", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (deferredStartupTimer != null)
                {
                    deferredStartupTimer.Stop();
                    deferredStartupTimer.Tick -= DeferredStartupTimer_Tick;
                    deferredStartupTimer.Dispose();
                    deferredStartupTimer = null;
                }

                this.Application.WindowActivate -= Application_WindowActivate;

                foreach (CustomTaskPane pane in panesByWindowHandle.Values)
                {
                    pane.VisibleChanged -= CustomTaskPane_VisibleChanged;
                }

                panesByWindowHandle.Clear();
            }
            catch (Exception ex)
            {
                WriteLog("ThisAddIn_Shutdown", ex);
            }
        }

        public void ToggleActiveWindowPane()
        {
            try
            {
                CleanupClosedWindowPanes();

                CustomTaskPane pane = EnsurePaneForWindow(this.Application.ActiveWindow);
                if (pane == null)
                {
                    MessageBox.Show("当前没有可用的 PowerPoint 窗口，无法显示 Ice 菜单。");
                    return;
                }

                pane.Visible = !pane.Visible;
            }
            catch (Exception ex)
            {
                WriteLog("ToggleActiveWindowPane", ex);
                MessageBox.Show("Ice 菜单初始化失败，请查看日志：" + logPath);
            }
        }

        public CustomTaskPane GetActiveWindowPane()
        {
            try
            {
                return EnsurePaneForWindow(this.Application.ActiveWindow);
            }
            catch (Exception ex)
            {
                WriteLog("GetActiveWindowPane", ex);
                return null;
            }
        }

        private void DeferredStartupTimer_Tick(object sender, EventArgs e)
        {
            if (deferredStartupTimer == null)
            {
                return;
            }

            deferredStartupTimer.Stop();
            deferredStartupTimer.Tick -= DeferredStartupTimer_Tick;
            deferredStartupTimer.Dispose();
            deferredStartupTimer = null;

            try
            {
                CustomTaskPane pane = EnsurePaneForWindow(this.Application.ActiveWindow);
                if (pane != null)
                {
                    pane.Visible = true;
                }
            }
            catch (Exception ex)
            {
                WriteLog("DeferredStartupTimer_Tick", ex);
            }
        }

        private void Application_WindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            try
            {
                CleanupClosedWindowPanes();

                if (defaultPaneVisible)
                {
                    EnsurePaneForWindow(wn);
                }
            }
            catch (Exception ex)
            {
                WriteLog("Application_WindowActivate", ex);
            }
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
            CustomTaskPane pane = this.CustomTaskPanes.Add(sidebarControl, PaneTitle, window);
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

            foreach (PowerPoint.DocumentWindow window in this.Application.Windows)
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
                this.CustomTaskPanes.Remove(pane);
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
            catch (Exception ex)
            {
                WriteLog("CustomTaskPane_VisibleChanged", ex);
            }
        }

        private bool LoadPaneVisibleState()
        {
            try
            {
                if (!File.Exists(paneStatePath))
                {
                    return false;
                }

                string savedState = File.ReadAllText(paneStatePath).Trim();
                return string.Equals(savedState, "True", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                WriteLog("LoadPaneVisibleState", ex);
                return false;
            }
        }

        private void WriteLog(string stage, Exception ex)
        {
            try
            {
                string logDirectory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(logDirectory) && !Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                string message =
                    "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " +
                    stage + Environment.NewLine +
                    ex + Environment.NewLine + Environment.NewLine;

                File.AppendAllText(logPath, message);
            }
            catch
            {
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
