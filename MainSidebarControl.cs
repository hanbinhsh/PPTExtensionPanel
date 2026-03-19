using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace PPTExtensionPanel
{
    public partial class MainSidebarControl : UserControl
    {
        private string configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PptExtensionPanelConfig.txt");
        private List<string> currentSelectedIds = new List<string>();
        private WebView2 wvContainer;

        private Button btnNavLayout;
        private Button btnNavIcon;

        public MainSidebarControl()
        {
            InitializeComponent();

            // --- 1. 修复底层按钮样式 ---
            btnAddTool.FlatStyle = FlatStyle.Flat;
            btnAddTool.BackColor = Color.FromArgb(64, 64, 64);
            btnAddTool.ForeColor = Color.White;
            btnAddTool.Click += btnAddTool_Click;

            SetupModernOfficeUI();

            flowLayoutPanel1.BackColor = Color.Transparent;
            pnlIconResult.BackColor = Color.Transparent;

            txtSearch.BackColor = Color.White;
            txtSearch.ForeColor = Color.Black;
            txtSearch.BorderStyle = BorderStyle.None;

            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.BackColor = Color.FromArgb(64, 64, 64);
            btnSearch.ForeColor = Color.White;

            btnSearch.Click += btnSearch_Click;

            // 绑定其他事件
            this.Resize += MainSidebarControl_Resize;
            this.Load += MainSidebarControl_Load;
            LoadSavedConfig();
        }

        private async void MainSidebarControl_Load(object sender, EventArgs e)
        {
            try
            {
                wvContainer = new WebView2();
                wvContainer.Visible = false; // 让它隐身工作，用户根本看不到浏览器界面
                this.Controls.Add(wvContainer); // 把它挂载到后台
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PptExtensionPanel", "WebView2Cache");
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await wvContainer.EnsureCoreWebView2Async(env);
                wvContainer.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
                wvContainer.CoreWebView2.NavigationCompleted += WvContainer_CoreWebView2_NavigationCompleted;
            }
            catch (Exception ex)
            {
                MessageBox.Show("后台引擎初始化失败: " + ex.Message);
            }
        }

        private void MainSidebarControl_Resize(object sender, EventArgs e)
        {
            if (flowLayoutPanel1 != null)
            {
                foreach (Control ctrl in flowLayoutPanel1.Controls)
                {
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

        private void RefreshPanel(List<PptTool> tools)
        {
            flowLayoutPanel1.Controls.Clear();
            foreach (var tool in tools)
            {
                Button btn = new Button();
                btn.Text = tool.Name;
                btn.Tag = tool.MsoId;
                btn.Width = flowLayoutPanel1.Width - 10;
                btn.Height = 40;
                btn.Margin = new Padding(3, 3, 3, 5);
                btn.UseVisualStyleBackColor = false;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.BackColor = Color.FromArgb(64, 64, 64);
                btn.ForeColor = Color.White;
                btn.Click += DynamicButton_Click;
                flowLayoutPanel1.Controls.Add(btn);
            }
        }

        private void DynamicButton_Click(object sender, EventArgs e)
        {
            Button clickedBtn = sender as Button;
            if (clickedBtn != null && clickedBtn.Tag != null)
            {
                string msoId = clickedBtn.Tag.ToString();
                try
                {
                    var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        var shapes = selection.ShapeRange;
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
                    Globals.ThisAddIn.Application.ActiveWindow.Activate();
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(msoId);
                }
                catch (Exception)
                {
                    MessageBox.Show("执行失败，请确保你已经在幻灯片上选中了至少两个形状！");
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            // 保护机制：防止后台引擎还没启动完毕用户就点了搜索
            if (wvContainer == null || wvContainer.CoreWebView2 == null)
            {
                MessageBox.Show("浏览器引擎正在努力启动中，请稍等一两秒再试！");
                return;
            }

            string keyword = Uri.EscapeDataString(txtSearch.Text.Trim());
            if (string.IsNullOrEmpty(keyword)) return;

            pnlIconResult.Controls.Clear();

            // MessageBox.Show("正在后台搜索: " + txtSearch.Text); 

            string searchUrl = $"https://www.iconfont.cn/search/index?searchType=icon&q={keyword}";
            wvContainer.CoreWebView2.Navigate(searchUrl);
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 直接读取底层的 JSON 格式字符串
                string json = e.WebMessageAsJson;

                // 此时的 json 会完美变成类似 ["data:image...","data:image..."] 的格式
                List<string> imageBase64List = JsonConvert.DeserializeObject<List<string>>(json);

                if (imageBase64List != null && imageBase64List.Count > 0)
                {
                    foreach (string base64Data in imageBase64List)
                    {
                        CreateIconPreview(base64Data);
                    }
                }
                else
                {
                    // 顺便加个提示，如果抓取结果是空的，我们在界面上能知道
                    MessageBox.Show("网页加载完成，但没有提取到图标，可能是网页结构变了或网络太慢。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("解析图标数据失败: " + ex.Message);
            }
        }

        private void CreateIconPreview(string base64Data)
        {
            string pureBase64 = base64Data.Substring(base64Data.IndexOf(",") + 1);
            byte[] imageBytes = Convert.FromBase64String(pureBase64);

            PictureBox pb = new PictureBox();
            pb.Width = 60;
            pb.Height = 60;
            pb.SizeMode = PictureBoxSizeMode.Zoom;
            pb.Margin = new Padding(5);
            pb.BorderStyle = BorderStyle.FixedSingle;
            pb.Cursor = Cursors.Hand;

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                pb.Image = Image.FromStream(ms);
            }

            pb.Click += (s, ev) => InsertBase64ImageToPpt(base64Data);

            if (pnlIconResult.InvokeRequired)
            {
                pnlIconResult.Invoke(new Action(() => pnlIconResult.Controls.Add(pb)));
            }
            else
            {
                pnlIconResult.Controls.Add(pb);
            }
        }

        private void InsertBase64ImageToPpt(string base64Data)
        {
            try
            {
                Globals.ThisAddIn.Application.ActiveWindow.Activate();
                var activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                string tempPath = Path.Combine(Path.GetTempPath(), $"temp_icon_{Guid.NewGuid()}.png");
                string pureBase64 = base64Data.Substring(base64Data.IndexOf(",") + 1);
                File.WriteAllBytes(tempPath, Convert.FromBase64String(pureBase64));

                activeSlide.Shapes.AddPicture(tempPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    100, 100, -1, -1);

                File.Delete(tempPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("插入图标失败: " + ex.Message);
            }
        }

        private void WvContainer_CoreWebView2_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                InjectIconExtractionScript();
            }
        }

        private async void InjectIconExtractionScript()
        {
            string jsScript = @"
        setTimeout(async function() {
            console.log('开始提取高清图标...');
            const finalImages = [];
            
            // 兼容性查找：Iconfont 的类名有时是 icon-item，有时是 block-icon-list 下的 li
            const iconLis = document.querySelectorAll('.icon-item, .block-icon-list li');

            for (const li of iconLis) {
                try {
                    const svgEl = li.querySelector('svg');
                    if (!svgEl) continue;

                    const useTag = svgEl.querySelector('use');
                    let targetSvg = svgEl;
                    if (useTag) {
                        const symbolId = useTag.getAttribute('xlink:href') || useTag.getAttribute('href');
                        if (symbolId) {
                            const symbolEl = document.querySelector(symbolId);
                            if (symbolEl) {
                                const tempSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
                                tempSvg.innerHTML = symbolEl.innerHTML;
                                const viewBox = symbolEl.getAttribute('viewBox');
                                if (viewBox) tempSvg.setAttribute('viewBox', viewBox);
                                targetSvg = tempSvg; 
                            }
                        }
                    }

                    const clone = targetSvg.cloneNode(true);
                    clone.setAttribute('width', '1000');
                    clone.setAttribute('height', '1000');
                    clone.removeAttribute('style');
                    clone.style.overflow = 'visible';

                    const parentColor = '#333333'; 
                    const allElements = clone.querySelectorAll('*');
                    allElements.forEach(el => {
                        if (el.getAttribute('fill') === 'currentColor') el.setAttribute('fill', parentColor);
                        if (el.getAttribute('stroke') === 'currentColor') el.setAttribute('stroke', parentColor);
                    });

                    const serializer = new XMLSerializer();
                    let svgString = serializer.serializeToString(clone);
                    if (!svgString.includes('xmlns=')) {
                        svgString = svgString.replace('<svg', '<svg xmlns=""http://www.w3.org/2000/svg""');
                    }
                    
                    const base64Svg = 'data:image/svg+xml;base64,' + window.btoa(unescape(encodeURIComponent(svgString)));
                    
                    const base64Png = await new Promise((resolve, reject) => {
                        const img = new Image();
                        img.onload = function() {
                            const canvas = document.createElement('canvas');
                            canvas.width = 1000; canvas.height = 1000;
                            const ctx = canvas.getContext('2d');
                            ctx.drawImage(img, 0, 0);
                            resolve(canvas.toDataURL('image/png'));
                        };
                        img.onerror = () => reject('加载失败');
                        img.src = base64Svg;
                    });

                    finalImages.push(base64Png);
                } catch (e) { console.error('提取单个图标出错', e); }
            }

            // 抓取完毕后发回 C#
            window.chrome.webview.postMessage(finalImages);
        }, 2500);
    ";

            await wvContainer.CoreWebView2.ExecuteScriptAsync(jsScript);
        }

        private void wvContainer_Click(object sender, EventArgs e)
        {

        }

        // 声明需要的变量
        private Panel panelLayout;
        private Panel panelIcon;
        private bool isLayoutSelected = true;

        private void SetupModernOfficeUI()
        {
            // 1. 动态创建两个完全透明的完美容器，它们会 100% 继承 Office 的深/浅色背景！
            panelLayout = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };
            panelIcon = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };

            // 2. 移花接木：把原本在 TabControl 里的所有控件，“偷”出来装进我们的透明容器里
            while (tabControl1.TabPages[0].Controls.Count > 0)
            {
                panelLayout.Controls.Add(tabControl1.TabPages[0].Controls[0]);
            }
            while (tabControl1.TabPages[1].Controls.Count > 0)
            {
                panelIcon.Controls.Add(tabControl1.TabPages[1].Controls[0]);
            }

            // 3. 卸磨杀驴：彻底从界面上删除罪魁祸首 TabControl，并且释放它的内存！
            this.Controls.Remove(tabControl1);
            tabControl1.Dispose();

            // 4. 把我们做好的两个透明面板加到主界面上
            this.Controls.Add(panelLayout);
            this.Controls.Add(panelIcon);

            // 5. 创建顶部导航栏（保持透明）
            Panel navPanel = new Panel { Dock = DockStyle.Top, Height = 35, BackColor = Color.Transparent };

            btnNavLayout = new Button { Text = "排版", Width = 90, Height = 35, Location = new Point(0, 0), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnNavLayout.FlatAppearance.BorderSize = 0;

            btnNavIcon = new Button { Text = "图标", Width = 90, Height = 35, Location = new Point(90, 0), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnNavIcon.FlatAppearance.BorderSize = 0;

            navPanel.Controls.Add(btnNavLayout);
            navPanel.Controls.Add(btnNavIcon);

            this.Controls.Add(navPanel);
            navPanel.SendToBack(); // 避免遮挡内容

            // 6. 绑定切换事件（不再使用 SelectedIndex，而是直接把对应的面板“拉到最前面”）
            btnNavLayout.Click += (s, e) => {
                isLayoutSelected = true;
                panelLayout.BringToFront(); // 让排版面板显示在最前面
                UpdateNavStyle();
            };
            btnNavIcon.Click += (s, e) => {
                isLayoutSelected = false;
                panelIcon.BringToFront();   // 让图标面板显示在最前面
                UpdateNavStyle();
            };

            // 7. 初始化：第一次打开时，把排版面板拉到最前面
            panelLayout.BringToFront();
            UpdateNavStyle();
        }

        // 💡 辅助方法：更新按钮字体（因为没有 TabControl 了，改用 isLayoutSelected 变量判断）
        private void UpdateNavStyle()
        {
            // 当前选中的页面字体加粗，未选中的是常规字体
            btnNavLayout.Font = new Font("Microsoft YaHei", 9F, isLayoutSelected ? FontStyle.Bold : FontStyle.Regular);
            btnNavIcon.Font = new Font("Microsoft YaHei", 9F, !isLayoutSelected ? FontStyle.Bold : FontStyle.Regular);
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnSearch_Click_1(object sender, EventArgs e)
        {

        }
    }
}