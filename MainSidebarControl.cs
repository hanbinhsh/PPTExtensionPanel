using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Windows.Forms;

namespace PPTExtensionPanel
{
    public partial class MainSidebarControl : UserControl
    {
        private string configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PptExtensionPanelConfig.txt");
        private List<string> currentSelectedIds = new List<string>();
        private WebView2 wvContainer;

        private Button btnNavLayout;
        private Button btnNavIcon;

        private DateTime lastSearchTime = DateTime.MinValue;

        private int defaultIconSize = 100;
        private string defaultIconColor = "ORIGINAL";

        private int currentPage = 1; // 当前页码
        private int totalPages = 1;  // 总页数（由 JS 动态读取）
        private string currentKeyword = ""; // 记录当前搜索词

        private bool isConfigLoading = false;

        public class IconResult
        {
            public int index { get; set; }
            public string png { get; set; }
        }

        public class SingleColorResult
        {
            public string type { get; set; }
            public string base64 { get; set; }
        }

        public class ListResult
        {
            public string type { get; set; }
            public int totalPages { get; set; }
            public List<IconResult> images { get; set; }
        }

        public MainSidebarControl()
        {
            InitializeComponent();

            if (numIconSize != null)
            {
                // 绑定值改变和键盘抬起事件（实现即输即存）
                numIconSize.ValueChanged += numIconSize_ValueChanged;
                numIconSize.KeyUp += numIconSize_KeyUp;
            }

            if (btnColorPicker != null) btnColorPicker.Click += btnColorPicker_Click;
            if (btnResetColor != null) btnResetColor.Click += btnResetColor_Click;
            if (btnPrevPage != null) btnPrevPage.Click += btnPrevPage_Click;
            if (btnNextPage != null) btnNextPage.Click += btnNextPage_Click;

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

            btnResetColor.FlatStyle = FlatStyle.Flat;
            btnResetColor.BackColor = Color.FromArgb(64, 64, 64);
            btnResetColor.ForeColor = Color.White;

            btnPrevPage.FlatStyle = FlatStyle.Flat;
            btnPrevPage.BackColor = Color.FromArgb(64, 64, 64);
            btnPrevPage.ForeColor = Color.White;

            btnNextPage.FlatStyle = FlatStyle.Flat;
            btnNextPage.BackColor = Color.FromArgb(64, 64, 64);
            btnNextPage.ForeColor = Color.White;

            // 绑定其他事件
            this.Resize += MainSidebarControl_Resize;
            this.Load += MainSidebarControl_Load;
            LoadSavedConfig();
            if (numIconSize != null)
            {
                numIconSize.Value = defaultIconSize;
            }
        }

        private async void MainSidebarControl_Load(object sender, EventArgs e)
        {
            // 💡 核心装甲 1：如果浏览器已经创建过了，直接忽略多余的加载请求
            if (wvContainer != null) return;

            try
            {
                wvContainer = new WebView2();
                wvContainer.Visible = false;
                this.Controls.Add(wvContainer);
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
                isConfigLoading = true;

                try
                {
                    string[] lines = File.ReadAllLines(configPath);

                    // 1. 读排版工具
                    if (lines.Length > 0 && !string.IsNullOrWhiteSpace(lines[0]))
                    {
                        currentSelectedIds = lines[0].Split(',').ToList();
                        List<PptTool> restoredTools = new List<PptTool>();
                        foreach (string id in currentSelectedIds)
                        {
                            restoredTools.Add(new PptTool { Name = GetToolNameById(id), MsoId = id });
                        }
                        RefreshPanel(restoredTools);
                    }

                    // 2. 读图标尺寸（此时即便触发 ValueChanged，也会被开头的 return 拦截）
                    if (lines.Length > 1 && int.TryParse(lines[1], out int savedSize))
                    {
                        defaultIconSize = savedSize;
                        if (numIconSize != null) numIconSize.Value = savedSize;
                    }

                    // 3. 读图标颜色
                    if (lines.Length > 2 && !string.IsNullOrWhiteSpace(lines[2]))
                    {
                        defaultIconColor = lines[2];
                    }

                    // 4. 更新界面颜色
                    UpdateColorUI();
                }
                finally
                {
                    isConfigLoading = false;
                }
            }
        }

        private void UpdateColorUI()
        {
            if (btnColorPicker != null)
            {
                if (defaultIconColor == "ORIGINAL")
                {
                    btnColorPicker.BackColor = Color.Transparent; // 原色模式下透明
                    btnColorPicker.Text = "X"; // 给个小标志代表当前是原色
                    btnColorPicker.ForeColor = Color.Gray;
                }
                else
                {
                    btnColorPicker.BackColor = ColorTranslator.FromHtml(defaultIconColor);
                    btnColorPicker.Text = "";
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

                    SaveConfig();
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
            if ((DateTime.Now - lastSearchTime).TotalMilliseconds < 500) return;
            lastSearchTime = DateTime.Now;

            if (wvContainer == null || wvContainer.CoreWebView2 == null)
            {
                MessageBox.Show("浏览器引擎尚未就绪，请稍等！");
                return;
            }

            currentKeyword = Uri.EscapeDataString(txtSearch.Text.Trim());
            if (string.IsNullOrEmpty(currentKeyword)) return;

            currentPage = 1; // 每次新搜索，强行回到第一页
            LoadIconPage();
        }

        private void btnPrevPage_Click(object sender, EventArgs e)
        {
            if (currentPage > 1)
            {
                currentPage--;
                LoadIconPage();
            }
        }

        private void btnNextPage_Click(object sender, EventArgs e)
        {
            if (currentPage < totalPages && !string.IsNullOrEmpty(currentKeyword))
            {
                currentPage++;
                LoadIconPage();
            }
        }

        private void LoadIconPage()
        {
            // 1. 触发加载 UI 反馈，全面禁用按钮防误触
            btnSearch.Text = "搜索中...";
            btnSearch.Enabled = false;
            if (btnPrevPage != null) btnPrevPage.Enabled = false;
            if (btnNextPage != null) btnNextPage.Enabled = false;
            if (lblPageNum != null) lblPageNum.Text = $"第 {currentPage} 页\n(加载中...)";

            pnlIconResult.Controls.Clear();
            Label lblLoading = new Label
            {
                Text = "正在极速抓取高清图标，请稍候...",
                AutoSize = true,
                Margin = new Padding(10),
                Font = new Font("Microsoft YaHei", 9F),
                ForeColor = SystemColors.GrayText
            };
            pnlIconResult.Controls.Add(lblLoading);

            // 2. 带上 page 参数发起翻页请求
            string searchUrl = $"https://www.iconfont.cn/search/index?searchType=icon&q={currentKeyword}&page={currentPage}";
            wvContainer.CoreWebView2.Navigate(searchUrl);
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.WebMessageAsJson;

                // 拦截：如果是单张上色图片，直接插入并结束
                if (json.Contains("\"type\":\"single\""))
                {
                    var singleResult = JsonConvert.DeserializeObject<SingleColorResult>(json);
                    if (singleResult != null && !string.IsNullOrEmpty(singleResult.base64))
                    {
                        InsertBase64ImageToPpt(singleResult.base64);
                    }
                    return;
                }

                // --- 以下是列表加载完毕后的逻辑 ---
                btnSearch.Text = "搜索";
                btnSearch.Enabled = true;
                pnlIconResult.Controls.Clear();

                if (json.Contains("\"type\":\"list\""))
                {
                    var listResult = JsonConvert.DeserializeObject<ListResult>(json);

                    // 💡 更新总页数，并刷新 UI 指示器
                    totalPages = listResult.totalPages > 0 ? listResult.totalPages : 1;
                    if (lblPageNum != null) lblPageNum.Text = $"第 {currentPage} / {totalPages} 页";

                    // 💡 精准控制上下页按钮状态
                    if (btnPrevPage != null) btnPrevPage.Enabled = (currentPage > 1);
                    if (btnNextPage != null) btnNextPage.Enabled = (currentPage < totalPages);

                    if (listResult.images != null && listResult.images.Count > 0)
                    {
                        foreach (var item in listResult.images)
                        {
                            CreateIconPreview(item);
                        }
                    }
                    else
                    {
                        ShowEmptyMessage();
                    }
                }
                else
                {
                    ShowEmptyMessage();
                }
            }
            catch (Exception ex)
            {
                btnSearch.Text = "搜索";
                btnSearch.Enabled = true;
                if (btnPrevPage != null) btnPrevPage.Enabled = (currentPage > 1);
            }
        }

        private void ShowEmptyMessage()
        {
            Label lblEmpty = new Label { Text = "未找到图标。", AutoSize = true, Margin = new Padding(10), Font = new Font("Microsoft YaHei", 9F) };
            pnlIconResult.Controls.Add(lblEmpty);
            if (btnPrevPage != null) btnPrevPage.Enabled = (currentPage > 1);
            if (btnNextPage != null) btnNextPage.Enabled = false;
            if (lblPageNum != null) lblPageNum.Text = $"第 {currentPage} 页";
        }

        private void CreateIconPreview(IconResult item)
        {
            string pureBase64 = item.png.Substring(item.png.IndexOf(",") + 1);
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

            pb.Click += (s, ev) => TriggerInsertIcon(item.index, item.png);

            if (pnlIconResult.InvokeRequired) pnlIconResult.Invoke(new Action(() => pnlIconResult.Controls.Add(pb)));
            else pnlIconResult.Controls.Add(pb);
        }

        private void TriggerInsertIcon(int index, string originalBase64)
        {
            if (defaultIconColor == "ORIGINAL")
            {
                // 如果是原色，不用麻烦 JS，直接用原本存下来的高清图插入，瞬间完成！
                InsertBase64ImageToPpt(originalBase64);
            }
            else
            {
                // 如果选了颜色，命令 JS 在后台重新渲染一张，画完后 JS 会主动发消息回来
                string script = $"window.processSingleIcon({index}, '{defaultIconColor}')";
                wvContainer.CoreWebView2.ExecuteScriptAsync(script);
            }
        }

        private void InsertBase64ImageToPpt(string base64Data)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Activate();
            var activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            string tempPath = Path.Combine(Path.GetTempPath(), $"temp_icon_{Guid.NewGuid()}.png");
            string pureBase64 = base64Data.Substring(base64Data.IndexOf(",") + 1);
            File.WriteAllBytes(tempPath, Convert.FromBase64String(pureBase64));
            float size = (float)defaultIconSize;
            activeSlide.Shapes.AddPicture(tempPath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                100, 100, size, size);

            File.Delete(tempPath);
        }

        private void WvContainer_CoreWebView2_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                InjectIconExtractionScript();
            }
            else
            {
                // 💡 核心装甲 3：如果没有网络，或者网页打不开
                btnSearch.Text = "搜索";
                btnSearch.Enabled = true;
                pnlIconResult.Controls.Clear();

                Label lblError = new Label
                {
                    Text = "网页加载失败，请检查你的网络连接！",
                    AutoSize = true,
                    Margin = new Padding(10),
                    Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular),
                    ForeColor = Color.Red // 用红色醒目提示
                };
                pnlIconResult.Controls.Add(lblError);
            }
        }

        private async void InjectIconExtractionScript()
        {
            string jsScript = @"
setTimeout(async function() {
    let attempts = 0;
    window.extractedSvgs = []; 
    
    // 专门为 C# 上色待命的画师
    window.processSingleIcon = function(index, color) {
        try {
            const rawSvg = window.extractedSvgs[index];
            if (!rawSvg) return;
            const parser = new DOMParser();
            const doc = parser.parseFromString(rawSvg, 'image/svg+xml');
            const clone = doc.documentElement;

            clone.setAttribute('fill', color);
            clone.style.color = color;
            const allElements = clone.querySelectorAll('*');
            allElements.forEach(el => {
                const tagName = el.tagName.toLowerCase();
                const fill = el.getAttribute('fill');
                const stroke = el.getAttribute('stroke');
                if ((fill && fill !== 'none') || ['path', 'rect', 'circle', 'polygon', 'ellipse'].includes(tagName)) {
                    el.setAttribute('fill', color);
                }
                if (stroke && stroke !== 'none') {
                    el.setAttribute('stroke', color);
                }
            });

            const serializer = new XMLSerializer();
            let newSvgString = serializer.serializeToString(clone);
            if (!newSvgString.includes('xmlns=')) {
                newSvgString = newSvgString.replace('<svg', '<svg xmlns=""http://www.w3.org/2000/svg""');
            }
            const base64Svg = 'data:image/svg+xml;base64,' + window.btoa(unescape(encodeURIComponent(newSvgString)));
            const img = new Image();
            img.onload = function() {
                const canvas = document.createElement('canvas');
                canvas.width = 1000; canvas.height = 1000;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                window.chrome.webview.postMessage({ type: 'single', base64: canvas.toDataURL('image/png') });
            };
            img.src = base64Svg;
        } catch(e) {}
    };

    // 💡 极速轮询与列表提取
    const checkExist = setInterval(async function() {
        const iconLis = document.querySelectorAll('.icon-item, .block-icon-list li');
        
        if (iconLis.length > 0) {
            clearInterval(checkExist); 
            
            // 💡 核心：读取网页上的总页数
            let parsedTotalPages = 1;
            const totalSpan = document.querySelector('.block-pagination .total');
            if (totalSpan) {
                const match = totalSpan.innerText.match(/\d+/);
                if (match) parsedTotalPages = parseInt(match[0], 10);
            }

            const finalImages = [];
            for (let i = 0; i < iconLis.length; i++) {
                try {
                    const li = iconLis[i];
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

                    const serializer = new XMLSerializer();
                    let svgString = serializer.serializeToString(clone);
                    if (!svgString.includes('xmlns=')) {
                        svgString = svgString.replace('<svg', '<svg xmlns=""http://www.w3.org/2000/svg""');
                    }
                    
                    window.extractedSvgs.push(svgString);
                    const currentIndex = window.extractedSvgs.length - 1;

                    const base64Svg = 'data:image/svg+xml;base64,' + window.btoa(unescape(encodeURIComponent(svgString)));
                    const base64Png = await new Promise((resolve) => {
                        const img = new Image();
                        img.onload = function() {
                            const canvas = document.createElement('canvas');
                            canvas.width = 1000; canvas.height = 1000;
                            const ctx = canvas.getContext('2d');
                            ctx.drawImage(img, 0, 0);
                            resolve(canvas.toDataURL('image/png'));
                        };
                        img.onerror = () => resolve(null);
                        img.src = base64Svg;
                    });

                    if(base64Png) {
                        finalImages.push({ index: currentIndex, png: base64Png });
                    }
                } catch (e) { }
            }

            // 💡 核心：把总页数和图片打包成一个 List 对象发给 C#
            window.chrome.webview.postMessage({
                type: 'list',
                totalPages: parsedTotalPages,
                images: finalImages
            });
        } 
        else {
            attempts++;
            if (attempts > 100) { 
                clearInterval(checkExist);
                window.chrome.webview.postMessage({ type: 'list', totalPages: 1, images: [] }); 
            }
        }
    }, 100); 
}, 500); 
";
            await wvContainer.CoreWebView2.ExecuteScriptAsync(jsScript);
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

            if (pnlIconResult != null) pnlIconResult.BringToFront();
            if (tableLayoutPanel1 != null) tableLayoutPanel1.SendToBack();

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

            btnNavLayout.UseVisualStyleBackColor = false;
            btnNavIcon.UseVisualStyleBackColor = false;

            btnNavLayout.FlatStyle = FlatStyle.Flat;
            btnNavLayout.FlatAppearance.BorderSize = 0;
            btnNavLayout.BackColor = Color.Transparent;
            btnNavLayout.FlatAppearance.MouseOverBackColor = Color.FromArgb(80, 80, 80);
            btnNavLayout.FlatAppearance.MouseDownBackColor = Color.FromArgb(60, 60, 60);

            btnNavIcon.FlatStyle = FlatStyle.Flat;
            btnNavIcon.FlatAppearance.BorderSize = 0;
            btnNavIcon.BackColor = Color.Transparent;
            btnNavIcon.FlatAppearance.MouseOverBackColor = Color.FromArgb(80, 80, 80);
            btnNavIcon.FlatAppearance.MouseDownBackColor = Color.FromArgb(60, 60, 60);

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

        private void SaveConfig()
        {
            if (isConfigLoading) return;

            try
            {
                File.WriteAllLines(configPath, new string[] {
            string.Join(",", currentSelectedIds),
            defaultIconSize.ToString(),
            defaultIconColor // 把颜色存入本地文件
        });
            }
            catch { }
        }

        private void numIconSize_ValueChanged(object sender, EventArgs e)
        {
            defaultIconSize = (int)numIconSize.Value;
            SaveConfig();
        }

        private void numIconSize_KeyUp(object sender, KeyEventArgs e)
        {
            // 只要键盘抬起，哪怕没按回车，也强行读取文本并保存！
            if (int.TryParse(numIconSize.Text, out int newSize))
            {
                if (newSize >= numIconSize.Minimum && newSize <= numIconSize.Maximum)
                {
                    defaultIconSize = newSize;
                    SaveConfig();
                }
            }
        }

        private void btnColorPicker_Click(object sender, EventArgs e)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                cd.FullOpen = true;
                if (defaultIconColor != "ORIGINAL") cd.Color = ColorTranslator.FromHtml(defaultIconColor);

                if (cd.ShowDialog() == DialogResult.OK)
                {
                    defaultIconColor = ColorTranslator.ToHtml(cd.Color); // 转成 #FFFFFF 格式
                    UpdateColorUI();
                    SaveConfig();
                }
            }
        }

        private void btnResetColor_Click(object sender, EventArgs e)
        {
            defaultIconColor = "ORIGINAL";
            UpdateColorUI();
            SaveConfig();
        }
    }
}