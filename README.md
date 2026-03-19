# 🚀 PowerPoint 现代增强侧边栏插件 (PPT Extension Panel)

这是一个基于 C# VSTO 开发的 PowerPoint 增强插件。它深度集成了原生 Office 界面风格，不仅提供可高度自定义的快捷排版工具，还内置了一个基于 WebView2 的无感极速图标搜索引擎，让演示文稿的制作如虎添翼。

## ✨ 核心特性 (Features)

### 📐 1. 快捷排版引擎 (Layout Tools)
* **高度自定义:** 用户可自行添加、移除常用的对齐和分布工具（如左对齐、水平居中、横向分布等）。
* **极简交互:** 抛弃繁琐的 Ribbon 菜单，一键对选中的多个形状执行排版操作。
* **状态记忆:** 用户的工具栏配置会自动保存在本地，下次打开 PPT 完美复原。

### 🎨 2. Iconfont 极速图标库 (Icon Engine)
这是本插件的核心黑科技，基于 `Microsoft.Web.WebView2` 打造的无头（Headless）抓取引擎：
* **⚡ 极速轮询渲染:** 摒弃传统的“死等”延迟，注入 JS 脚本以 0.1 秒级极速轮询 DOM，网络良好时实现毫秒级“秒出”图标。
* **🌈 智能按需上色 (On-Demand Coloring):** * 列表预览时，**完美保留原图的多色细节**。
  * 用户不仅可自定义图标边长（px），还能通过原生调色板选择颜色。
  * 在点击插入的瞬间，引擎会深入 SVG 底层暴力覆写 `fill` 和 `stroke`，将重新渲染后的定制图标瞬间打入幻灯片。
* **📚 深度分页解析:** 自动抓取并解析网页总页数，提供完美的上一页/下一页无缝翻页体验。
* **🛡️ 纯净拦截装甲:** 针对 WinForms 幽灵事件、异步 Promise 黑洞以及网络异常，设计了多层防抖与状态拦截机制，确保插件永远不会卡死或闪退。

### 💎 3. 现代化极简 UI (Modern UI)
* 彻底抛弃了 WinForms 丑陋的自带 TabControl 白底边框。
* 利用纯代码实现的 Z-Order 层级与透明 Panel 交替技术，让侧边栏完美融入 Office 原生的亮色/暗黑主题。
* 输入框即输即存（无需回车），交互手感如丝般顺滑。

---

## 🛠️ 技术栈与依赖 (Tech Stack)

* **框架:** C# / .NET Framework (VSTO)
* **核心组件:** Windows Forms (WinForms)
* **依赖库 (NuGet):**
  * `Microsoft.Web.WebView2` (用于后台静默网页渲染与 JS 交互)
  * `Newtonsoft.Json` (用于 C# 与 JS 之间的高效数据通讯)
* **目标应用:** Microsoft PowerPoint

---

## 📦 安装与运行 (Installation)

使用：
前往Release下载最新版本压缩包，解压并安装exe。
**在加载项选项卡中找到Ice菜单即可显示侧边栏菜单。**

开发：
1. 克隆或下载本仓库代码。
2. 使用 Visual Studio 打开 `.sln` 解决方案文件（请确保已安装 **Office/SharePoint 开发** 工作负载）。
3. 在 NuGet 包管理器中还原所需的 `WebView2` 和 `Newtonsoft.Json` 依赖。
4. 按下 `F5` 编译并运行，Visual Studio 将自动启动 PowerPoint 并挂载本插件。
5. 在 PPT 的功能区点击打开侧边栏，尽情体验！

---

## 📂 核心代码结构简述 (Architecture)

* `MainSidebarControl.cs`: 侧边栏主界面，负责所有的 UI 渲染、主题继承与事件绑定。
* **配置管理:** 采用本地 `.txt` 轻量级存储，记录工具 ID、图标边长、用户自定义颜色，并通过拦截器解决文件读写竞态条件。
* **JS 注入层:** 在 `InjectIconExtractionScript()` 中，向 WebView2 注入包含 DOM 解析、SVG 重绘、Canvas 转换 Base64 的全套异步前端代码。
* **异步通讯层:** 通过 `WebMessageReceived` 精准区分“整页列表抓取”和“单张按需着色”指令，保障主线程的极致流畅。

---

## 📝 TODO / 未来计划

- [ ] 支持更多主流图标网站的引擎切换。
- [ ] 优化无网络环境下的缓存加载策略。
- [ ] 增加针对 PPT 形状的更多一键特效（如统一调色、快捷对齐到幻灯片等）。

---

## 📄 许可协议 (License)

本项目仅供学习与交流使用。图标抓取逻辑依赖于公开网页结构，若目标网站（如 Iconfont）更改 DOM 结构，可能需要同步更新解析脚本。