# 🚀 PowerPoint 现代增强侧边栏插件 (PPT Extension Panel)

这是一个基于 C# VSTO 开发的 PowerPoint 增强插件。它深度集成了原生 Office 界面风格，不仅提供可高度自定义的快捷排版工具，还内置了一个基于 WebView2 的无感极速图标搜索引擎（基于阿里巴巴矢量图标库），让演示文稿的制作如虎添翼。

## ✨ 核心特性 (Features)

### 📐 1. 快捷排版

1. 自定义侧边排版菜单
2. 按需添加按钮，设置自动保存

### 🎨 2. Iconfont 图标库

1. 图标搜索
2. 图标着色
3. 图标大小调整
4. 上一页下一页

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