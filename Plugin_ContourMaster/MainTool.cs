using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Core;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

// 注册插件入口和命令类
[assembly: ExtensionApplication(typeof(Plugin_ContourMaster.MainTool))]
[assembly: CommandClass(typeof(Plugin_ContourMaster.MainTool))]

namespace Plugin_ContourMaster
{
    /// <summary>
    /// 插件核心入口类：实现 ICadTool 接口以集成到主程序，同时提供初始化和命令行入口。
    /// </summary>
    public class MainTool : ICadTool, IExtensionApplication
    {
        // 授权标记：用于判断是否通过主程序合法启动
        private static bool _isAuthorized = false;

        #region --- IExtensionApplication 接口实现 (初始化逻辑) ---

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        /// <summary>
        /// 插件加载时的初始化逻辑
        /// </summary>
        public void Initialize()
        {
            InitializeOpenCvPath();
        }

        public void Terminate() { }

        /// <summary>
        /// 核心初始化：确保 OpenCV 库能够正确加载
        /// </summary>
        private void InitializeOpenCvPath()
        {
            try
            {
                string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(assemblyPath))
                {
                    SetDllDirectory(assemblyPath);
                    string path = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process);
                    if (path != null && !path.Contains(assemblyPath))
                    {
                        Environment.SetEnvironmentVariable("PATH", assemblyPath + ";" + path, EnvironmentVariableTarget.Process);
                    }
                }
            }
            catch { }
        }

        #endregion

        #region --- ICadTool 接口实现 (供主程序加载调用) ---

        public string ToolName => "像素轮廓专家";

        public string IconCode => "\uE114";

        public string Description => "从 CAD 元素中提取像素级的闭合轮廓线。";

        public string Category { get; set; } = "图形处理";

        public ImageSource ToolPreview { get; set; }

        /// <summary>
        /// 宿主身份验证：防止插件被非法平台加载
        /// </summary>
        public bool VerifyHost(Guid hostGuid)
        {
            // 固定对比 CadAtlasManager 主程序的 GUID
            return hostGuid == new Guid("A7F3E2B1-4D5E-4B8C-9F0A-1C2B3D4E5F6B");
        }

        /// <summary>
        /// 执行入口：当用户在主程序面板点击图标时触发
        /// </summary>
        public void Execute()
        {
            _isAuthorized = true;
            UI.ContourMainControl.ShowTool();
        }

        #endregion

        #region --- AutoCAD 命令行入口 (支持独立运行调试) ---

        /// <summary>
        /// 命令行入口：支持 PXQLK_PANEL 或简写 LK 命令
        /// </summary>
        [CommandMethod("PXQLK_PANEL")]
        [CommandMethod("PKPL")]
        public void MainCommandEntry()
        {
#if STANDALONE || DEBUG
            // 调试或独立版模式下直接运行
            ShowUIInternal();
#else
            // Release 授权模式下检查授权标记
            if (_isAuthorized)
            {
                ShowUIInternal();
            }
            else
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\n[错误] 该插件为 智汇CAD全流程管理系统 授权版。");
                    doc.Editor.WriteMessage("\n[提示] 请先启动主程序并在主面板中点击插件图标运行。");
                }
            }
#endif
        }

        /// <summary>
        /// 统一的 UI 启动逻辑调用
        /// </summary>
        private void ShowUIInternal()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Editor.WriteMessage($"\n[{ToolName}] 正在启动...");
            UI.ContourMainControl.ShowTool();
        }

        #endregion
    }
}