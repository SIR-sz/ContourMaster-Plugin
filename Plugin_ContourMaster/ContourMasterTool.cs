using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Core;
using System;
using System.Windows.Media;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plugin_ContourMaster
{
    public class ContourMasterTool : ICadTool
    {
        // 授权状态标记
        private static bool _isAuthorized = false;

        // --- 插件元数据 ---
        public string ToolName => "像素轮廓专家";
        public string Description => "基于像素分析提取位图轮廓并转化为CAD矢量线。";
        public string IconCode => "\uE114";
        public string Category { get; set; } = "图形处理"; //

        // 预览图属性（主程序会自动加载同名PNG填充）
        public ImageSource ToolPreview { get; set; } //

        // 1. 握手校验：主程序加载时会调用此方法
        public bool VerifyHost(Guid hostGuid)
        {
#if STANDALONE
            // 【独立版逻辑】：Debug/Standalone 模式下，直接赋予权限
            _isAuthorized = true;
#else
            // 【授权版逻辑】：Release 模式下，必须匹配主程序 GUID
            _isAuthorized = (hostGuid == PluginContract.HostGuid);
#endif
            return _isAuthorized;
        }

        // 2. 执行入口：点击主程序面板按钮时触发
        public void Execute()
        {
            // 严谨起见，执行前检查授权
            if (!_isAuthorized) return;

            OpenMainUI();
        }

        // 3. 快捷键入口：用户在CAD命令行输入 PXQLK 触发
        [CommandMethod("PXQLK")]
        public void ShortcutCommand()
        {
#if !STANDALONE
            // 【授权版限制】：如果是 Release 编译，防止绕过主程序直接运行
            if (!_isAuthorized)
            {
                var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\n[受限] 此插件为 CadAtlasManager 专属工具，请通过主面板启动！");
                return;
            }
#endif
            // 如果是 STANDALONE 模式，这里会直接跳过检查进入 UI
            OpenMainUI();
        }

        // 4. UI 启动逻辑
        private void OpenMainUI()
        {
            try
            {
                // TODO: 实例化您的 WPF 窗口或面板
                // var win = new ContourWindow();
                // AcApp.ShowModelessWindow(win);

                var ed = AcApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("\n[像素轮廓专家] 功能已激活。");
            }
            catch (System.Exception ex)
            {
                if (AcApp.DocumentManager.MdiActiveDocument != null)
                {
                    AcApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n[异常] {ex.Message}");
                }
            }
        }
    }
}