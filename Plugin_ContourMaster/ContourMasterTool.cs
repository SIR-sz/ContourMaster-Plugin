using CadAtlasManager.Core;
using Plugin_ContourMaster.UI;
using System;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plugin_ContourMaster
{
    public class ContourMasterTool : ICadTool
    {
        private static bool _isAuthorized = false;
        public string ToolName => "像素轮廓专家";
        public string Description => "像素级轮廓提取工具。";
        public string IconCode => "\uE114";
        public string Category { get; set; } = "图形处理";
        public System.Windows.Media.ImageSource ToolPreview { get; set; }

        public bool VerifyHost(Guid hostGuid)
        {
#if STANDALONE
            _isAuthorized = true; // Debug/Standalone 模式下免授权
#else
            _isAuthorized = (hostGuid == PluginContract.HostGuid); // Release 模式下需校验
#endif
            return _isAuthorized;
        }

        public void Execute()
        {
            if (_isAuthorized) OpenUI();
        }

        private void OpenUI()
        {
            // 逻辑由 MainCommand 统一管理调色板，或在此直接弹窗
            AcApp.DocumentManager.MdiActiveDocument.SendStringToExecute("PXQLK_PANEL ", true, false, false);
        }
    }
}