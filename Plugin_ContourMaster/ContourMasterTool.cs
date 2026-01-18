using CadAtlasManager.Core;
using Plugin_ContourMaster.UI;
using System;
using System.Windows.Media;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plugin_ContourMaster
{
    public class ContourMasterTool : ICadTool
    {
        private static bool _isAuthorized = false;
        public string ToolName => "轮廓图像文字识别";// 按钮显示的名称
        public string Description => "轮廓提取、图片转ACD、图片文字转多行文本。";
        public string IconCode => "\uE114";// 布局图标代码
        public string Category { get; set; } // 由主程序加载时根据目录自动分配                                         

        public ImageSource ToolPreview { get; set; }  // 当目录下有同名图片时，CadAtlasManager 主程序会自动填充它。

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