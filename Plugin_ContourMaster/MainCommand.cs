using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms.Integration;

// 注册命令类
[assembly: CommandClass(typeof(Plugin_ContourMaster.MainCommand))]
[assembly: ExtensionApplication(typeof(Plugin_ContourMaster.MainCommand))]

namespace Plugin_ContourMaster
{
    public class MainCommand : IExtensionApplication
    {
        static PaletteSet _ps = null;

        public void Initialize()
        {
#if STANDALONE
            // 独立版在加载时可以给用户一点反馈
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\n>>> 像素轮廓专家 (独立定制版) 加载成功！");
#endif
        }

        public void Terminate() { }

        // 在 MainCommand.cs 的 ShowContourPanel 方法中更新：
        [CommandMethod("PXQLK_PANEL")]
        public void ShowContourPanel()
        {
            if (_ps == null)
            {
                Guid myPaletteId = new Guid("F3A8E9B2-C12D-4C11-8D9A-2B3C4D5E6F7A");
                _ps = new PaletteSet("像素轮廓工具箱", myPaletteId);
                _ps.MinimumSize = new System.Drawing.Size(300, 500);
                _ps.Style = PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.ShowCloseButton;

                // 加载刚才创建的 WPF 控件
                var control = new ContourMainControl();
                ElementHost host = new ElementHost { Child = control, Dock = System.Windows.Forms.DockStyle.Fill };
                _ps.Add("算法参数", host);
            }
            _ps.Visible = true;
        }
    }
}