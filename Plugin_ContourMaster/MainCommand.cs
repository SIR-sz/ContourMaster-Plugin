using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Plugin_ContourMaster.UI; // 引用 UI 命名空间
using System;
using System.Drawing;
using System.Windows.Forms.Integration;

[assembly: CommandClass(typeof(Plugin_ContourMaster.MainCommand))]

namespace Plugin_ContourMaster
{
    public class MainCommand
    {
        private static PaletteSet _ps = null;

        [CommandMethod("PXQLK_PANEL")]
        [CommandMethod("XSFLK")]
        public void ShowPanel()
        {
            if (_ps == null)
            {
                _ps = new PaletteSet("像素轮廓工具", new Guid("F3A8E9B2-C12D-4C11-8D9A-2B3C4D5E6F7A"));
                _ps.Size = new Size(300, 600);

                // 实例化 UI
                var control = new ContourMainControl();
                ElementHost host = new ElementHost { Child = control, Dock = System.Windows.Forms.DockStyle.Fill };
                _ps.Add("算法设置", host);
            }
            _ps.Visible = true;
        }
    }
}