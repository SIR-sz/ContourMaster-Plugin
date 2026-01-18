using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Plugin_ContourMaster.UI;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms.Integration;

// 标记插件入口类
[assembly: ExtensionApplication(typeof(Plugin_ContourMaster.PluginEntry))]
// 标记命令类
[assembly: CommandClass(typeof(Plugin_ContourMaster.MainCommand))]

namespace Plugin_ContourMaster
{
    /// <summary>
    /// 插件初始化入口
    /// </summary>
    public class PluginEntry : IExtensionApplication
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        public void Initialize()
        {
            InitializeOpenCvPath();
        }

        public void Terminate()
        {
        }

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
                    // 强制将插件目录加入 DLL 搜索路径
                    SetDllDirectory(assemblyPath);

                    // 更新进程环境变量 PATH
                    string path = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process);
                    if (path != null && !path.Contains(assemblyPath))
                    {
                        Environment.SetEnvironmentVariable("PATH", assemblyPath + ";" + path, EnvironmentVariableTarget.Process);
                    }
                }
            }
            catch
            {
                // 静默处理初始化异常
            }
        }
    }

    /// <summary>
    /// 命令交互类
    /// </summary>
    public class MainCommand
    {
        private static PaletteSet _ps = null;

        /// <summary>
        /// 显示插件面板
        /// </summary>
        [CommandMethod("PXQLK_PANEL")]
        [CommandMethod("LK")]
        public void ShowPanel()
        {
            if (_ps == null)
            {
                _ps = new PaletteSet("图像轮廓文字识别工具", new Guid("F3A8E9B2-C12D-4C11-8D9A-2B3C4D5E6F7A"));
                _ps.Size = new Size(300, 600);

                var control = new ContourMainControl();
                ElementHost host = new ElementHost { Child = control, Dock = System.Windows.Forms.DockStyle.Fill };
                _ps.Add("算法设置", host);
            }
            _ps.Visible = true;
        }
    }
}