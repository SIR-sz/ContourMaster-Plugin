using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Plugin_ContourMaster.UI;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms.Integration;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
using WpfWindow = System.Windows.Window;

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
        private static WpfWindow _toolWindow = null;

        [CommandMethod("PXQLK_PANEL")]
        [CommandMethod("LK")]
        public void ShowPanel()
        {
            // 如果窗口已存在且未被销毁，则激活它
            if (_toolWindow != null && _toolWindow.IsVisible)
            {
                _toolWindow.Activate();
                return;
            }

            _toolWindow = new WpfWindow
            {
                Title = "轮廓图像文字识别工具",
                Content = new ContourMainControl(),
                Width = 320,
                Height = 750,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
                ResizeMode = System.Windows.ResizeMode.CanMinimize,
                ShowInTaskbar = false
            };

            // 使用 AutoCAD 方式显示非模态窗口
            AcApp.ShowModelessWindow(_toolWindow);
        }
    }
}