using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Plugin_ContourMaster.UI;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices; // 必须引用，用于调用 Win32 API
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
        // 引用内核库，用于强制指定 DLL 搜索路径
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        public void Initialize()
        {
            InitializeOpenCvPath();
        }

        public void Terminate()
        {
        }

        private void InitializeOpenCvPath()
        {
            try
            {
                // 1. 获取插件所在的绝对路径
                string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                if (!string.IsNullOrEmpty(assemblyPath))
                {
                    // 2. 检查关键文件是否存在（诊断用）
                    string externDllPath = Path.Combine(assemblyPath, "OpenCvSharpExtern.dll");
                    if (!File.Exists(externDllPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"警告：未在路径下找到 {externDllPath}");
                    }

                    // 3. 核心修复：强制将插件目录加入 DLL 搜索路径
                    SetDllDirectory(assemblyPath);

                    // 4. 双重保险：同时更新进程的环境变量
                    string path = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process);
                    if (!path.Contains(assemblyPath))
                    {
                        Environment.SetEnvironmentVariable("PATH", assemblyPath + ";" + path, EnvironmentVariableTarget.Process);
                    }
                }
            }
            catch (System.Exception ex)
            {
                // 这里使用 System.Exception 避免和 AutoCAD.Runtime.Exception 冲突
                System.Diagnostics.Debug.WriteLine("OpenCV 初始化严重失败: " + ex.Message);
            }
        }
    }

    /// <summary>
    /// 命令交互类
    /// </summary>
    public class MainCommand
    {
        private static PaletteSet _ps = null;

        [CommandMethod("PXQLK_PANEL")]
        [CommandMethod("LK")]
        public void ShowPanel()
        {
            // --- 调试代码开始 ---
            string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string currentPathEnv = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Process);

            // 检查 OpenCV 关键文件是否存在
            string externDll = Path.Combine(assemblyPath, "OpenCvSharpExtern.dll");
            bool exists = File.Exists(externDll);

            // 弹出 AutoCAD 标准诊断对话框
            Application.ShowAlertDialog(
                $"【插件诊断信息】\n\n" +
                $"1. 插件所在目录:\n{assemblyPath}\n\n" +
                $"2. OpenCV核心文件状态: {(exists ? "✅ 已找到" : "❌ 未找到")}\n\n" +
                $"3. PATH环境是否包含该目录: {(currentPathEnv.Contains(assemblyPath) ? "✅ 已包含" : "❌ 未包含")}"
            );
            // --- 调试代码结束 ---

            if (_ps == null)
            {
                _ps = new PaletteSet("像素轮廓工具", new Guid("F3A8E9B2-C12D-4C11-8D9A-2B3C4D5E6F7A"));
                _ps.Size = new Size(300, 600);

                var control = new ContourMainControl();
                ElementHost host = new ElementHost { Child = control, Dock = System.Windows.Forms.DockStyle.Fill };
                _ps.Add("算法设置", host);
            }
            _ps.Visible = true;
        }
    }
}