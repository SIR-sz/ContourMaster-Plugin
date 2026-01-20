using Autodesk.AutoCAD.ApplicationServices; // 新增：用于 Application.MainWindow
using Plugin_ContourMaster.Models;
using Plugin_ContourMaster.Services;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop; // 新增：用于 WindowInteropHelper

namespace Plugin_ContourMaster.UI
{


    // ✨ 确保是 public，否则 XAML 找不到
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b ? !b : false;

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b ? !b : false;
    }

    // ✨ 新增：确保是 public
    public class InverseBooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b && b ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => throw new NotImplementedException();
    }

    public partial class ContourMainControl : UserControl
    {
        public ContourSettings Settings { get; set; } = new ContourSettings();
        private static ContourFloatingWindow _floatingWin = null;
        public ContourMainControl()
        {
            InitializeComponent();
            this.DataContext = Settings;

            // 同步初始值到密码框
            if (PwdApiKey != null) PwdApiKey.Password = Settings.BaiduApiKey;
            if (PwdSecretKey != null) PwdSecretKey.Password = Settings.BaiduSecretKey;
        }

        private void PwdApiKey_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (sender is PasswordBox pBox && Settings.BaiduApiKey != pBox.Password)
                Settings.BaiduApiKey = pBox.Password;
        }

        private void PwdSecretKey_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (sender is PasswordBox pBox && Settings.BaiduSecretKey != pBox.Password)
                Settings.BaiduSecretKey = pBox.Password;
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e) => RunEngineAction(engine => engine.ProcessContour());

        private async void BtnImage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var engine = new ContourEngine(Settings);
                await engine.ProcessReferencedImage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}");
            }
        }

        private void RunEngineAction(Action<ContourEngine> action)
        {
            try { action.Invoke(new ContourEngine(Settings)); }
            catch (Exception ex) { MessageBox.Show($"异常: {ex.Message}"); }
        }
        // 在类中增加此方法
        private void BtnAbout_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            AboutWindow aboutWin = new AboutWindow();
            // 设置为主窗口的子窗口，使其总是在最前
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(aboutWin);
        }
        // 在 UI/ContourMainControl.xaml.cs 中添加/替换以下静态字段和方法

        private static Autodesk.AutoCAD.Windows.PaletteSet _ps = null;

        /// <summary>
        /// [核心方法]: 启动悬浮工具窗口。
        /// 逻辑：检查实例是否存在，若不存在则新建并绑定 AutoCAD 为所有者，实现非模态显示。
        /// </summary>
        public static void ShowTool()
        {
            try
            {
                // 1. 如果窗口已存在且已加载，则直接激活并带到前台
                if (_floatingWin != null && _floatingWin.IsLoaded)
                {
                    _floatingWin.Activate();
                    return;
                }

                // 2. 初始化设置并创建窗口实例
                // 注意：ContourSettings 构造函数会自动从注册表加载保存的位置
                var settings = new ContourSettings();
                _floatingWin = new ContourFloatingWindow(settings);

                // 3. 绑定所有者关系
                // 通过 WindowInteropHelper 获取 AutoCAD 主句柄，确保窗口随 CAD 最小化/还原
                IntPtr acadHwnd = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Handle;
                WindowInteropHelper helper = new WindowInteropHelper(_floatingWin);
                helper.Owner = acadHwnd;

                // 4. 调用 AutoCAD API 显示非模态窗口
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessWindow(_floatingWin);
            }
            catch (System.Exception ex)
            {
                // 捕获启动过程中的任何异常
                System.Windows.MessageBox.Show($"[ContourMaster] 启动悬浮窗失败: {ex.Message}");
            }
        }
        // --- UI/ContourMainControl.xaml.cs ---

        /// <summary>
        /// 处理标题栏鼠标左键按下事件，实现窗口拖动
        /// </summary>
        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
            {
                // 找到当前的承载窗口并执行拖动
                System.Windows.Window.GetWindow(this)?.DragMove();
            }
        }


        /// <summary>
        /// 处理自定义关闭按钮点击事件
        /// </summary>
        private void BtnClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 关闭承载该用户控件的窗口
            System.Windows.Window.GetWindow(this)?.Close();
        }
    }
}