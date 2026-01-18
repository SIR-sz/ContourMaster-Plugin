using Plugin_ContourMaster.Models;
using Plugin_ContourMaster.Services;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

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
    }
}