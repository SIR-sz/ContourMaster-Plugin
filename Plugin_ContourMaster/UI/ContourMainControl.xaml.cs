using Plugin_ContourMaster.Models;
using Plugin_ContourMaster.Services;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace Plugin_ContourMaster.UI
{
    // ✨ 修复 4：添加转换器类解决 XAML 编译错误
    public class InverseBooleanConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b ? !b : false;

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
            => value is bool b ? !b : false;
    }
    public partial class ContourMainControl : UserControl
    {
        // ✨ 定义的是 Settings（大写）
        public ContourSettings Settings { get; set; } = new ContourSettings();

        public ContourMainControl()
        {
            InitializeComponent();
            this.DataContext = Settings;
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            RunEngineAction(engine => engine.ProcessContour());
        }

        // 仅保留图片识别事件
        private void BtnImage_Click(object sender, RoutedEventArgs e)
        {
            RunEngineAction(engine => engine.ProcessReferencedImage());
        }


        private void RunEngineAction(System.Action<ContourEngine> action)
        {
            try
            {
                // ✨ 统一使用 Settings
                var engine = new ContourEngine(Settings);
                action.Invoke(engine);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}", "轮廓提取", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}