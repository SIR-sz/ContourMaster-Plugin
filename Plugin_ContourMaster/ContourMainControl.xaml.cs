using System.Windows;
using System.Windows.Controls;

namespace Plugin_ContourMaster
{
    public partial class ContourMainControl : UserControl
    {
        public ContourSettings Settings { get; set; }

        public ContourMainControl()
        {
            InitializeComponent();
            Settings = new ContourSettings();
            this.DataContext = Settings;
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            // 第二步我们要实现的：像素处理核心算法
            MessageBox.Show($"准备处理！当前阈值: {Settings.Threshold}");
        }
    }
}