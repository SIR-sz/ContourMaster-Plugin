using Microsoft.Win32; // 必须引用此命名空间以读写注册表
using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Plugin_ContourMaster.Models
{
    // ✨ OCR 引擎类型枚举
    public enum OcrEngineType
    {
        Tesseract,
        Baidu
    }

    [Obfuscation(Feature = "renaming", Exclude = true)]
    public class ContourSettings : INotifyPropertyChanged
    {
        // 注册表存储路径
        private const string RegPath = @"Software\Plugin_ContourMaster\Settings";

        private double _threshold = 128.0;
        private double _simplifyTolerance = 0.5;
        private int _smoothLevel = 2;
        private string _layerName = "LK_XS";
        private int _precisionLevel = 5;
        private bool _isOcrMode = false;
        private OcrEngineType _selectedOcrEngine = OcrEngineType.Tesseract;
        private string _baiduApiKey = "";
        private string _baiduSecretKey = "";
        private double _winTop = 100;
        private double _left = 100;
        private double _winWidth = 320;
        private double _winHeight = 600;
        /// <summary>
        /// 窗口顶部纵坐标位置
        /// </summary>
        public double WinTop
        {
            get => _winTop;
            set { _winTop = value; OnPropertyChanged(); SaveToRegistry("WinTop", value); }
        }

        /// <summary>
        /// 窗口左侧横坐标位置
        /// </summary>
        public double WinLeft
        {
            get => _left;
            set { _left = value; OnPropertyChanged(); SaveToRegistry("WinLeft", value); }
        }

        /// <summary>
        /// 窗口宽度
        /// </summary>
        public double WinWidth
        {
            get => _winWidth;
            set { _winWidth = value; OnPropertyChanged(); SaveToRegistry("WinWidth", value); }
        }

        /// <summary>
        /// 窗口高度
        /// </summary>
        public double WinHeight
        {
            get => _winHeight;
            set { _winHeight = value; OnPropertyChanged(); SaveToRegistry("WinHeight", value); }
        }
        public ContourSettings()
        {
            // 初始化时自动加载已保存的配置
            LoadFromRegistry();
        }

        #region 自动记忆属性 (Setter 均触发保存)

        public double Threshold
        {
            get => _threshold;
            set { _threshold = value; OnPropertyChanged(); SaveToRegistry("Threshold", value); }
        }

        public double SimplifyTolerance
        {
            get => _simplifyTolerance;
            set { _simplifyTolerance = value; OnPropertyChanged(); SaveToRegistry("SimplifyTolerance", value); }
        }

        public int SmoothLevel
        {
            get => _smoothLevel;
            set { _smoothLevel = value; OnPropertyChanged(); SaveToRegistry("SmoothLevel", value); }
        }

        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); SaveToRegistry("LayerName", value); }
        }

        public int PrecisionLevel
        {
            get => _precisionLevel;
            set { _precisionLevel = value; OnPropertyChanged(); SaveToRegistry("PrecisionLevel", value); }
        }

        public bool IsOcrMode
        {
            get => _isOcrMode;
            set { _isOcrMode = value; OnPropertyChanged(); SaveToRegistry("IsOcrMode", value); }
        }

        public OcrEngineType SelectedOcrEngine
        {
            get => _selectedOcrEngine;
            set { _selectedOcrEngine = value; OnPropertyChanged(); SaveToRegistry("SelectedOcrEngine", value.ToString()); }
        }

        public string BaiduApiKey
        {
            get => _baiduApiKey;
            set { _baiduApiKey = value; OnPropertyChanged(); SaveToRegistry("BaiduApiKey", value); }
        }

        public string BaiduSecretKey
        {
            get => _baiduSecretKey;
            set { _baiduSecretKey = value; OnPropertyChanged(); SaveToRegistry("BaiduSecretKey", value); }
        }

        #endregion

        #region 注册表持久化逻辑

        /// <summary>
        /// 修改后的 LoadFromRegistry 方法：增加窗口位置信息的读取
        /// </summary>
        private void LoadFromRegistry()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegPath))
                {
                    if (key == null) return;

                    // 加载原有的设置
                    _layerName = key.GetValue("LayerName", "LK_XS").ToString();
                    _baiduApiKey = key.GetValue("BaiduApiKey", "").ToString();
                    _baiduSecretKey = key.GetValue("BaiduSecretKey", "").ToString();
                    _threshold = Convert.ToDouble(key.GetValue("Threshold", 128.0));
                    _simplifyTolerance = Convert.ToDouble(key.GetValue("SimplifyTolerance", 0.5));
                    _smoothLevel = Convert.ToInt32(key.GetValue("SmoothLevel", 2));
                    _precisionLevel = Convert.ToInt32(key.GetValue("PrecisionLevel", 5));
                    _isOcrMode = Convert.ToBoolean(key.GetValue("IsOcrMode", false));

                    string ocrTypeStr = key.GetValue("SelectedOcrEngine", "Tesseract").ToString();
                    if (Enum.TryParse(ocrTypeStr, out OcrEngineType ocrType))
                        _selectedOcrEngine = ocrType;

                    // --- 新增：加载窗口位置和大小 ---
                    _winTop = Convert.ToDouble(key.GetValue("WinTop", 100.0));
                    _left = Convert.ToDouble(key.GetValue("WinLeft", 100.0));
                    _winWidth = Convert.ToDouble(key.GetValue("WinWidth", 320.0));
                    _winHeight = Convert.ToDouble(key.GetValue("WinHeight", 600.0));
                }
            }
            catch { /* 忽略读取异常 */ }
        }

        private void SaveToRegistry(string name, object value)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegPath))
                {
                    if (key != null)
                    {
                        // 注册表底层支持多种类型，这里直接存入 object，系统会自动处理常见类型
                        key.SetValue(name, value ?? "");
                    }
                }
            }
            catch { /* 忽略写入异常 */ }
        }

        #endregion

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}