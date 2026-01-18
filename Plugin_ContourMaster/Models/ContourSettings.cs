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

        private void LoadFromRegistry()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegPath))
                {
                    if (key == null) return;

                    // 加载字符串类型
                    _layerName = key.GetValue("LayerName", "LK_XS").ToString();
                    _baiduApiKey = key.GetValue("BaiduApiKey", "").ToString();
                    _baiduSecretKey = key.GetValue("BaiduSecretKey", "").ToString();

                    // 加载数值类型 (处理转换)
                    _threshold = Convert.ToDouble(key.GetValue("Threshold", 128.0));
                    _simplifyTolerance = Convert.ToDouble(key.GetValue("SimplifyTolerance", 0.5));
                    _smoothLevel = Convert.ToInt32(key.GetValue("SmoothLevel", 2));
                    _precisionLevel = Convert.ToInt32(key.GetValue("PrecisionLevel", 5));
                    _isOcrMode = Convert.ToBoolean(key.GetValue("IsOcrMode", false));

                    // 加载枚举类型
                    string ocrTypeStr = key.GetValue("SelectedOcrEngine", "Tesseract").ToString();
                    if (Enum.TryParse(ocrTypeStr, out OcrEngineType ocrType))
                        _selectedOcrEngine = ocrType;
                }
            }
            catch { /* 忽略读取异常，使用默认值 */ }
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