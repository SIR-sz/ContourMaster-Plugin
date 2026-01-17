using Microsoft.Win32; // 必须引用此命名空间以读写注册表
using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Plugin_ContourMaster.Models
{
    // ✨ 新增：OCR 引擎类型枚举
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

        #region 自动记忆属性
        public string BaiduApiKey
        {
            get => _baiduApiKey;
            set
            {
                _baiduApiKey = value;
                OnPropertyChanged();
                SaveToRegistry("BaiduApiKey", value); // 变更即保存到注册表
            }
        }

        public string BaiduSecretKey
        {
            get => _baiduSecretKey;
            set
            {
                _baiduSecretKey = value;
                OnPropertyChanged();
                SaveToRegistry("BaiduSecretKey", value); // 变更即保存到注册表
            }
        }
        #endregion

        public bool IsOcrMode
        {
            get => _isOcrMode;
            set { _isOcrMode = value; OnPropertyChanged(); }
        }

        public OcrEngineType SelectedOcrEngine
        {
            get => _selectedOcrEngine;
            set { _selectedOcrEngine = value; OnPropertyChanged(); }
        }

        public double Threshold
        {
            get => _threshold;
            set { _threshold = value; OnPropertyChanged(); }
        }

        public double SimplifyTolerance
        {
            get => _simplifyTolerance;
            set { _simplifyTolerance = value; OnPropertyChanged(); }
        }

        public int SmoothLevel
        {
            get => _smoothLevel;
            set { _smoothLevel = value; OnPropertyChanged(); }
        }

        public string LayerName
        {
            get => _layerName;
            set { _layerName = value; OnPropertyChanged(); }
        }

        public int PrecisionLevel
        {
            get => _precisionLevel;
            set { _precisionLevel = value; OnPropertyChanged(); }
        }

        #region 注册表持久化逻辑
        private void LoadFromRegistry()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegPath))
                {
                    if (key != null)
                    {
                        _baiduApiKey = key.GetValue("BaiduApiKey", "").ToString();
                        _baiduSecretKey = key.GetValue("BaiduSecretKey", "").ToString();
                    }
                }
            }
            catch { /* 忽略读取异常 */ }
        }

        private void SaveToRegistry(string name, string value)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegPath))
                {
                    if (key != null)
                    {
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