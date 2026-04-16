using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Timers;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AutoSaveAddIn
{
    public partial class ThisAddIn
    {
        private Timer _saveTimer;

        private bool _isCurrentSaveEnabled = false;
        public bool IsCurrentSaveEnabled
        {
            get => _isCurrentSaveEnabled;
            set
            {
                _isCurrentSaveEnabled = value;
                System.Diagnostics.Debug.WriteLine($"current save enabled: {_isCurrentSaveEnabled}");
            }
        }

        private bool _isAllSaveEnabled = false;
        public bool IsAllSaveEnabled
        {
            get => _isAllSaveEnabled;
            set
            {
                _isAllSaveEnabled = value;
                System.Diagnostics.Debug.WriteLine($"all save enabled: {_isAllSaveEnabled}");
            }
        }

        private int _enabledIndex = 0;
        public int EnabledIndex
        {
            get => _enabledIndex;
            set
            {
                _enabledIndex = value;
                System.Diagnostics.Debug.WriteLine($"index: {_enabledIndex}");
            }
        }

        private int _saveInterval = 20;
        public int SaveInterval
        {
            get => _saveInterval;
            set
            {
                _saveInterval = value;
                System.Diagnostics.Debug.WriteLine($"save interval: {_saveInterval}");
            }
        }

        private string _settingsPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AutoSaveAddIn", "autoSaveAddInSettings.json");
        private UserSettings _settings;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                LoadSettings();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Startup error: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                SaveSettings();
            }
            catch (Exception ex)
            {

            }
        }

        private void LoadSettings()
        {
            try
            {
                _settings = JsonSerializationUtility.DeserializeFromJson<UserSettings>(_settingsPath);

                if (_settings == null)
                {
                    _settings = new UserSettings();
                }
            }
            catch
            {
                _settings = new UserSettings();
            }
            
            _enabledIndex = _settings.EnabledIndex;
            _isCurrentSaveEnabled = _settings.IsCurrentSaving;
            _isAllSaveEnabled = _settings.IsAllSaving;
            _saveInterval = _settings.SaveInterval;
        }

        private void SaveSettings()
        {
            try
            {
                _settings.EnabledIndex = _enabledIndex;
                _settings.IsCurrentSaving = _isCurrentSaveEnabled;
                _settings.IsAllSaving = _isAllSaveEnabled;
                _settings.SaveInterval = _saveInterval;

                JsonSerializationUtility.SerializeToJson(_settings, _settingsPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        #region Код, автоматически созданный VSTO 

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
