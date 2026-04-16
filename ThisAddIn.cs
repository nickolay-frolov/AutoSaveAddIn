using System;
using System.IO;
using System.Timers;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoSaveAddIn
{
    public partial class ThisAddIn
    {
        private Timer _saveTimer;
        private Timer _debugTimer;

        private bool _hasChanges = false;
        private bool _saveAsWasShown = false;

        private DateTime _lastRestartTime;
        private int _currentIntervalMs;

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
                RestartTimer();
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

                _saveTimer = new Timer();
                _saveTimer.AutoReset = false;
                _saveTimer.Elapsed += SaveTimer_Elapsed;

#if DEBUG
                _debugTimer = new Timer(1000);
                _debugTimer.AutoReset = true;
                _debugTimer.Elapsed += DebugTimer_Elapsed;
                _debugTimer.Start();
#endif

                Application.SheetChange += Application_SheetChange;
                Application.SheetSelectionChange += Application_SheetSelectionChange;
                Application.WindowDeactivate += Application_WindowDeactivate;
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
            _saveInterval = _settings.SaveInterval;
        }

        private void SaveSettings()
        {
            try
            {
                _settings.EnabledIndex = _enabledIndex;
                _settings.IsCurrentSaving = _isCurrentSaveEnabled;
                _settings.SaveInterval = _saveInterval;

                JsonSerializationUtility.SerializeToJson(_settings, _settingsPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
            }
        }

        private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            if (EnabledIndex == 0)
            {
                return;
            }

            _hasChanges = true;
            RestartTimer();
        }

        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            if (EnabledIndex == 0) return;

            _hasChanges = true;
            RestartTimer();
        }

        private void Application_WindowDeactivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            if (EnabledIndex == 0)
            {
                return;
            }

            if (_hasChanges)
            {
                PerformSave();
            }
        }

        private void RestartTimer()
        {
            _saveTimer.Stop();

            _currentIntervalMs = SaveInterval * 1000;
            _saveTimer.Interval = _currentIntervalMs;

            _lastRestartTime = DateTime.Now;

            _saveTimer.Start();
        }


        private void SaveTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_hasChanges)
            {
                PerformSave();
            }
        }

#if DEBUG
        private void DebugTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_saveTimer.Enabled)
            {
                var elapsed = (DateTime.Now - _lastRestartTime).TotalMilliseconds;
                var remaining = _currentIntervalMs - elapsed;

                if (remaining < 0)
                    remaining = 0;

                System.Diagnostics.Debug.WriteLine(
                    $"[AutoSave DEBUG] Осталось: {remaining / 1000:0.0} сек");
            }
        }
#endif

        private void PerformSave()
        {
            try
            {if (IsCurrentSaveEnabled)
                {
                    var wb = Application.ActiveWorkbook;
                    if (wb != null)
                        SaveWorkbook(wb);
                }
            }
            catch
            {

            }

            _hasChanges = false;
        }

        private void SaveWorkbook(Excel.Workbook wb)
        {
            if (wb.ReadOnly)
            {
                return;
            }

            if (string.IsNullOrEmpty(wb.Path))
            {
                if (_saveAsWasShown)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Автосохранение отключено для '{wb.Name}' до тех пор, пока пользователь не сохранит файл вручную.");
                    return;
                }

                _saveAsWasShown = true;

                try
                {
                    bool result = wb.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();

                    if (!result)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"Юзер закрыл \"Сохранить как...\" для '{wb.Name}'. Автосохранение не будет работать.");

                        var excelHandle = new IntPtr(Application.Hwnd);
                        var owner = new ExcelWindow(excelHandle);

                        System.Windows.Forms.MessageBox.Show(
                            owner,
                            "Файл не был сохранён. Автосохранение не работает, пока вы не сохраните книгу вручную.",
                            "Автосохранение отключено",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning
                        );


                        return;
                    }

                    System.Diagnostics.Debug.WriteLine($"Пользователь сохранил новую книгу: {wb.Name}");
                }
                catch
                {
                    return;
                }
            }

            wb.Save();
            System.Diagnostics.Debug.WriteLine($"Сохранена книга: {wb.Name}");
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
