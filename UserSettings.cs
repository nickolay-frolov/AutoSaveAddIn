using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoSaveAddIn
{
    public class UserSettings
    {
        public int EnabledIndex { get; set; } = 0;
        public int SaveInterval { get; set; } = 20; // Интервал автосохранения в секундах
        public bool IsCurrentSaving { get; set; } = true;
        public bool IsAllSaving { get; set; } = true;
    }
}