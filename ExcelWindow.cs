using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoSaveAddIn
{
    class ExcelWindow : System.Windows.Forms.NativeWindow
    {
        public ExcelWindow(IntPtr handle)
        {
            AssignHandle(handle);
        }
    }
}
