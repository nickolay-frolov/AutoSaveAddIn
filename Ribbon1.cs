using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoSaveAddIn
{
    public partial class Ribbon1
    {
        private Dictionary<int, string> _onOffComboboxItems = new Dictionary<int, string>
        {
            { 0, "Выкл." },
            { 1, "" },
        };

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            onOffCB_LoadText();
            saveIntervalEB_LoadText();
            ToggleUIEnabled();
        }

        private void onOffCB_LoadText()
        {
            if (_onOffComboboxItems.TryGetValue(Globals.ThisAddIn.EnabledIndex, out string text))
            {
                onOffCB.Text = text;
            }
            else
            {
                if(_onOffComboboxItems.TryGetValue(0, out string defaultText))
                {
                    onOffCB.Text = defaultText;
                    Globals.ThisAddIn.EnabledIndex = 0;
                }
            }
        }

        private void saveIntervalEB_LoadText()
        {
            saveIntervalEB.Text = Globals.ThisAddIn.SaveInterval.ToString();
        }

        private void ToggleUIEnabled()
        {
            if (Globals.ThisAddIn.EnabledIndex == 0) // Если "Выкл."
            {
                saveIntervalEB.Enabled = false;
            }
            else // Если "Вкл."
            {
                saveIntervalEB.Enabled = true;
            }
        }

        private void onOffCB_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string selectedText = onOffCB.Text;
            Globals.ThisAddIn.EnabledIndex = _onOffComboboxItems
                .FirstOrDefault(pair => pair.Value == selectedText).Key;
            ToggleUIEnabled();
        }

        private void saveIntervalEB_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (int.TryParse(saveIntervalEB.Text, out int interval))
            {
                Globals.ThisAddIn.SaveInterval = interval;
            }
            else
            {
               saveIntervalEB.Text = Globals.ThisAddIn.SaveInterval.ToString();
            }
        }
    }
}