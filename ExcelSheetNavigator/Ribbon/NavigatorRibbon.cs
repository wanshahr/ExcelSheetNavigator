using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ExcelSheetNavigator.Ribbon
{
    /// <summary>
    /// Ribbon UI (defined in NavigatorRibbon.xml) that adds a "Sheet Navigator"
    /// toggle button to the Home tab so the panel can be shown or hidden, and
    /// reopened after the user closes it.
    /// </summary>
    [ComVisible(true)]
    public class NavigatorRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelSheetNavigator.Ribbon.NavigatorRibbon.xml");
        }

        /// <summary>Called by Office when the ribbon is loaded.</summary>
        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        /// <summary>Toggle button action - shows or hides the navigator panel.</summary>
        public void OnToggleNavigator(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.ShowNavigator(pressed);
        }

        /// <summary>Keeps the toggle button's pressed state in sync with the panel.</summary>
        public bool GetTogglePressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.IsNavigatorVisible;
        }

        /// <summary>Requests a refresh of the toggle button's pressed state.</summary>
        public void InvalidateToggle()
        {
            try
            {
                if (_ribbon != null)
                {
                    _ribbon.InvalidateControl("btnToggleNavigator");
                }
            }
            catch (COMException)
            {
                // Ribbon not available yet; ignore.
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
