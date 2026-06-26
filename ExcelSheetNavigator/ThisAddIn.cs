using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelSheetNavigator.Controls;
using ExcelSheetNavigator.Models;
using ExcelSheetNavigator.Ribbon;
using ExcelSheetNavigator.Services;
using Tools = Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetNavigator
{
    public partial class ThisAddIn
    {
        private const string PaneTitle = "Sheet Navigator";
        private const int MinPanelWidth = 140;
        private const int MaxPanelWidth = 600;

        // One task pane per Excel window (Excel 2013+ uses one window per workbook),
        // keyed by the window handle, so the navigator appears in every open file.
        private readonly Dictionary<int, Tools.CustomTaskPane> _panes =
            new Dictionary<int, Tools.CustomTaskPane>();

        private ExcelService _excelService;
        private NavigatorRibbon _ribbon;

        private string _dockPosition = "Right";
        private int _panelWidth = 220;
        private bool _navigatorVisible = true;
        private bool _compact;

        // Shared across every pane and persisted between sessions.
        private readonly HashSet<string> _favorites = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _excelService = new ExcelService(this.Application);
            LoadSettings();
            WireExcelEvents();

            if (_navigatorVisible)
            {
                Tools.CustomTaskPane pane = EnsurePaneForActiveWindow();
                if (pane != null)
                {
                    PopulateControl(pane.Control as SheetNavigatorControl);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            CaptureCurrentWidth();
            TrySaveSettings();
            UnwireExcelEvents();

            foreach (Tools.CustomTaskPane pane in _panes.Values)
            {
                try { pane.Dispose(); }
                catch (COMException) { }
                catch (ObjectDisposedException) { }
            }
            _panes.Clear();
        }

        /// <summary>Provides the ribbon toggle button to Office.</summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new NavigatorRibbon();
            return _ribbon;
        }

        #region Settings

        private void LoadSettings()
        {
            try
            {
                _dockPosition = NormalizeDock(Properties.Settings.Default.DockPosition);
                _panelWidth = Clamp(Properties.Settings.Default.PanelWidth, MinPanelWidth, MaxPanelWidth);
                _navigatorVisible = Properties.Settings.Default.PaneVisible;
                _compact = Properties.Settings.Default.Compact;
                LoadStringSet(Properties.Settings.Default.Favorites, _favorites);
            }
            catch (Exception)
            {
                // Fall back to the defaults already assigned to the fields.
            }
        }

        private void TrySaveSettings()
        {
            try
            {
                Properties.Settings.Default.DockPosition = _dockPosition;
                Properties.Settings.Default.PanelWidth = _panelWidth;
                Properties.Settings.Default.PaneVisible = _navigatorVisible;
                Properties.Settings.Default.Compact = _compact;
                Properties.Settings.Default.Favorites = JoinStringSet(_favorites, false);
                Properties.Settings.Default.Save();
            }
            catch (Exception)
            {
                // Persisting settings is best-effort.
            }
        }

        private static void LoadStringSet(string raw, HashSet<string> target)
        {
            target.Clear();
            if (string.IsNullOrEmpty(raw))
            {
                return;
            }
            foreach (var part in raw.Split('\n'))
            {
                string value = part.Trim('\r');
                if (value.Length > 0)
                {
                    target.Add(value);
                }
            }
        }

        private static string JoinStringSet(HashSet<string> source, bool excludeSentinels)
        {
            var parts = new List<string>();
            foreach (var value in source)
            {
                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }
                if (excludeSentinels && value.IndexOf('\u0000') >= 0)
                {
                    continue; // sentinel section keys are not XML-persistable
                }
                parts.Add(value);
            }
            return string.Join("\n", parts);
        }

        private void CaptureCurrentWidth()
        {
            foreach (Tools.CustomTaskPane pane in _panes.Values)
            {
                try
                {
                    if (IsHorizontalDock(pane.DockPosition) && pane.Width > 0)
                    {
                        _panelWidth = Clamp(pane.Width, MinPanelWidth, MaxPanelWidth);
                        break;
                    }
                }
                catch (COMException) { }
                catch (ObjectDisposedException) { }
            }
        }

        #endregion

        #region Task pane management

        private Tools.CustomTaskPane EnsurePaneForActiveWindow()
        {
            Excel.Window window = null;
            try { window = this.Application.ActiveWindow; }
            catch (COMException) { return null; }

            if (window == null)
            {
                return null;
            }

            int hwnd;
            try { hwnd = window.Hwnd; }
            catch (COMException) { return null; }

            PruneClosedPanes();

            Tools.CustomTaskPane existing;
            if (_panes.TryGetValue(hwnd, out existing))
            {
                return existing;
            }

            return CreatePane(hwnd, window);
        }

        private Tools.CustomTaskPane CreatePane(int hwnd, Excel.Window window)
        {
            var control = new SheetNavigatorControl();
            control.SheetActivationRequested += Control_SheetActivationRequested;
            control.RefreshRequested += Control_RefreshRequested;
            control.DockPositionChangeRequested += Control_DockPositionChangeRequested;
            control.PreferencesChanged += Control_PreferencesChanged;
            control.SetFavorites(_favorites);
            control.SetCompactMode(_compact);

            Tools.CustomTaskPane pane;
            try
            {
                pane = this.CustomTaskPanes.Add(control, PaneTitle, window);
            }
            catch (Exception)
            {
                control.Dispose();
                return null;
            }

            try
            {
                pane.DockPosition = ParseDock(_dockPosition);
                ApplyWidth(pane);
            }
            catch (COMException) { }

            pane.DockPositionChanged += Pane_DockPositionChanged;
            pane.VisibleChanged += Pane_VisibleChanged;
            _panes[hwnd] = pane;

            control.SetDockSelection(_dockPosition);

            try { pane.Visible = _navigatorVisible; }
            catch (COMException) { }

            return pane;
        }

        private void PruneClosedPanes()
        {
            List<int> dead = null;
            foreach (KeyValuePair<int, Tools.CustomTaskPane> entry in _panes)
            {
                bool alive = true;
                try
                {
                    object w = entry.Value.Window;
                    alive = w != null;
                }
                catch (COMException) { alive = false; }
                catch (ObjectDisposedException) { alive = false; }
                catch (InvalidComObjectException) { alive = false; }

                if (!alive)
                {
                    if (dead == null)
                    {
                        dead = new List<int>();
                    }
                    dead.Add(entry.Key);
                }
            }

            if (dead == null)
            {
                return;
            }

            foreach (int key in dead)
            {
                try { _panes[key].Dispose(); }
                catch (COMException) { }
                catch (ObjectDisposedException) { }
                _panes.Remove(key);
            }
        }

        private void ApplyWidth(Tools.CustomTaskPane pane)
        {
            try
            {
                if (IsHorizontalDock(pane.DockPosition))
                {
                    pane.Width = _panelWidth;
                }
            }
            catch (COMException) { }
            catch (ArgumentException) { }
        }

        private SheetNavigatorControl GetActivePaneControl()
        {
            Tools.CustomTaskPane pane = EnsurePaneForActiveWindow();
            return pane == null ? null : pane.Control as SheetNavigatorControl;
        }

        private void PopulateControl(SheetNavigatorControl control)
        {
            if (control == null)
            {
                return;
            }

            List<WorksheetInfo> sheets = _excelService.GetWorksheets();
            control.SetSheets(sheets);
        }

        private void RefreshActivePane()
        {
            if (!_navigatorVisible)
            {
                return;
            }

            Tools.CustomTaskPane pane = EnsurePaneForActiveWindow();
            if (pane != null)
            {
                PopulateControl(pane.Control as SheetNavigatorControl);
            }
        }

        #endregion

        #region Public API used by the ribbon

        /// <summary>Shows or hides the navigator panel in the active window.</summary>
        public void ShowNavigator(bool show)
        {
            _navigatorVisible = show;

            if (show)
            {
                Tools.CustomTaskPane pane = EnsurePaneForActiveWindow();
                if (pane != null)
                {
                    PopulateControl(pane.Control as SheetNavigatorControl);
                }
            }

            foreach (Tools.CustomTaskPane pane in _panes.Values)
            {
                try { pane.Visible = show; }
                catch (COMException) { }
            }

            TrySaveSettings();
            InvalidateRibbon();
        }

        /// <summary>Whether the navigator panel is currently visible.</summary>
        public bool IsNavigatorVisible
        {
            get { return _navigatorVisible; }
        }

        /// <summary>Applies a new dock position to every pane and persists it.</summary>
        public void SetDockPosition(string dock)
        {
            _dockPosition = NormalizeDock(dock);
            Office.MsoCTPDockPosition position = ParseDock(_dockPosition);

            foreach (Tools.CustomTaskPane pane in _panes.Values)
            {
                try
                {
                    pane.DockPosition = position;
                    ApplyWidth(pane);
                }
                catch (COMException) { }
            }

            foreach (Tools.CustomTaskPane pane in _panes.Values)
            {
                var control = pane.Control as SheetNavigatorControl;
                if (control != null)
                {
                    control.SetDockSelection(_dockPosition);
                }
            }

            TrySaveSettings();
        }

        #endregion

        #region Control event handlers

        private void Control_SheetActivationRequested(object sender, SheetEventArgs e)
        {
            bool activated = _excelService.ActivateWorksheet(e.SheetName);
            var control = sender as SheetNavigatorControl;
            if (!activated && control != null)
            {
                control.ShowError("Unable to activate sheet");
            }
            // On success the Excel SheetActivate event refreshes the highlight.
        }

        private void Control_RefreshRequested(object sender, EventArgs e)
        {
            RefreshActivePane();
        }

        private void Control_DockPositionChangeRequested(object sender, DockEventArgs e)
        {
            SetDockPosition(e.DockPosition);
        }

        private void Control_PreferencesChanged(object sender, EventArgs e)
        {
            var source = sender as SheetNavigatorControl;
            if (source == null)
            {
                return;
            }

            _favorites.Clear();
            foreach (var name in source.GetFavorites())
            {
                _favorites.Add(name);
            }
            _compact = source.CompactMode;

            // Keep every other open pane in sync.
            foreach (Tools.CustomTaskPane pane in _panes.Values)
            {
                var control = pane.Control as SheetNavigatorControl;
                if (control != null && !ReferenceEquals(control, source))
                {
                    control.SetFavorites(_favorites);
                    control.SetCompactMode(_compact);
                }
            }

            TrySaveSettings();
        }

        private void Pane_DockPositionChanged(object sender, EventArgs e)
        {
            var pane = sender as Tools.CustomTaskPane;
            if (pane == null)
            {
                return;
            }

            try
            {
                _dockPosition = DockToString(pane.DockPosition);
                CaptureCurrentWidth();
                var control = pane.Control as SheetNavigatorControl;
                if (control != null)
                {
                    control.SetDockSelection(_dockPosition);
                }
                TrySaveSettings();
            }
            catch (COMException) { }
        }

        private void Pane_VisibleChanged(object sender, EventArgs e)
        {
            var pane = sender as Tools.CustomTaskPane;
            if (pane == null)
            {
                return;
            }

            try { _navigatorVisible = pane.Visible; }
            catch (COMException) { return; }

            InvalidateRibbon();
        }

        #endregion

        #region Excel event handlers

        private void OnWorkbookOpen(Excel.Workbook wb)
        {
            RefreshActivePane();
        }

        private void OnNewWorkbook(Excel.Workbook wb)
        {
            RefreshActivePane();
        }

        private void OnWorkbookActivate(Excel.Workbook wb)
        {
            RefreshActivePane();
        }

        private void OnWindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            RefreshActivePane();
        }

        private void OnSheetActivate(object sh)
        {
            if (!_navigatorVisible)
            {
                return;
            }

            string name = GetSheetName(sh);
            SheetNavigatorControl control = GetActivePaneControl();
            if (control != null && name != null)
            {
                control.SetActiveSheet(name);
            }
        }

        private void OnWorkbookNewSheet(Excel.Workbook wb, object sh)
        {
            RefreshActivePane();
        }

        private void OnSheetBeforeDelete(object sh)
        {
            if (!_navigatorVisible)
            {
                return;
            }

            // The sheet still exists during this event; defer the refresh so the
            // list reflects the workbook after the deletion completes.
            SheetNavigatorControl control = GetActivePaneControl();
            if (control != null && control.IsHandleCreated)
            {
                control.BeginInvoke(new Action(RefreshActivePane));
            }
            else
            {
                RefreshActivePane();
            }
        }

        #endregion

        #region Event wiring

        private void WireExcelEvents()
        {
            Excel.Application app = this.Application;
            app.WorkbookOpen += OnWorkbookOpen;
            ((Excel.AppEvents_Event)app).NewWorkbook += OnNewWorkbook;
            app.WorkbookActivate += OnWorkbookActivate;
            app.WindowActivate += OnWindowActivate;
            app.SheetActivate += OnSheetActivate;
            app.WorkbookNewSheet += OnWorkbookNewSheet;
            app.SheetBeforeDelete += OnSheetBeforeDelete;
        }

        private void UnwireExcelEvents()
        {
            try
            {
                Excel.Application app = this.Application;
                app.WorkbookOpen -= OnWorkbookOpen;
                ((Excel.AppEvents_Event)app).NewWorkbook -= OnNewWorkbook;
                app.WorkbookActivate -= OnWorkbookActivate;
                app.WindowActivate -= OnWindowActivate;
                app.SheetActivate -= OnSheetActivate;
                app.WorkbookNewSheet -= OnWorkbookNewSheet;
                app.SheetBeforeDelete -= OnSheetBeforeDelete;
            }
            catch (COMException) { }
        }

        #endregion

        #region Helpers

        private void InvalidateRibbon()
        {
            if (_ribbon != null)
            {
                _ribbon.InvalidateToggle();
            }
        }

        private static string GetSheetName(object sh)
        {
            var worksheet = sh as Excel.Worksheet;
            if (worksheet == null)
            {
                return null;
            }
            try { return worksheet.Name; }
            catch (COMException) { return null; }
        }

        private static bool IsHorizontalDock(Office.MsoCTPDockPosition position)
        {
            return position == Office.MsoCTPDockPosition.msoCTPDockPositionLeft ||
                   position == Office.MsoCTPDockPosition.msoCTPDockPositionRight;
        }

        private static Office.MsoCTPDockPosition ParseDock(string dock)
        {
            switch ((dock ?? "Right").Trim().ToLowerInvariant())
            {
                case "left":
                    return Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                case "floating":
                    return Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
                default:
                    return Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            }
        }

        private static string DockToString(Office.MsoCTPDockPosition position)
        {
            switch (position)
            {
                case Office.MsoCTPDockPosition.msoCTPDockPositionLeft:
                    return "Left";
                case Office.MsoCTPDockPosition.msoCTPDockPositionFloating:
                    return "Floating";
                default:
                    return "Right";
            }
        }

        private static string NormalizeDock(string dock)
        {
            return DockToString(ParseDock(dock));
        }

        private static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
