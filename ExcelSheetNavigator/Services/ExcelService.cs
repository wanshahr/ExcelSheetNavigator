using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelSheetNavigator.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetNavigator.Services
{
    /// <summary>
    /// Business-logic layer that mediates all interaction with the Excel object
    /// model. The UI never touches Excel interop directly - it goes through here.
    /// Only sheet metadata (name / index / state) is read; cell contents are
    /// never accessed.
    /// </summary>
    public sealed class ExcelService
    {
        private readonly Excel.Application _application;

        public ExcelService(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// Returns metadata for every worksheet in the active workbook, ordered
        /// by tab index, with the active sheet flagged. Returns an empty list
        /// when no workbook is open.
        /// </summary>
        public List<WorksheetInfo> GetWorksheets()
        {
            var result = new List<WorksheetInfo>();

            Excel.Workbook workbook = null;
            try
            {
                workbook = _application.ActiveWorkbook;
            }
            catch (COMException)
            {
                // Excel is busy or no workbook context is available.
                return result;
            }

            if (workbook == null)
            {
                return result;
            }

            string activeName = GetActiveSheetName();

            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                int count = worksheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)worksheets[i];
                        result.Add(new WorksheetInfo
                        {
                            Name = sheet.Name,
                            Index = sheet.Index,
                            IsActive = string.Equals(sheet.Name, activeName, StringComparison.Ordinal),
                            IsVisible = sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible
                        });
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            Marshal.ReleaseComObject(sheet);
                        }
                    }
                }
            }
            catch (COMException)
            {
                // Workbook changed underneath us; return whatever we gathered.
            }
            finally
            {
                if (worksheets != null)
                {
                    Marshal.ReleaseComObject(worksheets);
                }
                Marshal.ReleaseComObject(workbook);
            }

            result.Sort((a, b) => a.Index.CompareTo(b.Index));
            return result;
        }

        /// <summary>
        /// Gets the name of the currently active worksheet, or null when none.
        /// </summary>
        public string GetActiveSheetName()
        {
            Excel.Worksheet active = null;
            try
            {
                active = _application.ActiveSheet as Excel.Worksheet;
                return active != null ? active.Name : null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                if (active != null)
                {
                    Marshal.ReleaseComObject(active);
                }
            }
        }

        /// <summary>
        /// Activates the worksheet with the supplied name in the active workbook.
        /// Returns false when the sheet cannot be found or is hidden (Excel
        /// refuses to activate a hidden sheet, and we never modify the workbook
        /// by un-hiding it).
        /// </summary>
        public bool ActivateWorksheet(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            Excel.Workbook workbook = null;
            Excel.Sheets worksheets = null;
            Excel.Worksheet sheet = null;
            try
            {
                workbook = _application.ActiveWorkbook;
                if (workbook == null)
                {
                    return false;
                }

                worksheets = workbook.Worksheets;
                sheet = (Excel.Worksheet)worksheets[name];

                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                {
                    return false;
                }

                sheet.Activate();
                return true;
            }
            catch (COMException)
            {
                // Sheet not found, renamed, or activation rejected by Excel.
                return false;
            }
            finally
            {
                if (sheet != null)
                {
                    Marshal.ReleaseComObject(sheet);
                }
                if (worksheets != null)
                {
                    Marshal.ReleaseComObject(worksheets);
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }

        /// <summary>
        /// Returns the name of the active workbook, or null when none is open.
        /// </summary>
        public string GetActiveWorkbookName()
        {
            Excel.Workbook workbook = null;
            try
            {
                workbook = _application.ActiveWorkbook;
                return workbook != null ? workbook.Name : null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }
    }
}
