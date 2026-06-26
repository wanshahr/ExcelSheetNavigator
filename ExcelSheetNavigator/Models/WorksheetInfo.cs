using System;

namespace ExcelSheetNavigator.Models
{
    /// <summary>
    /// Lightweight metadata describing a single worksheet. Holds only the
    /// information required for navigation - no cell data is ever read.
    /// </summary>
    public sealed class WorksheetInfo
    {
        /// <summary>The worksheet name as shown on the Excel tab.</summary>
        public string Name { get; set; }

        /// <summary>The 1-based tab index of the worksheet within the workbook.</summary>
        public int Index { get; set; }

        /// <summary>True when this worksheet is the currently active sheet.</summary>
        public bool IsActive { get; set; }

        /// <summary>
        /// True when the worksheet is visible. Hidden / very-hidden sheets are
        /// listed but cannot be activated without first being un-hidden.
        /// </summary>
        public bool IsVisible { get; set; }

        public override string ToString()
        {
            return Name ?? string.Empty;
        }
    }
}
