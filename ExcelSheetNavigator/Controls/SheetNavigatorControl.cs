using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelSheetNavigator.Models;

namespace ExcelSheetNavigator.Controls
{
    /// <summary>
    /// Event arguments carrying a single worksheet name.
    /// </summary>
    public sealed class SheetEventArgs : EventArgs
    {
        public SheetEventArgs(string sheetName)
        {
            SheetName = sheetName;
        }

        public string SheetName { get; }
    }

    /// <summary>
    /// Event arguments carrying a requested dock position ("Left", "Right" or "Floating").
    /// </summary>
    public sealed class DockEventArgs : EventArgs
    {
        public DockEventArgs(string dockPosition)
        {
            DockPosition = dockPosition;
        }

        public string DockPosition { get; }
    }

    /// <summary>
    /// The dockable navigation panel. Renders a searchable, flat worksheet list
    /// with favorites, hover/active row states and a status bar. It contains no
    /// Excel interop code - it raises events that the add-in turns into actions.
    /// </summary>
    public partial class SheetNavigatorControl : UserControl
    {
        // Native cue-banner support so the search box shows placeholder text.
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        private const int EM_SETCUEBANNER = 0x1501;

        // ---- Color system (modern slate palette) --------------------------
        private static readonly Color ColorPrimary = Color.FromArgb(37, 99, 235);    // #2563EB blue-600
        private static readonly Color ColorLight = Color.FromArgb(239, 246, 255);     // #EFF6FF blue-50 (active row)
        private static readonly Color ColorBackground = Color.White;                  // #FFFFFF surface
        private static readonly Color ColorText = Color.FromArgb(15, 23, 42);         // #0F172A slate-900
        private static readonly Color ColorMuted = Color.FromArgb(100, 116, 139);     // #64748B slate-500
        private static readonly Color ColorFaint = Color.FromArgb(148, 163, 184);     // #94A3B8 slate-400
        private static readonly Color ColorBorder = Color.FromArgb(226, 232, 240);    // #E2E8F0 slate-200
        private static readonly Color ColorError = Color.FromArgb(220, 38, 38);       // #DC2626 red-600
        private static readonly Color ColorHover = Color.FromArgb(241, 245, 249);     // #F1F5F9 slate-100 (hover)
        private static readonly Color ColorStar = Color.FromArgb(245, 158, 11);       // #F59E0B amber-500

        // Row layout metrics. Icon/name offsets and the row font vary with the
        // active density (see ApplyDensity); the rest stay constant.
        private const int AccentWidth = 3;
        private const int StarBoxWidth = 26;
        private int _iconLeft = 12;
        private int _iconSize = 15;
        private int _nameLeft = 34;
        private bool _compact;

        // Owner-draw fonts (disposed in the Designer's Dispose).
        private Font _boldFont;
        private Font _italicFont;
        private Font _ownedListFont;

        // Cached owner-draw brushes/pens - allocated once, never per row repaint.
        private SolidBrush _brushActiveBg;
        private SolidBrush _brushHoverBg;
        private SolidBrush _brushRowBg;
        private SolidBrush _brushAccent;
        private SolidBrush _brushStar;
        private Pen _penIconMuted;
        private Pen _penIconActive;
        private Pen _penStarOutline;

        // Debounce timer for the search box (disposed in the Designer's Dispose).
        private System.Windows.Forms.Timer _searchTimer;

        // Signature of the currently displayed sheet set, to skip redundant rebuilds.
        private string _sheetsSignature;

        private readonly List<WorksheetInfo> _allSheets = new List<WorksheetInfo>();
        private readonly HashSet<string> _favorites = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private string _activeSheetName;
        private string _statusText = "Ready";
        private bool _statusIsError;
        private int _hoverIndex = -1;
        private int _tooltipIndex = -1;
        private int _matchedCount;
        private bool _suppressDockEvent;

        /// <summary>Raised when the user requests activation of a worksheet.</summary>
        public event EventHandler<SheetEventArgs> SheetActivationRequested;

        /// <summary>Raised when the user clicks the refresh button.</summary>
        public event EventHandler RefreshRequested;

        /// <summary>Raised when the user picks a different dock position.</summary>
        public event EventHandler<DockEventArgs> DockPositionChangeRequested;

        /// <summary>Raised when favorites change, so the host can persist them.</summary>
        public event EventHandler PreferencesChanged;

        public SheetNavigatorControl()
        {
            InitializeComponent();

            _ownedListFont = lstSheets.Font;
            _boldFont = new Font(lstSheets.Font, FontStyle.Bold);
            _italicFont = new Font(lstSheets.Font, FontStyle.Italic);

            _brushActiveBg = new SolidBrush(ColorLight);
            _brushHoverBg = new SolidBrush(ColorHover);
            _brushRowBg = new SolidBrush(ColorBackground);
            _brushAccent = new SolidBrush(ColorPrimary);
            _brushStar = new SolidBrush(ColorStar);
            _penIconMuted = new Pen(ColorFaint, 1.4f);
            _penIconActive = new Pen(ColorPrimary, 1.4f);
            _penStarOutline = new Pen(ColorMuted, 1.3f);

            // Reduce flicker on the owner-drawn list.
            typeof(ListBox).InvokeMember(
                "DoubleBuffered",
                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null,
                lstSheets,
                new object[] { true });

            lstSheets.DrawItem += LstSheets_DrawItem;
            lstSheets.MouseClick += LstSheets_MouseClick;
            lstSheets.MouseDoubleClick += LstSheets_MouseClick;
            lstSheets.MouseMove += LstSheets_MouseMove;
            lstSheets.MouseLeave += LstSheets_MouseLeave;
            lstSheets.KeyDown += LstSheets_KeyDown;

            txtSearch.TextChanged += TxtSearch_TextChanged;
            btnClear.Click += BtnClear_Click;
            btnRefresh.Click += BtnRefresh_Click;
            btnDensity.Click += BtnDensity_Click;
            cmbDock.SelectedIndexChanged += CmbDock_SelectedIndexChanged;

            txtSearch.BackColor = ColorBackground;
            pnlSearch.Paint += PnlSearch_Paint;
            pnlSearch.Resize += (s, e) => pnlSearch.Invalidate();
            txtSearch.GotFocus += (s, e) => pnlSearch.Invalidate();
            txtSearch.LostFocus += (s, e) => pnlSearch.Invalidate();

            _searchTimer = new System.Windows.Forms.Timer { Interval = 250 };
            _searchTimer.Tick += SearchTimer_Tick;

            UpdateDensityButton();
            UpdateStatusLabel();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            SendMessage(txtSearch.Handle, EM_SETCUEBANNER, (IntPtr)1, "Search sheets...");
        }

        #region Public API

        /// <summary>
        /// Replaces the displayed worksheet list. Preserves the search filter,
        /// favorites and the active-sheet highlight.
        /// </summary>
        public void SetSheets(IList<WorksheetInfo> sheets)
        {
            string signature = BuildSignature(sheets);

            // Same set of sheets (e.g. switching between windows of one workbook):
            // skip the full rebuild and just resync the active highlight cheaply.
            if (sheets != null && signature == _sheetsSignature && _allSheets.Count > 0)
            {
                string activeName = null;
                foreach (var s in sheets)
                {
                    if (s.IsActive)
                    {
                        activeName = s.Name;
                    }
                }
                SetActiveSheet(activeName);
                return;
            }

            _sheetsSignature = signature;
            _allSheets.Clear();
            if (sheets != null)
            {
                _allSheets.AddRange(sheets);
            }

            _activeSheetName = null;
            foreach (var s in _allSheets)
            {
                if (s.IsActive)
                {
                    _activeSheetName = s.Name;
                }
            }

            _statusText = "Ready";
            _statusIsError = false;

            BuildRows();
            UpdateStatusLabel();
        }

        /// <summary>Updates which sheet is rendered as active and highlights it.</summary>
        public void SetActiveSheet(string sheetName)
        {
            _activeSheetName = sheetName;
            foreach (var s in _allSheets)
            {
                s.IsActive = string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase);
            }

            _statusText = "Ready";
            _statusIsError = false;

            SelectSheetRow(sheetName);
            lstSheets.Invalidate();
            UpdateStatusLabel();
        }

        /// <summary>Sets a normal status message (e.g. "Ready").</summary>
        public void SetStatus(string text)
        {
            _statusText = string.IsNullOrEmpty(text) ? "Ready" : text;
            _statusIsError = false;
            UpdateStatusLabel();
        }

        /// <summary>Shows an error message in the status bar (red).</summary>
        public void ShowError(string text)
        {
            _statusText = string.IsNullOrEmpty(text) ? "Error" : text;
            _statusIsError = true;
            UpdateStatusLabel();
        }

        /// <summary>Sets the dock dropdown without raising a change event.</summary>
        public void SetDockSelection(string dockPosition)
        {
            _suppressDockEvent = true;
            try
            {
                int index = cmbDock.Items.IndexOf(dockPosition);
                cmbDock.SelectedIndex = index >= 0 ? index : cmbDock.Items.IndexOf("Right");
            }
            finally
            {
                _suppressDockEvent = false;
            }
        }

        /// <summary>Replaces the set of favorite sheet names. Does not raise PreferencesChanged.</summary>
        public void SetFavorites(IEnumerable<string> names)
        {
            _favorites.Clear();
            if (names != null)
            {
                foreach (var n in names)
                {
                    if (!string.IsNullOrEmpty(n))
                    {
                        _favorites.Add(n);
                    }
                }
            }
            BuildRows();
        }

        /// <summary>Returns a snapshot of the current favorite sheet names.</summary>
        public IEnumerable<string> GetFavorites()
        {
            return new List<string>(_favorites);
        }

        #endregion

        #region Row building (search + favorites)

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            btnClear.Visible = txtSearch.TextLength > 0;
            // Debounce: rebuild only after the user pauses typing.
            _searchTimer.Stop();
            _searchTimer.Start();
        }

        private void SearchTimer_Tick(object sender, EventArgs e)
        {
            _searchTimer.Stop();
            BuildRows();
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            txtSearch.Clear();
            txtSearch.Focus();
        }

        private void BuildRows()
        {
            string filter = txtSearch.Text == null ? string.Empty : txtSearch.Text.Trim();

            // Flat list: favorites pinned on top, everything else in Excel tab order.
            var favorites = new List<WorksheetInfo>();
            var others = new List<WorksheetInfo>();
            foreach (var s in _allSheets)
            {
                if (filter.Length > 0 &&
                    !(s.Name != null && s.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    continue;
                }
                if (_favorites.Contains(s.Name))
                {
                    favorites.Add(s);
                }
                else
                {
                    others.Add(s);
                }
            }
            _matchedCount = favorites.Count + others.Count;

            lstSheets.BeginUpdate();
            try
            {
                lstSheets.Items.Clear();
                foreach (var s in favorites)
                {
                    lstSheets.Items.Add(s);
                }
                foreach (var s in others)
                {
                    lstSheets.Items.Add(s);
                }
            }
            finally
            {
                lstSheets.EndUpdate();
            }

            if (!string.IsNullOrEmpty(_activeSheetName))
            {
                SelectSheetRow(_activeSheetName);
            }

            UpdateSummary(filter.Length > 0);
            UpdateEmptyState(filter.Length > 0);
        }

        private static string BuildSignature(IList<WorksheetInfo> sheets)
        {
            if (sheets == null || sheets.Count == 0)
            {
                return string.Empty;
            }
            var sb = new System.Text.StringBuilder();
            foreach (var s in sheets)
            {
                sb.Append(s.Name).Append(s.IsVisible ? '1' : '0').Append('\u0001');
            }
            return sb.ToString();
        }

        private void SelectSheetRow(string sheetName)
        {
            for (int i = 0; i < lstSheets.Items.Count; i++)
            {
                if (lstSheets.Items[i] is WorksheetInfo info &&
                    string.Equals(info.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    if (lstSheets.SelectedIndex != i)
                    {
                        lstSheets.SelectedIndex = i;
                    }
                    return;
                }
            }
        }

        #endregion

        #region Summary / status / empty state

        private void UpdateSummary(bool filtering)
        {
            int total = _allSheets.Count;
            int hidden = 0;
            foreach (var s in _allSheets)
            {
                if (!s.IsVisible)
                {
                    hidden++;
                }
            }

            if (total == 0)
            {
                lblSummary.Text = "No worksheets";
                return;
            }

            string text;
            if (filtering && _matchedCount != total)
            {
                text = _matchedCount + " of " + total + " sheets";
            }
            else
            {
                text = total + (total == 1 ? " sheet" : " sheets");
                if (hidden > 0)
                {
                    text += " (" + hidden + " hidden)";
                }
            }
            lblSummary.Text = text;
        }

        private void UpdateStatusLabel()
        {
            string text = _statusText;
            if (!string.IsNullOrEmpty(_activeSheetName))
            {
                text += "      |      Active: " + _activeSheetName;
            }
            lblStatus.Text = text;
            lblStatus.ForeColor = _statusIsError ? ColorError : ColorMuted;
        }

        private void UpdateEmptyState(bool filtering)
        {
            if (_allSheets.Count == 0)
            {
                lblEmpty.Text = "No worksheets found";
                lblEmpty.Visible = true;
                lblEmpty.BringToFront();
            }
            else if (_matchedCount == 0)
            {
                lblEmpty.Text = "No sheets match search";
                lblEmpty.Visible = true;
                lblEmpty.BringToFront();
            }
            else
            {
                lblEmpty.Visible = false;
            }
        }

        #endregion

        #region Refresh (with loading feedback)

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            // Brief loading acknowledgement before the (synchronous) refresh.
            _statusText = "Refreshing\u2026";
            _statusIsError = false;
            UpdateStatusLabel();
            lblStatus.Update();

            RefreshRequested?.Invoke(this, EventArgs.Empty);
        }

        #endregion

        #region Search box chrome

        private void PnlSearch_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Rounded input field with a focus accent.
            var r = new Rectangle(8, 6, pnlSearch.Width - 16, pnlSearch.Height - 13);
            if (r.Width <= 0 || r.Height <= 0)
            {
                return;
            }

            bool focused = txtSearch.Focused;
            using (var path = RoundedRect(r, 8))
            using (var fill = new SolidBrush(ColorBackground))
            using (var border = new Pen(focused ? ColorPrimary : ColorBorder, focused ? 1.5f : 1f))
            {
                g.FillPath(fill, path);
                g.DrawPath(border, path);
            }

            // Magnifier glyph on the left (accented while focused).
            int cx = r.Left + 12;
            int cy = pnlSearch.Height / 2;
            using (var pen = new Pen(focused ? ColorPrimary : ColorMuted, 1.6f))
            {
                var circle = new Rectangle(cx - 5, cy - 5, 9, 9);
                g.DrawEllipse(pen, circle);
                g.DrawLine(pen, cx + 3, cy + 3, cx + 7, cy + 7);
            }
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        #endregion

        #region List rendering

        private void LstSheets_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= lstSheets.Items.Count)
            {
                return;
            }

            var info = lstSheets.Items[e.Index] as WorksheetInfo;
            if (info == null)
            {
                return;
            }

            DrawSheetRow(e, info);
        }

        private void DrawSheetRow(DrawItemEventArgs e, WorksheetInfo info)
        {
            bool active = info.IsActive;
            bool hover = e.Index == _hoverIndex;
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            bool favorite = _favorites.Contains(info.Name);

            SolidBrush back = active ? _brushActiveBg : ((hover || selected) ? _brushHoverBg : _brushRowBg);
            e.Graphics.FillRectangle(back, e.Bounds);

            // Left accent bar for the active sheet.
            if (active)
            {
                e.Graphics.FillRectangle(_brushAccent, e.Bounds.X, e.Bounds.Y, AccentWidth, e.Bounds.Height);
            }

            // Sheet icon.
            int iconY = e.Bounds.Y + (e.Bounds.Height - _iconSize) / 2;
            DrawSheetIcon(e.Graphics, e.Bounds.X + _iconLeft, iconY, _iconSize, _iconSize, active ? _penIconActive : _penIconMuted);

            // Star area (shown on hover or when already a favorite).
            int starLeft = e.Bounds.Right - StarBoxWidth;
            if (hover || favorite)
            {
                DrawStar(
                    e.Graphics, starLeft + StarBoxWidth / 2, e.Bounds.Y + e.Bounds.Height / 2, 5.5f,
                    favorite ? _brushStar : null, favorite ? null : _penStarOutline, favorite);
            }

            // Sheet name.
            int nameRight = starLeft - 4;
            var nameRect = new Rectangle(e.Bounds.X + _nameLeft, e.Bounds.Y, nameRight - (e.Bounds.X + _nameLeft), e.Bounds.Height);
            Font font = active ? _boldFont : (info.IsVisible ? lstSheets.Font : _italicFont);
            Color fore = info.IsVisible ? ColorText : ColorMuted;
            TextRenderer.DrawText(
                e.Graphics, info.Name, font, nameRect, fore,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }

        private static void DrawSheetIcon(Graphics g, int x, int y, int w, int h, Pen pen)
        {
            SmoothingMode previous = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Mini spreadsheet: a rounded body split into a 2x2 cell grid.
            var body = new Rectangle(x, y, w - 1, h - 1);
            using (var path = RoundedRect(body, 3))
            {
                g.DrawPath(pen, path);
            }
            int midX = x + (w - 1) / 2;
            int midY = y + (h - 1) / 2;
            g.DrawLine(pen, midX, y + 2, midX, y + h - 3);   // vertical divider
            g.DrawLine(pen, x + 2, midY, x + w - 3, midY);   // horizontal divider
            g.SmoothingMode = previous;
        }

        private static void DrawStar(Graphics g, float cx, float cy, float radius, Brush fill, Pen outline, bool filled)
        {
            const int points = 5;
            var pts = new PointF[points * 2];
            double rotation = -Math.PI / 2;
            double step = Math.PI / points;
            for (int i = 0; i < pts.Length; i++)
            {
                double r = (i % 2 == 0) ? radius : radius * 0.42;
                double angle = rotation + i * step;
                pts[i] = new PointF(
                    cx + (float)(Math.Cos(angle) * r),
                    cy + (float)(Math.Sin(angle) * r));
            }

            SmoothingMode previous = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (filled)
            {
                g.FillPolygon(fill, pts);
            }
            else
            {
                g.DrawPolygon(outline, pts);
            }
            g.SmoothingMode = previous;
        }

        #endregion

        #region Mouse / keyboard interaction

        private void LstSheets_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            int index = lstSheets.IndexFromPoint(e.Location);
            if (index < 0 || index >= lstSheets.Items.Count)
            {
                return;
            }

            var info = lstSheets.Items[index] as WorksheetInfo;
            if (info == null)
            {
                return;
            }

            // Click on the star toggles the favorite instead of activating.
            Rectangle bounds = lstSheets.GetItemRectangle(index);
            int starLeft = bounds.Right - StarBoxWidth;
            if (e.X >= starLeft)
            {
                ToggleFavorite(info.Name);
                return;
            }

            SheetActivationRequested?.Invoke(this, new SheetEventArgs(info.Name));
        }

        private void ToggleFavorite(string sheetName)
        {
            if (_favorites.Contains(sheetName))
            {
                _favorites.Remove(sheetName);
            }
            else
            {
                _favorites.Add(sheetName);
            }
            BuildRows();
            PreferencesChanged?.Invoke(this, EventArgs.Empty);
        }

        private void LstSheets_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (lstSheets.SelectedItem is WorksheetInfo info)
                {
                    SheetActivationRequested?.Invoke(this, new SheetEventArgs(info.Name));
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down)
            {
                MoveSelection(1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                MoveSelection(-1);
                e.Handled = true;
            }
        }

        private void MoveSelection(int direction)
        {
            int count = lstSheets.Items.Count;
            if (count == 0)
            {
                return;
            }

            int index = lstSheets.SelectedIndex + direction;
            if (index < 0)
            {
                index = 0;
            }
            else if (index >= count)
            {
                index = count - 1;
            }
            lstSheets.SelectedIndex = index;
        }

        private void LstSheets_MouseMove(object sender, MouseEventArgs e)
        {
            int index = lstSheets.IndexFromPoint(e.Location);
            if (index != _hoverIndex)
            {
                int previous = _hoverIndex;
                _hoverIndex = index;
                InvalidateItem(previous);
                InvalidateItem(_hoverIndex);
            }

            bool interactive = index >= 0 && index < lstSheets.Items.Count;
            lstSheets.Cursor = interactive ? Cursors.Hand : Cursors.Default;

            UpdateTooltip(index);
        }

        private void LstSheets_MouseLeave(object sender, EventArgs e)
        {
            if (_hoverIndex != -1)
            {
                int previous = _hoverIndex;
                _hoverIndex = -1;
                InvalidateItem(previous);
            }
            _tooltipIndex = -1;
            lstSheets.Cursor = Cursors.Default;
            toolTip.Hide(lstSheets);
        }

        private void InvalidateItem(int index)
        {
            if (index >= 0 && index < lstSheets.Items.Count)
            {
                lstSheets.Invalidate(lstSheets.GetItemRectangle(index));
            }
        }

        private void UpdateTooltip(int index)
        {
            if (index == _tooltipIndex)
            {
                return;
            }

            _tooltipIndex = index;
            if (index < 0 || index >= lstSheets.Items.Count ||
                !(lstSheets.Items[index] is WorksheetInfo info))
            {
                toolTip.Hide(lstSheets);
                return;
            }

            // Show a tooltip only when the name is actually truncated.
            int available = lstSheets.Width - _nameLeft - StarBoxWidth - 4;
            Size measured = TextRenderer.MeasureText(info.Name, lstSheets.Font);
            if (measured.Width > available)
            {
                toolTip.Show(info.Name, lstSheets, _nameLeft, lstSheets.GetItemRectangle(index).Bottom);
            }
            else
            {
                toolTip.Hide(lstSheets);
            }
        }

        #endregion

        #region Density (compact mode)

        /// <summary>Whether the list is currently using the compact (denser) row layout.</summary>
        public bool CompactMode
        {
            get { return _compact; }
        }

        /// <summary>Applies the compact/comfortable row density. Does not raise PreferencesChanged.</summary>
        public void SetCompactMode(bool compact)
        {
            ApplyDensity(compact);
        }

        private void BtnDensity_Click(object sender, EventArgs e)
        {
            ApplyDensity(!_compact);
            PreferencesChanged?.Invoke(this, EventArgs.Empty);
        }

        private void ApplyDensity(bool compact)
        {
            _compact = compact;

            float fontSize;
            if (compact)
            {
                _iconLeft = 10;
                _iconSize = 13;
                _nameLeft = 28;
                fontSize = 8.25f;
                lstSheets.ItemHeight = 21;
            }
            else
            {
                _iconLeft = 12;
                _iconSize = 15;
                _nameLeft = 34;
                fontSize = 9.75f;
                lstSheets.ItemHeight = 30;
            }

            // Swap the row font and its bold/italic variants to match the density.
            var listFont = new Font("Segoe UI", fontSize, FontStyle.Regular);
            lstSheets.Font = listFont;
            if (_ownedListFont != null) { _ownedListFont.Dispose(); }
            _ownedListFont = listFont;
            if (_boldFont != null) { _boldFont.Dispose(); }
            _boldFont = new Font(listFont, FontStyle.Bold);
            if (_italicFont != null) { _italicFont.Dispose(); }
            _italicFont = new Font(listFont, FontStyle.Italic);

            UpdateDensityButton();
            lstSheets.Invalidate();
        }

        private void UpdateDensityButton()
        {
            btnDensity.ForeColor = _compact ? ColorPrimary : ColorMuted;
            btnDensity.BackColor = _compact ? ColorLight : Color.White;
            toolTip.SetToolTip(btnDensity, _compact ? "Switch to comfortable rows" : "Switch to compact rows");
        }

        #endregion

        #region Dock

        private void CmbDock_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppressDockEvent || cmbDock.SelectedItem == null)
            {
                return;
            }

            DockPositionChangeRequested?.Invoke(this, new DockEventArgs(cmbDock.SelectedItem.ToString()));
        }

        #endregion
    }
}
