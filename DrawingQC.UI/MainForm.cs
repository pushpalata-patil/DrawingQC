using System.Diagnostics;

namespace DrawingQC.UI;

public sealed class MainForm : Form
{
    private readonly TextBox _zipBox = new();
    private readonly TextBox _outBox = new();
    private readonly Button _browseZip = new() { Text = "Browse..."};
    private readonly Button _browseOut = new() { Text = "Browse..."};
    private readonly Button _runBtn = new() { Text = "Run QC" };
    private readonly Button _openExcelBtn = new() { Text = "Open Excel", Enabled = false };
    private readonly ProgressBar _progress = new() { Minimum = 0, Maximum = 100 };
    private readonly Label _statusLabel = new() { Text = "Select a zip file to begin." };
    private readonly Label _countsLabel = new() { Text = "" };
    private readonly DataGridView _grid = new();

    private string _lastReportPath = "";

    public MainForm()
    {
        Text = "Drawing QC - PDF vs File-name checker";
        Width = 1000;
        Height = 640;
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(820, 480);
        Font = new Font("Segoe UI", 9f);

        BuildLayout();
        WireEvents();
    }

    private void BuildLayout()
    {
        // ---- Top input panel ----
        var top = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 3,
            RowCount = 3,
            Padding = new Padding(12, 12, 12, 6),
            Height = 130,
            AutoSize = true,
        };
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

        _zipBox.Dock = DockStyle.Fill;
        _outBox.Dock = DockStyle.Fill;
        _browseZip.Dock = DockStyle.Fill;
        _browseOut.Dock = DockStyle.Fill;

        top.Controls.Add(new Label { Text = "Zip file:", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 0);
        top.Controls.Add(_zipBox, 1, 0);
        top.Controls.Add(_browseZip, 2, 0);

        top.Controls.Add(new Label { Text = "Excel out:", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 1);
        top.Controls.Add(_outBox, 1, 1);
        top.Controls.Add(_browseOut, 2, 1);

        var actionPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
        _runBtn.Width = 110;
        _runBtn.Height = 30;
        _runBtn.BackColor = Color.FromArgb(0x37, 0x47, 0x5A);
        _runBtn.ForeColor = Color.White;
        _runBtn.FlatStyle = FlatStyle.Flat;
        _openExcelBtn.Width = 110;
        _openExcelBtn.Height = 30;
        actionPanel.Controls.Add(_runBtn);
        actionPanel.Controls.Add(_openExcelBtn);
        top.Controls.Add(actionPanel, 1, 2);

        Controls.Add(top);

        // ---- Grid ----
        _grid.Dock = DockStyle.Fill;
        _grid.ReadOnly = true;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.RowHeadersVisible = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _grid.BackgroundColor = Color.White;
        _grid.BorderStyle = BorderStyle.None;
        _grid.EnableHeadersVisualStyles = false;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0x37, 0x47, 0x5A);
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
        _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        _grid.ColumnHeadersHeight = 30;

        _grid.Columns.Add(MakeCol("SrNo", "Sr.No", 55, false));
        _grid.Columns.Add(MakeCol("Pdf", "Support PDF", 0, true));
        _grid.Columns.Add(MakeCol("D1", "1st drawing name", 160, false));
        _grid.Columns.Add(MakeCol("D2", "2nd drawing name", 160, false));
        _grid.Columns.Add(MakeCol("Status", "Status", 90, false));
        _grid.Columns.Add(MakeCol("Missing", "Not found in PDF", 170, false));

        var gridHost = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 6, 12, 6) };
        gridHost.Controls.Add(_grid);
        Controls.Add(gridHost);
        gridHost.BringToFront();

        // ---- Bottom status bar ----
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(12, 4, 12, 8) };
        _progress.Dock = DockStyle.Top;
        _progress.Height = 18;
        _statusLabel.Dock = DockStyle.Top;
        _statusLabel.Height = 20;
        _countsLabel.Dock = DockStyle.Top;
        _countsLabel.Height = 20;
        _countsLabel.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
        bottom.Controls.Add(_countsLabel);
        bottom.Controls.Add(_statusLabel);
        bottom.Controls.Add(_progress);
        Controls.Add(bottom);
        bottom.BringToFront();
    }

    private static DataGridViewTextBoxColumn MakeCol(string name, string header, int width, bool fill)
    {
        var col = new DataGridViewTextBoxColumn { Name = name, HeaderText = header };
        if (fill)
        {
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            col.FillWeight = 200;
        }
        else
        {
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            col.Width = width;
        }
        return col;
    }

    private void WireEvents()
    {
        _browseZip.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "Zip files (*.zip)|*.zip|All files (*.*)|*.*", Title = "Select the drawings zip" };
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                _zipBox.Text = dlg.FileName;
                _outBox.Text = DefaultOutputFor(dlg.FileName);
            }
        };

        _browseOut.Click += (_, _) =>
        {
            using var dlg = new SaveFileDialog { Filter = "Excel workbook (*.xlsx)|*.xlsx", Title = "Save QC report as", FileName = Path.GetFileName(_outBox.Text) };
            if (!string.IsNullOrWhiteSpace(_outBox.Text))
                dlg.InitialDirectory = Path.GetDirectoryName(_outBox.Text);
            if (dlg.ShowDialog(this) == DialogResult.OK)
                _outBox.Text = dlg.FileName;
        };

        _openExcelBtn.Click += (_, _) =>
        {
            if (File.Exists(_lastReportPath))
                Process.Start(new ProcessStartInfo(_lastReportPath) { UseShellExecute = true });
        };

        _runBtn.Click += async (_, _) => await RunAsync();
    }

    private static string DefaultOutputFor(string zipPath) =>
        Path.Combine(Path.GetDirectoryName(zipPath) ?? ".",
            Path.GetFileNameWithoutExtension(zipPath) + " - QC Report.xlsx");

    private async Task RunAsync()
    {
        string zip = _zipBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(zip) || !File.Exists(zip))
        {
            MessageBox.Show(this, "Please select a valid zip file.", "Drawing QC", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (string.IsNullOrWhiteSpace(_outBox.Text))
            _outBox.Text = DefaultOutputFor(zip);
        string outPath = _outBox.Text.Trim();

        SetBusy(true);
        _grid.Rows.Clear();
        _countsLabel.Text = "";
        _progress.Value = 0;

        var progress = new Progress<(int done, int total, string name)>(p =>
        {
            _progress.Maximum = Math.Max(p.total, 1);
            _progress.Value = Math.Min(p.done, _progress.Maximum);
            _statusLabel.Text = $"Reading {p.done}/{p.total}:  {p.name}";
        });

        try
        {
            var results = await Task.Run(() => QcEngine.Analyze(zip, progress));

            // Always show results in the grid first, so nothing is lost even if the
            // Excel file happens to be open/locked.
            PopulateGrid(results);
            int matched = results.Count(r => r.FinalStatus == "Matched");
            int unmatched = results.Count(r => r.FinalStatus == "Unmatched");
            int duplicate = results.Count(r => r.FinalStatus == "Duplicate");
            _countsLabel.Text = $"Total: {results.Count}     Matched: {matched}     Unmatched: {unmatched}     Duplicate: {duplicate}";

            _statusLabel.Text = "Writing Excel report...";
            string saved = await Task.Run(() => WriteReportResilient(results, outPath));
            _lastReportPath = saved;
            _openExcelBtn.Enabled = true;

            _statusLabel.Text = string.Equals(saved, outPath, StringComparison.OrdinalIgnoreCase)
                ? $"Done. Report saved to: {saved}"
                : $"The chosen file was open in Excel, so the report was saved to a new copy: {saved}";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "Failed.";
            MessageBox.Show(this, "Error: " + ex.Message,
                "Drawing QC", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    /// <summary>
    /// Saves the report; if the target file is open/locked in Excel, falls back to a
    /// timestamped copy in the same folder and returns the path actually written.
    /// </summary>
    private static string WriteReportResilient(List<PdfResult> results, string outPath)
    {
        try
        {
            QcEngine.WriteReport(results, outPath);
            return outPath;
        }
        catch (IOException)
        {
            string dir = Path.GetDirectoryName(outPath) ?? ".";
            string name = Path.GetFileNameWithoutExtension(outPath);
            string alt = Path.Combine(dir, $"{name} {DateTime.Now:yyyy-MM-dd HHmmss}.xlsx");
            QcEngine.WriteReport(results, alt);
            return alt;
        }
    }

    private void PopulateGrid(List<PdfResult> results)
    {
        var green = Color.FromArgb(0xC6, 0xEF, 0xCE);
        var greenText = Color.FromArgb(0x00, 0x61, 0x00);
        var red = Color.FromArgb(0xFF, 0xC7, 0xCE);
        var redText = Color.FromArgb(0x9C, 0x00, 0x06);
        var amber = Color.FromArgb(0xFF, 0xEB, 0x9C);
        var amberText = Color.FromArgb(0x9C, 0x63, 0x00);

        _grid.SuspendLayout();
        int sr = 1;
        foreach (var r in results)
        {
            int idx = _grid.Rows.Add(
                sr++,
                r.FileName,
                string.IsNullOrWhiteSpace(r.Drawing1) ? "-" : r.Drawing1,
                string.IsNullOrWhiteSpace(r.Drawing2) ? "-" : r.Drawing2,
                r.FinalStatus,
                string.IsNullOrWhiteSpace(r.NotFoundInPdf) ? "-" : r.NotFoundInPdf);

            Color back, fore;
            switch (r.FinalStatus)
            {
                case "Matched": back = green; fore = greenText; break;
                case "Duplicate": back = amber; fore = amberText; break;
                default: back = red; fore = redText; break;
            }
            var style = _grid.Rows[idx].DefaultCellStyle;
            style.BackColor = back;
            style.ForeColor = fore;
            style.SelectionBackColor = ControlPaint.Dark(back);
            style.SelectionForeColor = fore;
        }
        _grid.ResumeLayout();
    }

    private void SetBusy(bool busy)
    {
        _runBtn.Enabled = !busy;
        _browseZip.Enabled = !busy;
        _browseOut.Enabled = !busy;
        _runBtn.Text = busy ? "Working..." : "Run QC";
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }
}
