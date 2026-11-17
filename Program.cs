using System.Drawing.Drawing2D;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ZipmodRenamer;

static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

class Item
{
    public string Path;
    public string NewPath;
    public string Status = "";
    public Dictionary<string, string> Info;
}

class LogEntry
{
    public string Old { get; set; }
    public string New { get; set; }
}

class LogFile
{
    public string Time { get; set; }
    public List<LogEntry> Renames { get; set; } = new();
}

class MainForm : Form
{
    const string DefaultPattern = "[{author}] {name} v{version}";
    const string LogName = "zipmod_rename_log.json";
    const int MaxStem = 150;

    static readonly Regex InvalidChars = new(@"[<>:""/\\|?*\x00-\x1f]");

    readonly TextBox folderBox = new() { Dock = DockStyle.Fill };
    readonly TextBox patternBox = new() { Dock = DockStyle.Fill };
    readonly CheckBox authorChk = new() { Text = "Author", Checked = true, AutoSize = true };
    readonly CheckBox nameChk = new() { Text = "Name", Checked = true, AutoSize = true };
    readonly CheckBox versionChk = new() { Text = "Version", Checked = true, AutoSize = true };
    readonly CheckBox guidChk = new() { Text = "GUID", Checked = false, AutoSize = true };
    readonly CheckBox gameChk = new() { Text = "Game", Checked = false, AutoSize = true };
    readonly ThemeToggle darkToggle = new();
    readonly ToolTip tip = new();
    bool darkMode;
    bool suppress = true;

    [DllImport("dwmapi.dll")]
    static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
    static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);
    readonly CheckBox recursiveChk = new() { Text = "Include subfolders", Checked = true, AutoSize = true };
    readonly CheckBox onlyChangesChk = new() { Text = "Only show files that will change", Checked = true, AutoSize = true };
    readonly Button scanBtn = new() { Text = "1. Scan / Preview", AutoSize = true };
    readonly Button renameBtn = new() { Text = "2. Rename", AutoSize = true, Enabled = false };
    readonly Button undoBtn = new() { Text = "Undo last rename", AutoSize = true };
    readonly Label statusLbl = new() { AutoSize = true, Padding = new Padding(12, 6, 0, 0) };
    readonly ListView list = new()
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        GridLines = true
    };

    List<Item> items = new();

    public MainForm()
    {
        Text = "Zipmod Renamer";
        Width = 1050;
        Height = 650;
        MinimumSize = new Size(800, 450);

        list.Columns.Add("Current filename", 350);
        list.Columns.Add("New filename", 420);
        list.Columns.Add("Status", 200);

        var browseBtn = new Button { Text = "Browse...", AutoSize = true };
        browseBtn.Click += (_, _) => Browse();
        scanBtn.Click += async (_, _) => await ScanAsync();
        renameBtn.Click += (_, _) => DoRename();
        undoBtn.Click += async (_, _) => await UndoAsync();
        onlyChangesChk.CheckedChanged += (_, _) => RefreshList();

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 4,
            Padding = new Padding(8)
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "Folder:", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) }, 0, 0);
        table.Controls.Add(folderBox, 1, 0);
        table.Controls.Add(browseBtn, 2, 0);

        darkToggle.Anchor = AnchorStyles.Right;
        darkToggle.Margin = new Padding(14, 0, 0, 0);
        tip.SetToolTip(darkToggle, "Toggle dark mode");
        table.Controls.Add(darkToggle, 3, 0);

        table.Controls.Add(new Label { Text = "Name pattern:", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) }, 0, 1);
        table.Controls.Add(patternBox, 1, 1);
        table.SetColumnSpan(patternBox, 2);

        var parts = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        parts.Controls.Add(new Label { Text = "Include in name:", AutoSize = true, Padding = new Padding(0, 3, 8, 0) });
        parts.Controls.Add(authorChk);
        parts.Controls.Add(nameChk);
        parts.Controls.Add(versionChk);
        parts.Controls.Add(guidChk);
        parts.Controls.Add(gameChk);
        table.Controls.Add(parts, 0, 2);
        table.SetColumnSpan(parts, 3);

        var opts = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        opts.Controls.Add(recursiveChk);
        opts.Controls.Add(onlyChangesChk);
        table.Controls.Add(opts, 0, 3);
        table.SetColumnSpan(opts, 3);

        foreach (var chk in new[] { authorChk, nameChk, versionChk, guidChk, gameChk })
            chk.CheckedChanged += (_, _) => RebuildPattern();
        patternBox.TextChanged += (_, _) => Recompute();

        var btns = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        btns.Controls.Add(scanBtn);
        btns.Controls.Add(renameBtn);
        btns.Controls.Add(undoBtn);
        btns.Controls.Add(statusLbl);
        table.Controls.Add(btns, 0, 4);
        table.SetColumnSpan(btns, 3);

        // Fill control first, then the Top-docked panel
        Controls.Add(list);
        Controls.Add(table);

        // Dark mode setup
        typeof(ListView).GetProperty("DoubleBuffered",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.SetValue(list, true);
        list.DrawColumnHeader += ListDrawHeader;
        list.DrawSubItem += ListDrawSubItem;
        list.SizeChanged += (_, _) => FitLastColumn();
        list.HandleCreated += (_, _) => ApplyNative();

        darkToggle.SetCheckedSilently(LoadDarkPref());
        ApplyTheme(darkToggle.Checked);
        darkToggle.CheckedChanged += (_, _) =>
        {
            ApplyTheme(darkToggle.Checked);
            SaveDarkPref(darkToggle.Checked);
        };

        suppress = false;
        RebuildPattern();
    }

    // ------------------------------------------------------------- theming --

    static string PrefPath => System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ZipmodRenamer", "theme.txt");

    static bool LoadDarkPref()
    {
        try { if (File.Exists(PrefPath)) return File.ReadAllText(PrefPath).Trim() == "dark"; } catch { }
        try
        {
            using var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            if (k?.GetValue("AppsUseLightTheme") is int v) return v == 0;
        }
        catch { }
        return false;
    }

    static void SaveDarkPref(bool dark)
    {
        try
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(PrefPath));
            File.WriteAllText(PrefPath, dark ? "dark" : "light");
        }
        catch { }
    }

    Color FieldBack => darkMode ? Color.FromArgb(45, 45, 45) : SystemColors.Window;
    Color PageFore => darkMode ? Color.FromArgb(230, 230, 230) : SystemColors.ControlText;

    void ApplyTheme(bool dark)
    {
        darkMode = dark;
        var back = dark ? Color.FromArgb(32, 32, 32) : SystemColors.Control;

        BackColor = back;
        ForeColor = PageFore;
        ThemeControls(this, back);

        list.BackColor = FieldBack;
        list.ForeColor = PageFore;
        list.GridLines = !dark;
        list.OwnerDraw = dark;   // custom drawing only in dark mode; light mode stays fully native
        list.Invalidate();

        ApplyNative();
    }

    void ThemeControls(Control parent, Color back)
    {
        foreach (Control c in parent.Controls)
        {
            switch (c)
            {
                case ListView:
                    break;
                case Button b:
                    if (darkMode)
                    {
                        b.FlatStyle = FlatStyle.Flat;
                        b.BackColor = Color.FromArgb(60, 60, 60);
                        b.ForeColor = PageFore;
                        b.FlatAppearance.BorderColor = Color.FromArgb(95, 95, 95);
                    }
                    else
                    {
                        b.FlatStyle = FlatStyle.Standard;
                        b.UseVisualStyleBackColor = true;
                        b.ForeColor = SystemColors.ControlText;
                    }
                    break;
                case TextBox t:
                    t.BackColor = FieldBack;
                    t.ForeColor = PageFore;
                    t.BorderStyle = darkMode ? BorderStyle.FixedSingle : BorderStyle.Fixed3D;
                    break;
                default:
                    c.BackColor = back;
                    c.ForeColor = PageFore;
                    break;
            }
            ThemeControls(c, back);
        }
    }

    // Title bar + scrollbars (needs window handles)
    void ApplyNative()
    {
        if (IsHandleCreated)
        {
            int v = darkMode ? 1 : 0;
            DwmSetWindowAttribute(Handle, 20, ref v, sizeof(int));   // Win10 2004+ / Win11
            DwmSetWindowAttribute(Handle, 19, ref v, sizeof(int));   // older Win10 builds
        }
        if (list.IsHandleCreated)
            SetWindowTheme(list.Handle, darkMode ? "DarkMode_Explorer" : "Explorer", null);
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ApplyNative();
    }

    void ListDrawHeader(object sender, DrawListViewColumnHeaderEventArgs e)
    {
        using var bg = new SolidBrush(Color.FromArgb(52, 52, 52));
        using var pen = new Pen(Color.FromArgb(85, 85, 85));
        e.Graphics.FillRectangle(bg, e.Bounds);
        e.Graphics.DrawLine(pen, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
        e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        TextRenderer.DrawText(e.Graphics, e.Header.Text, list.Font, Rectangle.Inflate(e.Bounds, -6, 0),
            PageFore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
    }

    void ListDrawSubItem(object sender, DrawListViewSubItemEventArgs e)
    {
        var bgColor = e.Item.Selected ? Color.FromArgb(62, 76, 102) : list.BackColor;
        using var bg = new SolidBrush(bgColor);
        e.Graphics.FillRectangle(bg, e.Bounds);
        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, list.Font, Rectangle.Inflate(e.Bounds, -4, 0),
            PageFore, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
    }

    // Make the last column fill the remaining width (no blank column on the right)
    void FitLastColumn()
    {
        if (list.Columns.Count < 3) return;
        int used = list.Columns[0].Width + list.Columns[1].Width;
        list.Columns[2].Width = Math.Max(120, list.ClientSize.Width - used);
    }

    // Builds the pattern text from the checkboxes, then refreshes the preview
    void RebuildPattern()
    {
        var p = new List<string>();
        if (authorChk.Checked) p.Add("[{author}]");
        if (nameChk.Checked) p.Add("{name}");
        if (versionChk.Checked) p.Add("v{version}");
        if (guidChk.Checked) p.Add("({guid})");
        if (gameChk.Checked) p.Add("({game})");

        bool old = suppress;
        suppress = true;
        patternBox.Text = string.Join(" ", p);
        suppress = old;
        Recompute();
    }

    // Re-applies the current pattern to already-scanned files (no re-reading of zips)
    void Recompute()
    {
        if (suppress || items.Count == 0 || !scanBtn.Enabled) return;
        foreach (var it in items)
        {
            it.NewPath = it.Path;
            if (it.Info != null) it.Status = "";
        }
        ComputeTargets(items, patternBox.Text.Trim());
        RefreshList();
        UpdateSummary();
    }

    void UpdateSummary()
    {
        int changes = items.Count(i => i.NewPath != i.Path);
        statusLbl.Text = $"{items.Count} zipmods, {changes} to rename";
        renameBtn.Enabled = changes > 0;
    }

    // ------------------------------------------------------------ helpers --

    static string Clean(string text)
    {
        text = InvalidChars.Replace(text ?? "", "");
        text = Regex.Replace(text, @"\s+", " ").Trim().TrimEnd('.', ' ');
        return text;
    }

    static string BuildStem(string pattern, Dictionary<string, string> info)
    {
        string stem = pattern;
        string ver = Clean(info.GetValueOrDefault("version", ""));
        string verWithV = ver.Length == 0 ? "" :
            (Regex.IsMatch(ver, @"^[vV]\d") ? ver : "v" + ver);
        stem = stem.Replace("v{version}", verWithV);
        foreach (var key in new[] { "name", "author", "version", "guid", "game" })
            stem = stem.Replace("{" + key + "}", Clean(info.GetValueOrDefault(key, "")));

        stem = Regex.Replace(stem, @"\[\s*\]", "");
        stem = Regex.Replace(stem, @"\(\s*\)", "");
        stem = Regex.Replace(stem, @"\s+", " ").Trim().TrimEnd('.', ' ');
        if (stem.Length > MaxStem) stem = stem.Substring(0, MaxStem).TrimEnd('.', ' ');
        return stem;
    }

    static Dictionary<string, string> ReadManifest(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        var entry =
            zip.Entries.FirstOrDefault(e => e.FullName.Equals("manifest.xml", StringComparison.OrdinalIgnoreCase))
            ?? zip.Entries.FirstOrDefault(e => e.FullName.EndsWith("/manifest.xml", StringComparison.OrdinalIgnoreCase));
        if (entry == null) throw new Exception("no manifest.xml");

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);

        var fields = new Dictionary<string, string>
        {
            ["guid"] = "", ["name"] = "", ["version"] = "", ["author"] = "", ["game"] = ""
        };
        foreach (var el in doc.Descendants())
        {
            var tag = el.Name.LocalName.ToLowerInvariant();
            if (fields.ContainsKey(tag) && fields[tag] == "")
                fields[tag] = el.Value.Trim();
        }
        return fields;
    }

    void Browse()
    {
        using var dlg = new FolderBrowserDialog { Description = "Select your mods folder" };
        if (dlg.ShowDialog() == DialogResult.OK) folderBox.Text = dlg.SelectedPath;
    }

    // --------------------------------------------------------------- scan --

    async Task ScanAsync()
    {
        string root = folderBox.Text.Trim();
        if (!Directory.Exists(root))
        {
            MessageBox.Show("Please choose a valid folder first.", Text);
            return;
        }

        scanBtn.Enabled = renameBtn.Enabled = undoBtn.Enabled = false;
        bool recursive = recursiveChk.Checked;
        string pattern = patternBox.Text.Trim();
        var progress = new Progress<string>(s => statusLbl.Text = s);

        items = await Task.Run(() =>
        {
            var opts = new EnumerationOptions
            {
                RecurseSubdirectories = recursive,
                IgnoreInaccessible = true
            };
            var files = Directory.EnumerateFiles(root, "*.zipmod", opts)
                .Where(f => f.EndsWith(".zipmod", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var result = new List<Item>();
            for (int i = 0; i < files.Count; i++)
            {
                if (i % 25 == 0) ((IProgress<string>)progress).Report($"Reading {i + 1}/{files.Count}...");
                var item = new Item { Path = files[i], NewPath = files[i] };
                try { item.Info = ReadManifest(files[i]); }
                catch (Exception ex) { item.Status = "Skipped: " + ex.Message; }
                result.Add(item);
            }
            ComputeTargets(result, pattern);
            return result;
        });

        RefreshList();
        scanBtn.Enabled = undoBtn.Enabled = true;
        UpdateSummary();
    }

    static void ComputeTargets(List<Item> all, string pattern)
    {
        var batch = new HashSet<string>(all.Select(i => i.Path), StringComparer.OrdinalIgnoreCase);
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenGuids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Files already in those folders that are not part of this batch
        foreach (var folder in all.Select(i => System.IO.Path.GetDirectoryName(i.Path)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                foreach (var f in Directory.EnumerateFileSystemEntries(folder))
                    if (!batch.Contains(f)) taken.Add(f);
            }
            catch { }
        }

        foreach (var item in all)
        {
            if (item.Info == null)
            {
                taken.Add(item.Path);
                continue;
            }

            string folder = System.IO.Path.GetDirectoryName(item.Path);
            string stem = BuildStem(pattern, item.Info);
            if (stem.Length == 0)
            {
                item.Status = "Skipped: nothing to name it with";
                taken.Add(item.Path);
                continue;
            }

            string dupNote = "";
            string guid = item.Info["guid"];
            if (guid.Length > 0 && !seenGuids.Add(guid)) dupNote = " (duplicate GUID)";

            string candidate = System.IO.Path.Combine(folder, stem + ".zipmod");
            int n = 2;
            while (taken.Contains(candidate))
                candidate = System.IO.Path.Combine(folder, $"{stem} ({n++}).zipmod");

            taken.Add(candidate);
            item.NewPath = candidate;

            if (candidate == item.Path) item.Status = "Already correct" + dupNote;
            else if (string.Equals(candidate, item.Path, StringComparison.OrdinalIgnoreCase)) item.Status = "Case change" + dupNote;
            else item.Status = (n == 2 ? "Will rename" : "Will rename, name in use") + dupNote;
        }
    }

    void RefreshList()
    {
        string root = folderBox.Text.Trim();
        list.BeginUpdate();
        list.Items.Clear();
        foreach (var it in items)
        {
            bool changed = it.NewPath != it.Path;
            if (onlyChangesChk.Checked && !changed && !it.Status.StartsWith("Skipped")) continue;
            var row = new ListViewItem(System.IO.Path.GetRelativePath(root, it.Path));
            row.SubItems.Add(changed ? System.IO.Path.GetRelativePath(root, it.NewPath) : "");
            row.SubItems.Add(it.Status);
            list.Items.Add(row);
        }
        list.EndUpdate();
        FitLastColumn();
    }

    // ------------------------------------------------------------- rename --

    static string TempName(string path) =>
        System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), $"__tmp_{Guid.NewGuid():N}.zipmod");

    void DoRename()
    {
        var todo = items.Where(i => i.NewPath != i.Path).ToList();
        if (todo.Count == 0) return;

        var confirm = MessageBox.Show(
            $"Rename {todo.Count} files?\n\nA log is saved so you can undo this afterwards.\n" +
            "Tip: back up your mods folder before the first run.",
            Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes) return;

        var staged = new List<(Item item, string tmp)>();
        var errors = new List<string>();

        // Phase 1: move everything to temp names so swaps / chains can't collide
        foreach (var it in todo)
        {
            string tmp = TempName(it.Path);
            try { File.Move(it.Path, tmp); staged.Add((it, tmp)); }
            catch (Exception ex) { errors.Add($"{System.IO.Path.GetFileName(it.Path)}: {ex.Message}"); }
        }

        // Phase 2: temp -> final
        var log = new LogFile { Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") };
        foreach (var (it, tmp) in staged)
        {
            try
            {
                File.Move(tmp, it.NewPath);
                log.Renames.Add(new LogEntry { Old = it.Path, New = it.NewPath });
            }
            catch (Exception ex)
            {
                try { File.Move(tmp, it.Path); } catch { }
                errors.Add($"{System.IO.Path.GetFileName(it.Path)}: {ex.Message}");
            }
        }

        if (log.Renames.Count > 0)
        {
            try
            {
                string logPath = System.IO.Path.Combine(folderBox.Text.Trim(), LogName);
                File.WriteAllText(logPath, JsonSerializer.Serialize(log, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }

        string msg = $"Renamed {log.Renames.Count} files.";
        if (errors.Count > 0) msg += $"\n\n{errors.Count} problems:\n" + string.Join("\n", errors.Take(10));
        MessageBox.Show(msg, Text);
        _ = ScanAsync();
    }

    async Task UndoAsync()
    {
        string logPath = System.IO.Path.Combine(folderBox.Text.Trim(), LogName);
        if (!File.Exists(logPath))
        {
            MessageBox.Show("No rename log found in this folder.", Text);
            return;
        }

        var log = JsonSerializer.Deserialize<LogFile>(File.ReadAllText(logPath));
        if (log == null || log.Renames.Count == 0) return;

        var confirm = MessageBox.Show($"Undo {log.Renames.Count} renames from {log.Time}?",
            Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (confirm != DialogResult.Yes) return;

        var staged = new List<(LogEntry entry, string tmp)>();
        var errors = new List<string>();

        foreach (var r in log.Renames)
        {
            string tmp = TempName(r.New);
            try { File.Move(r.New, tmp); staged.Add((r, tmp)); }
            catch (Exception ex) { errors.Add($"{System.IO.Path.GetFileName(r.New)}: {ex.Message}"); }
        }

        int done = 0;
        foreach (var (r, tmp) in staged)
        {
            try { File.Move(tmp, r.Old); done++; }
            catch (Exception ex)
            {
                try { File.Move(tmp, r.New); } catch { }
                errors.Add($"{System.IO.Path.GetFileName(r.Old)}: {ex.Message}");
            }
        }

        if (errors.Count == 0) File.Delete(logPath);

        string msg = $"Restored {done} files.";
        if (errors.Count > 0) msg += $"\n\n{errors.Count} problems:\n" + string.Join("\n", errors.Take(10));
        MessageBox.Show(msg, Text);
        await ScanAsync();
    }
}


// Sun / moon pill toggle with a short slide animation
class ThemeToggle : Control
{
    static readonly Color TrackLight = Color.FromArgb(178, 178, 178);
    static readonly Color TrackDark = Color.FromArgb(51, 51, 51);

    bool isChecked;
    float pos; // 0 = left (light), 1 = right (dark)
    readonly System.Windows.Forms.Timer timer = new() { Interval = 15 };

    public event EventHandler CheckedChanged;

    public ThemeToggle()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
                 ControlStyles.Selectable, true);
        Size = new Size(56, 28);
        Cursor = Cursors.Hand;
        TabStop = true;
        AccessibleName = "Dark mode";
        AccessibleRole = AccessibleRole.CheckButton;

        timer.Tick += (_, _) =>
        {
            float target = isChecked ? 1f : 0f;
            if (pos < target) pos = Math.Min(target, pos + 0.14f);
            else if (pos > target) pos = Math.Max(target, pos - 0.14f);
            if (pos == target) timer.Stop();
            Invalidate();
        };
    }

    public bool Checked
    {
        get => isChecked;
        set
        {
            if (isChecked == value) return;
            isChecked = value;
            timer.Start();
            CheckedChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    // Set the state without animating or raising CheckedChanged (used at startup)
    public void SetCheckedSilently(bool value)
    {
        isChecked = value;
        pos = value ? 1f : 0f;
        Invalidate();
    }

    protected override void OnClick(EventArgs e)
    {
        base.OnClick(e);
        Focus();
        Checked = !Checked;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.KeyCode == Keys.Space || e.KeyCode == Keys.Enter)
        {
            Checked = !Checked;
            e.Handled = true;
        }
    }

    protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
    protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

    protected override void Dispose(bool disposing)
    {
        if (disposing) timer.Dispose();
        base.Dispose(disposing);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(BackColor);

        float w = Width - 1, h = Height - 1;
        Color track = Lerp(TrackLight, TrackDark, pos);

        using (var path = Pill(0, 0, w, h))
        using (var b = new SolidBrush(track))
            g.FillPath(b, path);

        float d = h - 8f;
        float leftX = 4f, rightX = w - 4f - d;
        float cy = h / 2f;

        DrawSun(g, rightX + d / 2f, cy, 1f - pos);
        DrawMoon(g, leftX + d / 2f, cy, pos, track);

        float kx = leftX + (rightX - leftX) * pos;
        using (var b = new SolidBrush(Color.White))
            g.FillEllipse(b, kx, 4f, d, d);

        if (Focused && ShowFocusCues)
        {
            using var pen = new Pen(Color.FromArgb(130, 170, 230), 1.5f);
            using var path = Pill(1, 1, w - 2, h - 2);
            g.DrawPath(pen, path);
        }
    }

    static GraphicsPath Pill(float x, float y, float w, float h)
    {
        var p = new GraphicsPath();
        p.AddArc(x, y, h, h, 90, 180);
        p.AddArc(x + w - h, y, h, h, 270, 180);
        p.CloseFigure();
        return p;
    }

    static Color Lerp(Color a, Color b, float t) => Color.FromArgb(
        (int)(a.R + (b.R - a.R) * t),
        (int)(a.G + (b.G - a.G) * t),
        (int)(a.B + (b.B - a.B) * t));

    static void DrawSun(Graphics g, float cx, float cy, float alpha)
    {
        if (alpha < 0.02f) return;
        var col = Color.FromArgb((int)(255 * Math.Min(1f, alpha)), 255, 255, 255);
        using var brush = new SolidBrush(col);
        using var pen = new Pen(col, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        g.FillEllipse(brush, cx - 3.8f, cy - 3.8f, 7.6f, 7.6f);
        for (int i = 0; i < 8; i++)
        {
            double a = i * Math.PI / 4;
            float cos = (float)Math.Cos(a), sin = (float)Math.Sin(a);
            g.DrawLine(pen, cx + cos * 6.5f, cy + sin * 6.5f, cx + cos * 8.8f, cy + sin * 8.8f);
        }
    }

    static void DrawMoon(Graphics g, float cx, float cy, float alpha, Color track)
    {
        if (alpha < 0.02f) return;
        var col = Color.FromArgb((int)(255 * Math.Min(1f, alpha)), 255, 255, 255);
        using var moon = new SolidBrush(col);
        using var cut = new SolidBrush(track);
        g.FillEllipse(moon, cx - 6.5f, cy - 6.5f, 13f, 13f);
        g.FillEllipse(cut, cx - 6.5f + 4.2f, cy - 6.5f - 3.2f, 12f, 12f);
    }
}
