using System.Drawing.Drawing2D;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
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
    public bool Selected = true;
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
    readonly Button clearBtn = new() { Text = "Clear", AutoSize = true };
    readonly Button matcherBtn = new() { Text = "Card matcher...", AutoSize = true };
    readonly Button selAllBtn = new() { Text = "Select all", AutoSize = true };
    readonly Button selNoneBtn = new() { Text = "Select none", AutoSize = true };
    bool refreshing;
    MatcherForm matcher;
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
        clearBtn.Click += (_, _) => ClearList();
        matcherBtn.Click += (_, _) => OpenMatcher();
        onlyChangesChk.CheckedChanged += (_, _) => RefreshList();
        list.CheckBoxes = true;
        list.ItemChecked += ListItemChecked;
        selAllBtn.Click += (_, _) => SetAllChecked(true);
        selNoneBtn.Click += (_, _) => SetAllChecked(false);

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
        opts.Controls.Add(selAllBtn);
        opts.Controls.Add(selNoneBtn);
        table.Controls.Add(opts, 0, 3);
        table.SetColumnSpan(opts, 3);

        foreach (var chk in new[] { authorChk, nameChk, versionChk, guidChk, gameChk })
            chk.CheckedChanged += (_, _) => RebuildPattern();
        patternBox.TextChanged += (_, _) => Recompute();

        var btns = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        btns.Controls.Add(scanBtn);
        btns.Controls.Add(renameBtn);
        btns.Controls.Add(undoBtn);
        btns.Controls.Add(clearBtn);
        btns.Controls.Add(matcherBtn);
        btns.Controls.Add(statusLbl);
        table.Controls.Add(btns, 0, 4);
        table.SetColumnSpan(btns, 3);

        // Fill control first, then the Top-docked panel
        Controls.Add(list);
        Controls.Add(table);

        // Drag & drop a folder (or a file inside it) to scan it
        AllowDrop = true;
        EnableDrop(this);

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
            if (matcher != null && !matcher.IsDisposed) matcher.SetDark(darkToggle.Checked);
        };

        suppress = false;
        RebuildPattern();
    }

    void EnableDrop(Control c)
    {
        c.AllowDrop = true;
        c.DragEnter += OnDragEnterFolder;
        c.DragDrop += OnDropFolder;
        foreach (Control child in c.Controls) EnableDrop(child);
    }

    void OnDragEnterFolder(object sender, DragEventArgs e)
    {
        e.Effect = scanBtn.Enabled && e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    async void OnDropFolder(object sender, DragEventArgs e)
    {
        if (!scanBtn.Enabled) return;
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] paths || paths.Length == 0) return;

        // A single folder: scan the whole folder
        if (paths.Length == 1 && Directory.Exists(paths[0]))
        {
            folderBox.Text = paths[0];
            await ScanAsync();
            return;
        }

        // Otherwise load only exactly what was dropped
        var opts = new EnumerationOptions { RecurseSubdirectories = recursiveChk.Checked, IgnoreInaccessible = true };
        var files = new List<string>();
        foreach (var p in paths)
        {
            if (Directory.Exists(p))
                files.AddRange(Directory.EnumerateFiles(p, "*.zipmod", opts)
                    .Where(f => f.EndsWith(".zipmod", StringComparison.OrdinalIgnoreCase)));
            else if (p.EndsWith(".zipmod", StringComparison.OrdinalIgnoreCase) && File.Exists(p))
                files.Add(p);
        }
        files = files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        if (files.Count == 0)
        {
            statusLbl.Text = "No .zipmod files in what was dropped";
            return;
        }

        folderBox.Text = CommonFolder(files);
        await ScanAsync(files);
    }

    // The deepest folder that contains all the given files
    static string CommonFolder(List<string> files)
    {
        char sep = System.IO.Path.DirectorySeparatorChar;
        var parts = files.Select(f => System.IO.Path.GetDirectoryName(f).Split(sep)).ToList();
        var first = parts[0];
        int n = first.Length;
        foreach (var p in parts)
        {
            int i = 0;
            while (i < n && i < p.Length && string.Equals(first[i], p[i], StringComparison.OrdinalIgnoreCase)) i++;
            n = i;
        }
        string common = string.Join(sep.ToString(), first.Take(n));
        if (common.Length == 0) return System.IO.Path.GetDirectoryName(files[0]);
        if (common.Length == 2 && common[1] == ':') common += sep;   // "D:" -> "D:\"
        return common;
    }

    // Empty the list so you can start fresh (the Folder box is left alone)
    void ClearList()
    {
        if (!scanBtn.Enabled) return;   // ignore while a scan is running
        items = new List<Item>();
        list.Items.Clear();
        statusLbl.Text = "";
        renameBtn.Enabled = false;
    }

    void OpenMatcher()
    {
        if (matcher == null || matcher.IsDisposed)
        {
            matcher = new MatcherForm(folderBox.Text.Trim(), darkToggle.Checked);
            matcher.Show(this);
        }
        else
        {
            matcher.Activate();
        }
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
        const TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.Left |
                                      TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;

        var bgColor = e.Item.Selected ? Color.FromArgb(62, 76, 102) : list.BackColor;
        Rectangle cell = e.Bounds;
        bool firstWithCheck = e.ColumnIndex == 0 && list.CheckBoxes;

        if (firstWithCheck)
        {
            // Column 0 includes the checkbox area, so work out the whole cell ourselves
            var row = e.Item.GetBounds(ItemBoundsPortion.Entire);
            cell = new Rectangle(row.Left, row.Top, list.Columns[0].Width, row.Height);
        }

        using (var bg = new SolidBrush(bgColor))
            e.Graphics.FillRectangle(bg, cell);

        if (firstWithCheck)
        {
            int size = 13;
            var box = new Rectangle(cell.Left + 4, cell.Top + (cell.Height - size) / 2, size, size);
            using (var fill = new SolidBrush(Color.FromArgb(32, 32, 32)))
                e.Graphics.FillRectangle(fill, box);
            using (var border = new Pen(Color.FromArgb(150, 150, 150)))
                e.Graphics.DrawRectangle(border, box);

            if (e.Item.Checked)
            {
                var old = e.Graphics.SmoothingMode;
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var tick = new Pen(Color.FromArgb(235, 235, 235), 1.8f))
                    e.Graphics.DrawLines(tick, new[]
                    {
                        new Point(box.Left + 3, box.Top + 7),
                        new Point(box.Left + 5, box.Top + 10),
                        new Point(box.Left + 10, box.Top + 3)
                    });
                e.Graphics.SmoothingMode = old;
            }

            cell = new Rectangle(cell.Left + 22, cell.Top, cell.Width - 22, cell.Height);
        }

        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, list.Font, Rectangle.Inflate(cell, -4, 0),
            PageFore, flags);
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
        int selected = items.Count(i => i.Selected && i.NewPath != i.Path);
        statusLbl.Text = $"{items.Count} zipmods, {changes} to rename ({selected} selected)";
        renameBtn.Enabled = selected > 0;
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

    internal static Dictionary<string, string> ReadManifest(string path)
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

    async Task ScanAsync(List<string> only = null)
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
            List<string> files = only != null
                ? only.OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToList()
                : Directory.EnumerateFiles(root, "*.zipmod", opts)
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
        refreshing = true;
        list.BeginUpdate();
        list.Items.Clear();
        foreach (var it in items)
        {
            bool changed = it.NewPath != it.Path;
            if (onlyChangesChk.Checked && !changed && !it.Status.StartsWith("Skipped")) continue;
            var row = new ListViewItem(System.IO.Path.GetRelativePath(root, it.Path));
            row.SubItems.Add(changed ? System.IO.Path.GetRelativePath(root, it.NewPath) : "");
            row.SubItems.Add(it.Status);
            row.Tag = it;
            row.Checked = changed && it.Selected;
            list.Items.Add(row);
        }
        list.EndUpdate();
        refreshing = false;
        FitLastColumn();
    }

    // ------------------------------------------------------------- rename --

    void ListItemChecked(object sender, ItemCheckedEventArgs e)
    {
        if (refreshing) return;
        if (e.Item.Tag is Item it)
        {
            it.Selected = e.Item.Checked;
            UpdateSummary();
        }
    }

    // Select all / none applies to the rows currently shown that have a pending rename
    void SetAllChecked(bool value)
    {
        refreshing = true;
        foreach (ListViewItem row in list.Items)
        {
            if (row.Tag is Item it && it.NewPath != it.Path)
            {
                row.Checked = value;
                it.Selected = value;
            }
        }
        refreshing = false;
        UpdateSummary();
    }

    // Unchecked files stay where they are, so make sure no selected file's new name
    // collides with a file that isn't moving.
    static void ResolveConflicts(List<Item> todo)
    {
        var moving = new HashSet<string>(todo.Select(t => t.Path), StringComparer.OrdinalIgnoreCase);
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var folder in todo.Select(t => System.IO.Path.GetDirectoryName(t.Path)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                foreach (var f in Directory.EnumerateFileSystemEntries(folder))
                    if (!moving.Contains(f)) taken.Add(f);
            }
            catch { }
        }

        foreach (var it in todo)
        {
            string dir = System.IO.Path.GetDirectoryName(it.NewPath);
            string stem = System.IO.Path.GetFileNameWithoutExtension(it.NewPath);
            string candidate = it.NewPath;
            int n = 2;
            while (taken.Contains(candidate))
                candidate = System.IO.Path.Combine(dir, $"{stem} ({n++}).zipmod");
            taken.Add(candidate);
            it.NewPath = candidate;
        }
    }

    static string TempName(string path) =>
        System.IO.Path.Combine(System.IO.Path.GetDirectoryName(path), $"__tmp_{Guid.NewGuid():N}.zipmod");

    void DoRename()
    {
        var todo = items.Where(i => i.Selected && i.NewPath != i.Path).ToList();
        if (todo.Count == 0) return;
        ResolveConflicts(todo);

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


// ===================================================================
//  Card matcher
// ===================================================================

class ZipInfo
{
    public string Path, Name, Version, Author;
}

class ModRef
{
    public string Guid;
    public int Count;
    public SortedSet<string> Props = new();
}

class CardResult
{
    public string Path;
    public string Error;
    public string Note;
    public List<ModRef> Mods = new();
}

// Minimal MessagePack reader (enough to read Koikatsu card headers and extended data)
static class MsgPack
{
    public static object Read(byte[] d)
    {
        int p = 0;
        return Read(d, ref p, 0);
    }

    static ulong BE(byte[] d, ref int p, int n)
    {
        if (p + n > d.Length) throw new Exception("msgpack: unexpected end of data");
        ulong v = 0;
        for (int i = 0; i < n; i++) v = (v << 8) | d[p + i];
        p += n;
        return v;
    }

    static string Str(byte[] d, ref int p, int n)
    {
        if (n < 0 || p + n > d.Length) throw new Exception("msgpack: bad string length");
        string s = Encoding.UTF8.GetString(d, p, n);
        p += n;
        return s;
    }

    static byte[] Bin(byte[] d, ref int p, int n)
    {
        if (n < 0 || p + n > d.Length) throw new Exception("msgpack: bad binary length");
        var b = new byte[n];
        Array.Copy(d, p, b, 0, n);
        p += n;
        return b;
    }

    static List<object> Arr(byte[] d, ref int p, int n, int depth)
    {
        var l = new List<object>();
        for (int i = 0; i < n; i++) l.Add(Read(d, ref p, depth + 1));
        return l;
    }

    static Dictionary<string, object> Map(byte[] d, ref int p, int n, int depth)
    {
        var m = new Dictionary<string, object>();
        for (int i = 0; i < n; i++)
        {
            object k = Read(d, ref p, depth + 1);
            object v = Read(d, ref p, depth + 1);
            m[k?.ToString() ?? "null"] = v;
        }
        return m;
    }

    static object Ext(byte[] d, ref int p, int n)
    {
        p += 1; // type byte
        return Bin(d, ref p, n);
    }

    static object Read(byte[] d, ref int p, int depth)
    {
        if (depth > 64) throw new Exception("msgpack: nesting too deep");
        if (p >= d.Length) throw new Exception("msgpack: unexpected end of data");
        byte b = d[p++];
        int n;

        if (b <= 0x7f) return (long)b;
        if (b >= 0xe0) return (long)(sbyte)b;
        if ((b & 0xe0) == 0xa0) return Str(d, ref p, b & 0x1f);
        if ((b & 0xf0) == 0x90) return Arr(d, ref p, b & 0x0f, depth);
        if ((b & 0xf0) == 0x80) return Map(d, ref p, b & 0x0f, depth);

        switch (b)
        {
            case 0xc0: return null;
            case 0xc2: return false;
            case 0xc3: return true;

            case 0xc4: n = (int)BE(d, ref p, 1); return Bin(d, ref p, n);
            case 0xc5: n = (int)BE(d, ref p, 2); return Bin(d, ref p, n);
            case 0xc6: n = (int)BE(d, ref p, 4); return Bin(d, ref p, n);

            case 0xc7: n = (int)BE(d, ref p, 1); return Ext(d, ref p, n);
            case 0xc8: n = (int)BE(d, ref p, 2); return Ext(d, ref p, n);
            case 0xc9: n = (int)BE(d, ref p, 4); return Ext(d, ref p, n);

            case 0xca: return (double)BitConverter.Int32BitsToSingle((int)BE(d, ref p, 4));
            case 0xcb: return BitConverter.Int64BitsToDouble((long)BE(d, ref p, 8));

            case 0xcc: return (long)BE(d, ref p, 1);
            case 0xcd: return (long)BE(d, ref p, 2);
            case 0xce: return (long)BE(d, ref p, 4);
            case 0xcf: return (long)BE(d, ref p, 8);
            case 0xd0: return (long)(sbyte)BE(d, ref p, 1);
            case 0xd1: return (long)(short)BE(d, ref p, 2);
            case 0xd2: return (long)(int)BE(d, ref p, 4);
            case 0xd3: return (long)BE(d, ref p, 8);

            case 0xd4: return Ext(d, ref p, 1);
            case 0xd5: return Ext(d, ref p, 2);
            case 0xd6: return Ext(d, ref p, 4);
            case 0xd7: return Ext(d, ref p, 8);
            case 0xd8: return Ext(d, ref p, 16);

            case 0xd9: n = (int)BE(d, ref p, 1); return Str(d, ref p, n);
            case 0xda: n = (int)BE(d, ref p, 2); return Str(d, ref p, n);
            case 0xdb: n = (int)BE(d, ref p, 4); return Str(d, ref p, n);

            case 0xdc: n = (int)BE(d, ref p, 2); return Arr(d, ref p, n, depth);
            case 0xdd: n = (int)BE(d, ref p, 4); return Arr(d, ref p, n, depth);
            case 0xde: n = (int)BE(d, ref p, 2); return Map(d, ref p, n, depth);
            case 0xdf: n = (int)BE(d, ref p, 4); return Map(d, ref p, n, depth);
        }
        throw new Exception("msgpack: unknown type 0x" + b.ToString("x2"));
    }
}

// Reads the zipmod references (Sideloader resolve info) out of a Koikatsu character card
static class CardReader
{
    const string UarId = "com.bepis.sideloader.universalautoresolver";

    static long ToLong(object o) => o switch
    {
        long l => l,
        double d => (long)d,
        _ => 0
    };

    // Returns the offset just after the PNG's IEND chunk, or -1
    static int FindPngEnd(byte[] d)
    {
        byte[] sig = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        if (d.Length < 8) return -1;
        for (int i = 0; i < 8; i++) if (d[i] != sig[i]) return -1;

        long p = 8;
        while (p + 8 <= d.Length)
        {
            long len = ((long)d[p] << 24) | ((long)d[p + 1] << 16) | ((long)d[p + 2] << 8) | d[p + 3];
            string type = Encoding.ASCII.GetString(d, (int)p + 4, 4);
            p += 12 + len;
            if (type == "IEND") return p <= d.Length ? (int)p : d.Length;
            if (p > d.Length) return -1;
        }
        return -1;
    }

    public static CardResult Read(string path)
    {
        var r = new CardResult { Path = path };
        try
        {
            byte[] data = File.ReadAllBytes(path);
            int pos = FindPngEnd(data);
            if (pos < 0) throw new Exception("not a valid PNG file");
            if (pos >= data.Length) throw new Exception("no card data inside this PNG (not a Koikatsu card)");

            using var ms = new MemoryStream(data, pos, data.Length - pos);
            using var br = new BinaryReader(ms, Encoding.UTF8);

            br.ReadInt32();                      // product number
            string marker = br.ReadString();     // e.g. 【KoiKatuChara】
            if (!marker.StartsWith("【KoiKatu"))
                throw new Exception("not a Koikatsu card");
            if (marker.Contains("Clothes"))
                throw new Exception("outfit/coordinate cards aren't supported yet");

            br.ReadString();                     // card version
            int faceLen = br.ReadInt32();
            if (faceLen > 0) ms.Seek(faceLen, SeekOrigin.Current);

            int hdrLen = br.ReadInt32();
            byte[] hdr = br.ReadBytes(hdrLen);
            br.ReadInt64();                      // total data length
            long dataStart = ms.Position;

            // Block header: list of {name, version, pos, size}
            object root = MsgPack.Read(hdr);
            List<object> infos = null;
            if (root is Dictionary<string, object> rd && rd.TryGetValue("lstInfo", out var li)) infos = li as List<object>;
            else if (root is List<object> rl && rl.Count > 0) infos = rl[0] as List<object> ?? rl;
            if (infos == null) throw new Exception("couldn't read the card's block list");

            long extPos = -1, extSize = 0;
            foreach (var inf in infos)
            {
                string name = null; long bp = 0, bs = 0;
                if (inf is Dictionary<string, object> m)
                {
                    name = m.GetValueOrDefault("name") as string;
                    bp = ToLong(m.GetValueOrDefault("pos"));
                    bs = ToLong(m.GetValueOrDefault("size"));
                }
                else if (inf is List<object> a && a.Count >= 4)
                {
                    name = a[0] as string; bp = ToLong(a[2]); bs = ToLong(a[3]);
                }
                if (name == "KKEx") { extPos = bp; extSize = bs; break; }
            }

            if (extPos < 0)
            {
                r.Note = "No extended data (no modded items)";
                return r;
            }

            ms.Position = dataStart + extPos;
            byte[] ext = br.ReadBytes((int)extSize);
            var extRoot = MsgPack.Read(ext) as Dictionary<string, object>;
            if (extRoot == null || !extRoot.TryGetValue(UarId, out var plugin))
            {
                r.Note = "No zipmod references found";
                return r;
            }

            // PluginData is either [version, {data}] or {"version":..,"data":{..}}
            Dictionary<string, object> pdata = null;
            if (plugin is List<object> pl && pl.Count >= 2) pdata = pl[1] as Dictionary<string, object>;
            else if (plugin is Dictionary<string, object> pm) pdata = pm.GetValueOrDefault("data") as Dictionary<string, object> ?? pm;

            if (pdata == null || !pdata.TryGetValue("info", out var infoObj) || infoObj is not List<object> entries)
            {
                r.Note = "No zipmod references found";
                return r;
            }

            var map = new Dictionary<string, ModRef>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in entries)
            {
                object parsed = entry is byte[] bytes ? MsgPack.Read(bytes) : entry;
                if (parsed is not Dictionary<string, object> ri) continue;

                string guid = (ri.GetValueOrDefault("ModID") ?? ri.GetValueOrDefault("GUID")) as string;
                if (string.IsNullOrWhiteSpace(guid)) continue;

                if (!map.TryGetValue(guid, out var mr))
                    map[guid] = mr = new ModRef { Guid = guid };
                mr.Count++;

                string prop = ri.GetValueOrDefault("Property") as string;
                if (string.IsNullOrWhiteSpace(prop)) prop = ri.GetValueOrDefault("Slot") as string;
                if (!string.IsNullOrWhiteSpace(prop)) mr.Props.Add(prop);
            }

            r.Mods = map.Values.OrderBy(m => m.Guid, StringComparer.OrdinalIgnoreCase).ToList();
            if (r.Mods.Count == 0) r.Note = "No zipmod references found";
        }
        catch (Exception ex)
        {
            r.Error = ex.Message;
        }
        return r;
    }
}

// Dark theme helper for secondary windows
static class DarkTheme
{
    [DllImport("dwmapi.dll")]
    static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
    static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);

    public static Color Back(bool dark) => dark ? Color.FromArgb(32, 32, 32) : SystemColors.Control;
    public static Color Fore(bool dark) => dark ? Color.FromArgb(230, 230, 230) : SystemColors.ControlText;
    public static Color Field(bool dark) => dark ? Color.FromArgb(45, 45, 45) : SystemColors.Window;

    public static void Apply(Form form, bool dark)
    {
        form.BackColor = Back(dark);
        form.ForeColor = Fore(dark);
        Walk(form, dark);
        ApplyNative(form, dark);
    }

    static void Walk(Control parent, bool dark)
    {
        foreach (Control c in parent.Controls)
        {
            switch (c)
            {
                case ListView lv:
                    lv.BackColor = Field(dark);
                    lv.ForeColor = Fore(dark);
                    lv.GridLines = !dark;
                    lv.OwnerDraw = dark;
                    if (lv.IsHandleCreated) SetWindowTheme(lv.Handle, dark ? "DarkMode_Explorer" : "Explorer", null);
                    lv.Invalidate();
                    break;
                case Button b:
                    if (dark)
                    {
                        b.FlatStyle = FlatStyle.Flat;
                        b.BackColor = Color.FromArgb(60, 60, 60);
                        b.ForeColor = Fore(true);
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
                    t.BackColor = Field(dark);
                    t.ForeColor = Fore(dark);
                    t.BorderStyle = dark ? BorderStyle.FixedSingle : BorderStyle.Fixed3D;
                    break;
                default:
                    c.BackColor = Back(dark);
                    c.ForeColor = Fore(dark);
                    break;
            }
            Walk(c, dark);
        }
    }

    public static void ApplyNative(Form form, bool dark)
    {
        if (!form.IsHandleCreated) return;
        int v = dark ? 1 : 0;
        DwmSetWindowAttribute(form.Handle, 20, ref v, sizeof(int));
        DwmSetWindowAttribute(form.Handle, 19, ref v, sizeof(int));
    }

    // Hook up custom drawing (only used while the list's OwnerDraw is on, i.e. dark mode)
    public static void Attach(ListView lv)
    {
        const TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.Left |
                                      TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;

        lv.DrawColumnHeader += (s, e) =>
        {
            using var bg = new SolidBrush(Color.FromArgb(52, 52, 52));
            using var pen = new Pen(Color.FromArgb(85, 85, 85));
            e.Graphics.FillRectangle(bg, e.Bounds);
            e.Graphics.DrawLine(pen, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
            e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            TextRenderer.DrawText(e.Graphics, e.Header.Text, lv.Font, Rectangle.Inflate(e.Bounds, -6, 0),
                Fore(true), flags);
        };

        lv.DrawSubItem += (s, e) =>
        {
            var bgColor = e.Item.Selected ? Color.FromArgb(62, 76, 102) : lv.BackColor;
            using var bg = new SolidBrush(bgColor);
            e.Graphics.FillRectangle(bg, e.Bounds);
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, lv.Font, Rectangle.Inflate(e.Bounds, -4, 0),
                e.Item.ForeColor, flags);
        };

        lv.HandleCreated += (s, e) =>
            SetWindowTheme(lv.Handle, lv.OwnerDraw ? "DarkMode_Explorer" : "Explorer", null);
    }
}

class MatcherForm : Form
{
    readonly TextBox folderBox = new() { Dock = DockStyle.Fill };
    readonly Button browseBtn = new() { Text = "Browse...", AutoSize = true };
    readonly Button addBtn = new() { Text = "Add cards...", AutoSize = true };
    readonly Button addFolderBtn = new() { Text = "Add folder of cards...", AutoSize = true };
    readonly Button clearBtn = new() { Text = "Clear", AutoSize = true };
    readonly Button reindexBtn = new() { Text = "Re-scan mods folder", AutoSize = true };
    readonly Button copyGuidsBtn = new() { Text = "Copy missing GUIDs", AutoSize = true };
    readonly Button copyFilesBtn = new() { Text = "Copy needed zipmods to...", AutoSize = true };
    readonly CheckBox onlyMissingChk = new() { Text = "Only show missing", AutoSize = true };
    readonly Label statusLbl = new() { AutoSize = true, Padding = new Padding(8, 6, 0, 0) };
    readonly ListView list = new()
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        GridLines = true
    };

    bool dark;
    bool busy;
    Dictionary<string, List<ZipInfo>> index;
    string indexedFolder;
    int indexCount;
    readonly List<CardResult> cards = new();

    public MatcherForm(string modsFolder, bool darkMode)
    {
        dark = darkMode;
        Text = "Card matcher";
        Width = 1150;
        Height = 650;
        MinimumSize = new Size(850, 450);
        folderBox.Text = modsFolder;

        list.Columns.Add("Card", 220);
        list.Columns.Add("Status", 80);
        list.Columns.Add("Mod name", 230);
        list.Columns.Add("GUID", 270);
        list.Columns.Add("Used for", 170);
        list.Columns.Add("File", 300);

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 3,
            Padding = new Padding(8)
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "Mods folder:", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) }, 0, 0);
        table.Controls.Add(folderBox, 1, 0);
        table.Controls.Add(browseBtn, 2, 0);

        var row1 = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        row1.Controls.Add(addBtn);
        row1.Controls.Add(addFolderBtn);
        row1.Controls.Add(clearBtn);
        row1.Controls.Add(reindexBtn);
        table.Controls.Add(row1, 0, 1);
        table.SetColumnSpan(row1, 3);

        var row2 = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        row2.Controls.Add(onlyMissingChk);
        row2.Controls.Add(copyGuidsBtn);
        row2.Controls.Add(copyFilesBtn);
        row2.Controls.Add(statusLbl);
        table.Controls.Add(row2, 0, 2);
        table.SetColumnSpan(row2, 3);

        var hint = new Label
        {
            Text = "Drag and drop character cards (.png) or folders of cards anywhere on this window.",
            AutoSize = true,
            Padding = new Padding(0, 4, 0, 0)
        };
        table.Controls.Add(hint, 0, 3);
        table.SetColumnSpan(hint, 3);

        Controls.Add(list);
        Controls.Add(table);

        browseBtn.Click += async (_, _) => await BrowseAsync();
        addBtn.Click += async (_, _) => await AddFilesAsync();
        addFolderBtn.Click += async (_, _) => await AddFolderAsync();
        clearBtn.Click += (_, _) => { cards.Clear(); Render(); };
        reindexBtn.Click += async (_, _) => { await EnsureIndexAsync(true); Render(); };
        copyGuidsBtn.Click += (_, _) => CopyMissingGuids();
        copyFilesBtn.Click += (_, _) => CopyNeededFiles();
        onlyMissingChk.CheckedChanged += (_, _) => Render();
        list.SizeChanged += (_, _) => FitLastColumn();

        DarkTheme.Attach(list);

        AllowDrop = true;
        EnableDrop(this);

        DarkTheme.Apply(this, dark);
        Render();
    }

    public void SetDark(bool value)
    {
        dark = value;
        DarkTheme.Apply(this, dark);
        Render();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        DarkTheme.ApplyNative(this, dark);
    }

    // ---------------------------------------------------------- drag & drop --

    void EnableDrop(Control c)
    {
        if (c is TextBox) return;
        c.AllowDrop = true;
        c.DragEnter += OnDragEnter;
        c.DragDrop += OnDragDrop;
        foreach (Control child in c.Controls) EnableDrop(child);
    }

    void OnDragEnter(object sender, DragEventArgs e)
    {
        e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
    }

    async void OnDragDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths) await AddPathsAsync(paths);
    }

    // -------------------------------------------------------------- actions --

    async Task BrowseAsync()
    {
        using var dlg = new FolderBrowserDialog { Description = "Select your mods folder" };
        if (dlg.ShowDialog() != DialogResult.OK) return;
        folderBox.Text = dlg.SelectedPath;
        await EnsureIndexAsync(true);
        Render();
    }

    async Task AddFilesAsync()
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Select character cards",
            Filter = "Koikatsu cards (*.png)|*.png",
            Multiselect = true
        };
        if (dlg.ShowDialog() == DialogResult.OK) await AddPathsAsync(dlg.FileNames);
    }

    async Task AddFolderAsync()
    {
        using var dlg = new FolderBrowserDialog { Description = "Select a folder of character cards" };
        if (dlg.ShowDialog() == DialogResult.OK) await AddPathsAsync(new[] { dlg.SelectedPath });
    }

    void SetBusy(bool value)
    {
        busy = value;
        foreach (var b in new[] { browseBtn, addBtn, addFolderBtn, clearBtn, reindexBtn, copyGuidsBtn, copyFilesBtn })
            b.Enabled = !value;
    }

    async Task EnsureIndexAsync(bool force)
    {
        string folder = folderBox.Text.Trim();
        if (!Directory.Exists(folder))
        {
            index = null;
            indexedFolder = null;
            return;
        }
        if (!force && index != null && string.Equals(indexedFolder, folder, StringComparison.OrdinalIgnoreCase))
            return;

        SetBusy(true);
        var progress = new Progress<string>(t => statusLbl.Text = t);

        var result = await Task.Run(() =>
        {
            var idx = new Dictionary<string, List<ZipInfo>>(StringComparer.OrdinalIgnoreCase);
            var files = Directory.EnumerateFiles(folder, "*.zipmod",
                    new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true })
                .Where(f => f.EndsWith(".zipmod", StringComparison.OrdinalIgnoreCase))
                .ToList();

            for (int i = 0; i < files.Count; i++)
            {
                if (i % 25 == 0) ((IProgress<string>)progress).Report($"Indexing mods {i + 1}/{files.Count}...");
                try
                {
                    var m = MainForm.ReadManifest(files[i]);
                    string guid = m["guid"];
                    if (guid.Length == 0) continue;
                    if (!idx.TryGetValue(guid, out var l)) idx[guid] = l = new List<ZipInfo>();
                    l.Add(new ZipInfo { Path = files[i], Name = m["name"], Version = m["version"], Author = m["author"] });
                }
                catch { }
            }
            return (Index: idx, Count: files.Count);
        });

        index = result.Index;
        indexCount = result.Count;
        indexedFolder = folder;
        SetBusy(false);
    }

    async Task AddPathsAsync(IEnumerable<string> paths)
    {
        if (busy) return;

        string folder = folderBox.Text.Trim();
        if (!Directory.Exists(folder))
        {
            MessageBox.Show("Choose your mods folder first so the cards can be matched against it.", Text);
            return;
        }

        var files = new List<string>();
        foreach (var p in paths)
        {
            if (Directory.Exists(p))
                files.AddRange(Directory.EnumerateFiles(p, "*.png",
                    new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true }));
            else if (p.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                files.Add(p);
        }

        files = files
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(f => !cards.Any(c => string.Equals(c.Path, f, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (files.Count == 0)
        {
            statusLbl.Text = "No new .png cards found";
            return;
        }

        await EnsureIndexAsync(false);

        SetBusy(true);
        statusLbl.Text = $"Reading {files.Count} card(s)...";
        var read = await Task.Run(() => files.Select(CardReader.Read).ToList());
        cards.AddRange(read);
        SetBusy(false);
        Render();
    }

    // -------------------------------------------------------------- results --

    static ZipInfo Best(List<ZipInfo> zips) =>
        zips.OrderByDescending(z => Version.TryParse(z.Version, out var v) ? v : new Version(0, 0)).First();

    static string Join(IEnumerable<string> items, int max)
    {
        var l = items.ToList();
        string s = string.Join("; ", l.Take(max));
        return l.Count > max ? s + $" (+{l.Count - max} more)" : s;
    }

    void Render()
    {
        bool onlyMissing = onlyMissingChk.Checked;
        var red = dark ? Color.FromArgb(255, 130, 130) : Color.Firebrick;
        var orange = dark ? Color.FromArgb(240, 180, 90) : Color.DarkOrange;

        var allGuids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var missingGuids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        list.BeginUpdate();
        list.Items.Clear();

        void AddRow(string card, string status, string mod, string guid, string used, string file, Color? color)
        {
            var row = new ListViewItem(card);
            row.SubItems.Add(status);
            row.SubItems.Add(mod);
            row.SubItems.Add(guid);
            row.SubItems.Add(used);
            row.SubItems.Add(file);
            if (color.HasValue) row.ForeColor = color.Value;
            list.Items.Add(row);
        }

        foreach (var card in cards)
        {
            string cardName = System.IO.Path.GetFileName(card.Path);

            if (card.Error != null)
            {
                AddRow(cardName, "Error", card.Error, "", "", "", orange);
                continue;
            }
            if (card.Mods.Count == 0)
            {
                if (!onlyMissing) AddRow(cardName, "No mods", card.Note ?? "No zipmod references", "", "", "", null);
                continue;
            }

            foreach (var mod in card.Mods)
            {
                allGuids.Add(mod.Guid);
                List<ZipInfo> zips = null;
                bool known = index != null;
                bool found = known && index.TryGetValue(mod.Guid, out zips);
                if (known && !found) missingGuids.Add(mod.Guid);

                if (onlyMissing && (found || !known)) continue;

                string used = Join(mod.Props, 3);
                if (found)
                {
                    var best = Best(zips);
                    AddRow(cardName, "Found", best.Name, mod.Guid, used,
                        Join(zips.Select(z => System.IO.Path.GetFileName(z.Path)), 3), null);
                }
                else
                {
                    AddRow(cardName, known ? "MISSING" : "Unknown", "", mod.Guid, used, "", known ? red : orange);
                }
            }
        }

        list.EndUpdate();
        FitLastColumn();

        int foundCount = allGuids.Count - missingGuids.Count;
        statusLbl.Text = cards.Count == 0
            ? (index != null ? $"{indexCount} zipmods indexed" : "")
            : $"{cards.Count} card(s), {allGuids.Count} mod(s) needed, {foundCount} found, {missingGuids.Count} missing  ({indexCount} zipmods indexed)";
    }

    void FitLastColumn()
    {
        int last = list.Columns.Count - 1;
        if (last < 1) return;
        int used = 0;
        for (int i = 0; i < last; i++) used += list.Columns[i].Width;
        list.Columns[last].Width = Math.Max(150, list.ClientSize.Width - used);
    }

    IEnumerable<string> AllNeededGuids() =>
        cards.Where(c => c.Error == null)
             .SelectMany(c => c.Mods)
             .Select(m => m.Guid)
             .Distinct(StringComparer.OrdinalIgnoreCase);

    void CopyMissingGuids()
    {
        if (index == null)
        {
            MessageBox.Show("Choose a valid mods folder first.", Text);
            return;
        }
        var missing = AllNeededGuids().Where(g => !index.ContainsKey(g)).ToList();
        if (missing.Count == 0)
        {
            MessageBox.Show("Nothing is missing.", Text);
            return;
        }
        Clipboard.SetText(string.Join(Environment.NewLine, missing));
        statusLbl.Text = $"Copied {missing.Count} missing GUID(s) to the clipboard";
    }

    void CopyNeededFiles()
    {
        if (index == null)
        {
            MessageBox.Show("Choose a valid mods folder first.", Text);
            return;
        }
        var wanted = AllNeededGuids()
            .Where(g => index.ContainsKey(g))
            .Select(g => Best(index[g]).Path)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (wanted.Count == 0)
        {
            MessageBox.Show("None of the needed zipmods were found in your mods folder.", Text);
            return;
        }

        using var dlg = new FolderBrowserDialog { Description = "Choose where to copy the needed zipmods" };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        int copied = 0, skipped = 0;
        var errors = new List<string>();
        foreach (var src in wanted)
        {
            string dest = System.IO.Path.Combine(dlg.SelectedPath, System.IO.Path.GetFileName(src));
            if (File.Exists(dest)) { skipped++; continue; }
            try { File.Copy(src, dest); copied++; }
            catch (Exception ex) { errors.Add($"{System.IO.Path.GetFileName(src)}: {ex.Message}"); }
        }

        string msg = $"Copied {copied} zipmod(s)." + (skipped > 0 ? $" Skipped {skipped} that already existed." : "");
        if (errors.Count > 0) msg += $"\n\n{errors.Count} problems:\n" + string.Join("\n", errors.Take(10));
        MessageBox.Show(msg, Text);
    }
}
