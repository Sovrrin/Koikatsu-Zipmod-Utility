using System.IO.Compression;
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
    readonly TextBox patternBox = new() { Dock = DockStyle.Fill, Text = DefaultPattern };
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
            ColumnCount = 3,
            Padding = new Padding(8)
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "Folder:", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) }, 0, 0);
        table.Controls.Add(folderBox, 1, 0);
        table.Controls.Add(browseBtn, 2, 0);

        table.Controls.Add(new Label { Text = "Name pattern:", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 6, 0, 0) }, 0, 1);
        table.Controls.Add(patternBox, 1, 1);
        table.Controls.Add(new Label
        {
            Text = "{name} {author} {version} {guid} {game}",
            AutoSize = true,
            ForeColor = Color.Gray,
            Padding = new Padding(0, 6, 0, 0)
        }, 2, 1);

        var opts = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        opts.Controls.Add(recursiveChk);
        opts.Controls.Add(onlyChangesChk);
        table.Controls.Add(opts, 0, 2);
        table.SetColumnSpan(opts, 3);

        var btns = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        btns.Controls.Add(scanBtn);
        btns.Controls.Add(renameBtn);
        btns.Controls.Add(undoBtn);
        btns.Controls.Add(statusLbl);
        table.Controls.Add(btns, 0, 3);
        table.SetColumnSpan(btns, 3);

        // Fill control first, then the Top-docked panel
        Controls.Add(list);
        Controls.Add(table);
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
        foreach (var key in new[] { "name", "author", "version", "guid", "game" })
            stem = stem.Replace("{" + key + "}", Clean(info.GetValueOrDefault(key, "")));

        stem = Regex.Replace(stem, @"\[\s*\]", "");
        stem = Regex.Replace(stem, @"\(\s*\)", "");
        stem = Regex.Replace(stem.Trim(), @"\s+v$", "");
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
        string pattern = string.IsNullOrWhiteSpace(patternBox.Text) ? DefaultPattern : patternBox.Text.Trim();
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
        int changes = items.Count(i => i.NewPath != i.Path);
        statusLbl.Text = $"{items.Count} zipmods, {changes} to rename";
        scanBtn.Enabled = undoBtn.Enabled = true;
        renameBtn.Enabled = changes > 0;
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
