using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PrintCalc
{
    public sealed class PrintItem
    {
        public string Path { get; init; } = "";
        public string Name { get; init; } = "";
        public string Type { get; init; } = "";
        public int Pages { get; init; }
    }

    public class Form1 : Form
    {
        private readonly ListView fileList = new();
        private readonly TextBox searchBox = new();
        private readonly ComboBox mode = new();
        private readonly NumericUpDown copies = new();
        private readonly NumericUpDown simplexRate = new();
        private readonly NumericUpDown duplexRate = new();
        private readonly Label totalLabel = new();
        private readonly Label selectionLabel = new();
        private readonly Label statusLabel = new();
        private readonly Label pathLabel = new();
        private readonly Label versionLabel = new();
        private FileSystemWatcher? watcher;
        private readonly Dictionary<string, PrintItem> items = new(StringComparer.OrdinalIgnoreCase);
        private readonly string settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PrintDesk", "settings.ini");

        private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif"
        };

        private static readonly string AppVersion = "1.2.0";

        public Form1()
        {
            Text = "PrintDesk — Shop Print Manager";
            Width = 1180;
            Height = 760;
            MinimumSize = new Size(980, 650);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(244, 247, 251);
            Font = new Font("Segoe UI", 10F);

            LoadSettings();
            BuildUi();
            StartWhatsAppWatcher();
            CatchStartupErrors();
        }

        private void CatchStartupErrors()
        {
            // Keep startup resilient: folders may not exist yet and WhatsApp may be closed.
            try { RefreshWhatsAppFiles(); } catch (Exception ex) { LogError(ex); }
        }

        private void BuildUi()
        {
            Controls.Clear();

            var header = new Panel { Dock = DockStyle.Top, Height = 88, BackColor = Color.FromArgb(22, 31, 46) };
            var title = new Label
            {
                Text = "PrintDesk",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 24F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(22, 12)
            };
            header.Controls.Add(title);

            var subtitle = new Label
            {
                Text = "WhatsApp PDFs & photos → fast shop billing",
                ForeColor = Color.FromArgb(191, 204, 222),
                Font = new Font("Segoe UI", 10F),
                AutoSize = true,
                Location = new Point(25, 52)
            };
            header.Controls.Add(subtitle);

            statusLabel.Text = "● Starting…";
            statusLabel.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold);
            statusLabel.AutoSize = true;
            statusLabel.ForeColor = Color.FromArgb(132, 231, 166);
            statusLabel.Location = new Point(920, 28);
            header.Controls.Add(statusLabel);

            versionLabel.Text = "v" + AppVersion;
            versionLabel.ForeColor = Color.FromArgb(144, 155, 173);
            versionLabel.AutoSize = true;
            versionLabel.Location = new Point(920, 51);
            header.Controls.Add(versionLabel);
            Controls.Add(header);

            var toolbar = new Panel { Dock = DockStyle.Top, Height = 132, BackColor = Color.White, Padding = new Padding(18, 12, 18, 10) };

            AddText(toolbar, "MODE", 18, 8);
            mode.Items.AddRange(new object[] { "Simplex", "Duplex" });
            mode.SelectedIndex = 0;
            mode.DropDownStyle = ComboBoxStyle.DropDownList;
            mode.Location = new Point(18, 30); mode.Width = 110;
            toolbar.Controls.Add(mode);

            AddText(toolbar, "COPIES", 145, 8);
            copies.Minimum = 1; copies.Maximum = 9999; copies.Value = 1;
            copies.Location = new Point(145, 30); copies.Width = 70;
            toolbar.Controls.Add(copies);

            AddText(toolbar, "SIMPLEX / PAGE", 230, 8);
            simplexRate.Minimum = 0; simplexRate.Maximum = 100000; simplexRate.Value = 10;
            simplexRate.Location = new Point(230, 30); simplexRate.Width = 90;
            toolbar.Controls.Add(simplexRate);

            AddText(toolbar, "DUPLEX / SHEET", 335, 8);
            duplexRate.Minimum = 0; duplexRate.Maximum = 100000; duplexRate.Value = 15;
            duplexRate.Location = new Point(335, 30); duplexRate.Width = 90;
            toolbar.Controls.Add(duplexRate);

            var refresh = ButtonStyle("↻ Refresh", 440, 28, 110, Color.FromArgb(234, 239, 247), Color.FromArgb(36, 47, 64));
            refresh.Click += (_, _) => RefreshWhatsAppFiles();
            toolbar.Controls.Add(refresh);

            var add = ButtonStyle("＋ Add files", 560, 28, 112, Color.FromArgb(47, 111, 255), Color.White);
            add.Click += (_, _) => AddFilesDialog();
            toolbar.Controls.Add(add);

            var open = ButtonStyle("Open", 682, 28, 82, Color.FromArgb(234, 239, 247), Color.FromArgb(36, 47, 64));
            open.Click += (_, _) => OpenSelected();
            toolbar.Controls.Add(open);

            var print = ButtonStyle("Print", 774, 28, 82, Color.FromArgb(26, 169, 111), Color.White);
            print.Click += (_, _) => PrintSelected();
            toolbar.Controls.Add(print);

            var remove = ButtonStyle("Remove", 866, 28, 88, Color.FromArgb(253, 235, 236), Color.FromArgb(183, 55, 67));
            remove.Click += (_, _) => RemoveSelected();
            toolbar.Controls.Add(remove);

            AddText(toolbar, "SEARCH", 18, 72);
            searchBox.PlaceholderText = "Search filename…";
            searchBox.Location = new Point(75, 69); searchBox.Width = 325;
            toolbar.Controls.Add(searchBox);

            pathLabel.Text = "WhatsApp: detecting transfer folder…";
            pathLabel.ForeColor = Color.FromArgb(112, 123, 140);
            pathLabel.Font = new Font("Segoe UI", 8.5F);
            pathLabel.AutoEllipsis = true;
            pathLabel.Location = new Point(415, 73); pathLabel.Width = 540;
            toolbar.Controls.Add(pathLabel);
            Controls.Add(toolbar);

            var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(18, 14, 18, 110), BackColor = Color.FromArgb(244, 247, 251) };
            var card = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(10, 10, 10, 8) };

            fileList.Dock = DockStyle.Fill;
            fileList.View = View.Details;
            fileList.FullRowSelect = true;
            fileList.CheckBoxes = true;
            fileList.MultiSelect = true;
            fileList.GridLines = false;
            fileList.BorderStyle = BorderStyle.None;
            fileList.BackColor = Color.White;
            fileList.Font = new Font("Segoe UI", 10.5F);
            fileList.Columns.Add("TYPE", 86);
            fileList.Columns.Add("FILE", 520);
            fileList.Columns.Add("PAGES", 90);
            fileList.Columns.Add("PRICE / COPY", 130);
            fileList.DoubleClick += (_, _) => OpenSelected();
            fileList.ItemChecked += (_, _) => Recalculate();
            card.Controls.Add(fileList);
            body.Controls.Add(card);
            Controls.Add(body);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 102, BackColor = Color.FromArgb(22, 31, 46), Padding = new Padding(20, 10, 20, 10) };
            selectionLabel.Text = "0 selected";
            selectionLabel.ForeColor = Color.FromArgb(174, 188, 208);
            selectionLabel.Location = new Point(22, 17);
            selectionLabel.AutoSize = true;
            bottom.Controls.Add(selectionLabel);

            totalLabel.Text = "TOTAL  Rs. 0.00";
            totalLabel.ForeColor = Color.White;
            totalLabel.Font = new Font("Segoe UI Semibold", 25F, FontStyle.Bold);
            totalLabel.Location = new Point(20, 35);
            totalLabel.AutoSize = true;
            bottom.Controls.Add(totalLabel);

            var rule = new Label
            {
                Text = "Duplex: odd PDF page count +1 blank side → 2 pages = Rs.15",
                ForeColor = Color.FromArgb(182, 194, 211),
                AutoSize = true,
                Location = new Point(710, 48)
            };
            bottom.Controls.Add(rule);
            Controls.Add(bottom);

            mode.SelectedIndexChanged += (_, _) => { SaveSettings(); Recalculate(); };
            copies.ValueChanged += (_, _) => { SaveSettings(); Recalculate(); };
            simplexRate.ValueChanged += (_, _) => { SaveSettings(); Recalculate(); };
            duplexRate.ValueChanged += (_, _) => { SaveSettings(); Recalculate(); };
            searchBox.TextChanged += (_, _) => RebuildList();
        }

        private static void AddText(Control parent, string text, int x, int y)
        {
            parent.Controls.Add(new Label
            {
                Text = text, AutoSize = true,
                ForeColor = Color.FromArgb(111, 122, 139),
                Font = new Font("Segoe UI Semibold", 8F, FontStyle.Bold),
                Location = new Point(x, y)
            });
        }

        private static Button ButtonStyle(string text, int x, int y, int width, Color back, Color fore)
        {
            return new Button
            {
                Text = text, Location = new Point(x, y), Width = width, Height = 36,
                FlatStyle = FlatStyle.Flat, BackColor = back, ForeColor = fore,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold), Cursor = Cursors.Hand,
                FlatAppearance = { BorderSize = 0 }
            };
        }

        private void StartWhatsAppWatcher()
        {
            try
            {
                string? transfers = FindWhatsAppTransfers();
                if (transfers == null)
                {
                    statusLabel.Text = "● WhatsApp folder not found";
                    statusLabel.ForeColor = Color.FromArgb(255, 175, 95);
                    pathLabel.Text = "WhatsApp: open WhatsApp Desktop and click Refresh";
                    return;
                }

                watcher?.Dispose();
                watcher = new FileSystemWatcher(transfers, "*.*")
                {
                    IncludeSubdirectories = true,
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                    EnableRaisingEvents = true
                };
                watcher.Created += (_, e) => QueueAdd(e.FullPath);
                watcher.Renamed += (_, e) => QueueAdd(e.FullPath);

                statusLabel.Text = "● WhatsApp monitoring ON";
                statusLabel.ForeColor = Color.FromArgb(125, 234, 169);
                pathLabel.Text = "WhatsApp: " + transfers;
                pathLabel.AutoEllipsis = true;
            }
            catch (Exception ex)
            {
                statusLabel.Text = "● Monitoring unavailable";
                statusLabel.ForeColor = Color.FromArgb(255, 155, 155);
                LogError(ex);
            }
        }

        private static string? FindWhatsAppTransfers()
        {
            string sessions = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Packages", "5319275A.WhatsAppDesktop_cv1g1gvanyjgm", "LocalState", "sessions");

            try
            {
                if (!Directory.Exists(sessions)) return null;
                var dirs = Directory.EnumerateDirectories(sessions, "*", SearchOption.AllDirectories)
                    .Where(d => string.Equals(Path.GetFileName(d), "transfers", StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(d => SafeCreation(d))
                    .ToList();
                return dirs.FirstOrDefault();
            }
            catch { return null; }
        }

        private static DateTime SafeCreation(string path)
        {
            try { return Directory.GetCreationTimeUtc(path); } catch { return DateTime.MinValue; }
        }

        private void QueueAdd(string path)
        {
            if (IsDisposed || !IsSupported(path)) return;
            try { BeginInvoke(new Action(() => AddFile(path))); } catch { }
        }

        private static bool IsSupported(string path)
        {
            string ext = Path.GetExtension(path);
            return ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) || ImageExtensions.Contains(ext);
        }

        private void RefreshWhatsAppFiles()
        {
            try
            {
                StartWhatsAppWatcher();
                string? folder = FindWhatsAppTransfers();
                if (folder == null) return;
                foreach (var path in Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories))
                    if (IsSupported(path)) AddFile(path);
                statusLabel.Text = "● WhatsApp monitoring ON";
                statusLabel.ForeColor = Color.FromArgb(125, 234, 169);
            }
            catch (Exception ex)
            {
                LogError(ex);
                statusLabel.Text = "● Refresh failed";
                statusLabel.ForeColor = Color.FromArgb(255, 155, 155);
            }
        }

        private void AddFilesDialog()
        {
            using var dialog = new OpenFileDialog
            {
                Title = "Add PDF or photo files",
                Multiselect = true,
                Filter = "Print files|*.pdf;*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.gif|PDF|*.pdf|Images|*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.gif"
            };
            if (dialog.ShowDialog() != DialogResult.OK) return;
            foreach (var path in dialog.FileNames) AddFile(path);
        }

        private void AddFile(string path)
        {
            if (!File.Exists(path) || !IsSupported(path) || items.ContainsKey(path)) return;
            try
            {
                string ext = Path.GetExtension(path);
                int pages = ImageExtensions.Contains(ext) ? 1 : CountPdfPages(path);
                items[path] = new PrintItem { Path = path, Name = Path.GetFileName(path), Type = ImageExtensions.Contains(ext) ? "PHOTO" : "PDF", Pages = pages };
                RebuildList(selectPath: path);
            }
            catch (Exception ex) { LogError(ex); }
        }

        private void RebuildList(string? selectPath = null)
        {
            var checkedPaths = fileList.Items.Cast<ListViewItem>()
                .Where(x => x.Checked && x.Tag is string)
                .Select(x => (string)x.Tag!).ToHashSet(StringComparer.OrdinalIgnoreCase);

            fileList.BeginUpdate();
            fileList.Items.Clear();
            string query = searchBox.Text.Trim();
            foreach (var item in items.Values.OrderByDescending(x => SafeCreation(x.Path)))
            {
                if (query.Length > 0 && !item.Name.Contains(query, StringComparison.OrdinalIgnoreCase)) continue;
                decimal price = PriceFor(item.Pages);
                var row = new ListViewItem(item.Type);
                row.SubItems.Add(item.Name);
                row.SubItems.Add(item.Pages.ToString());
                row.SubItems.Add($"Rs. {price:N2}");
                row.Tag = item.Path;
                row.Checked = checkedPaths.Contains(item.Path) || string.Equals(selectPath, item.Path, StringComparison.OrdinalIgnoreCase);
                fileList.Items.Add(row);
            }
            fileList.EndUpdate();
            Recalculate();
        }

        private decimal PriceFor(int pages)
        {
            if (mode.SelectedIndex == 1)
            {
                int adjusted = pages % 2 == 0 ? pages : pages + 1;
                return (adjusted / 2m) * duplexRate.Value;
            }
            return pages * simplexRate.Value;
        }

        private void Recalculate()
        {
            decimal total = 0;
            int selected = 0;
            foreach (ListViewItem row in fileList.Items)
            {
                if (!row.Checked || row.Tag is not string path || !items.TryGetValue(path, out var item)) continue;
                selected++;
                total += PriceFor(item.Pages) * copies.Value;
            }
            selectionLabel.Text = $"{selected} selected • {items.Count} files loaded";
            totalLabel.Text = $"TOTAL  Rs. {total:N2}";
        }

        private IEnumerable<string> SelectedPaths() => fileList.Items.Cast<ListViewItem>()
            .Where(x => x.Checked && x.Tag is string)
            .Select(x => (string)x.Tag!);

        private void OpenSelected()
        {
            string? path = SelectedPaths().FirstOrDefault();
            if (path == null) return;
            try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
            catch (Exception ex) { LogError(ex); MessageBox.Show(this, "Could not open the selected file.", "PrintDesk", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void PrintSelected()
        {
            foreach (string path in SelectedPaths())
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        Verb = "print",
                        UseShellExecute = true,
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    });
                }
                catch (Exception ex) { LogError(ex); }
            }
        }

        private void RemoveSelected()
        {
            foreach (string path in SelectedPaths().ToList()) items.Remove(path);
            RebuildList();
        }

        private int CountPdfPages(string path)
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream);
            string text = reader.ReadToEnd();
            return Math.Max(1, Regex.Matches(text, @"/Type\s*/Page\b").Count);
        }

        private void SaveSettings()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
                var sb = new StringBuilder();
                sb.AppendLine($"Mode={(mode.SelectedIndex == 1 ? "Duplex" : "Simplex")}");
                sb.AppendLine($"Copies={copies.Value}");
                sb.AppendLine($"Simplex={simplexRate.Value}");
                sb.AppendLine($"Duplex={duplexRate.Value}");
                File.WriteAllText(settingsPath, sb.ToString());
            }
            catch { }
        }

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(settingsPath)) return;
                foreach (string line in File.ReadAllLines(settingsPath))
                {
                    var parts = line.Split('=', 2);
                    if (parts.Length != 2) continue;
                    if (parts[0] == "Mode") mode.SelectedIndex = parts[1] == "Duplex" ? 1 : 0;
                    if (parts[0] == "Copies" && decimal.TryParse(parts[1], out var c)) copies.Value = Math.Min(copies.Maximum, Math.Max(copies.Minimum, c));
                    if (parts[0] == "Simplex" && decimal.TryParse(parts[1], out var s)) simplexRate.Value = Math.Min(simplexRate.Maximum, Math.Max(simplexRate.Minimum, s));
                    if (parts[0] == "Duplex" && decimal.TryParse(parts[1], out var d)) duplexRate.Value = Math.Min(duplexRate.Maximum, Math.Max(duplexRate.Minimum, d));
                }
            }
            catch { }
        }

        private void LogError(Exception ex)
        {
            try
            {
                string log = Path.Combine(Path.GetTempPath(), "PrintDesk-error.log");
                File.AppendAllText(log, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex}\r\n");
            }
            catch { }
        }

        private static DateTime SafeCreation(string path)
        {
            try { return File.GetCreationTimeUtc(path); } catch { return DateTime.MinValue; }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            SaveSettings();
            watcher?.Dispose();
            base.OnFormClosed(e);
        }

        [STAThread]
        public static void Main()
        {
            try
            {
                ApplicationConfiguration.Initialize();
                Application.ThreadException += (_, e) =>
                {
                    try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "PrintDesk-error.log"), $"{DateTime.Now}: {e.Exception}\r\n"); } catch { }
                    MessageBox.Show("PrintDesk encountered an error. Details were saved to %TEMP%\\PrintDesk-error.log.", "PrintDesk", MessageBoxButtons.OK, MessageBoxIcon.Error);
                };
                AppDomain.CurrentDomain.UnhandledException += (_, e) =>
                {
                    try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "PrintDesk-error.log"), $"{DateTime.Now}: {e.ExceptionObject}\r\n"); } catch { }
                };
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                try { File.WriteAllText(Path.Combine(Path.GetTempPath(), "PrintDesk-error.log"), ex.ToString()); } catch { }
                MessageBox.Show("PrintDesk could not start. See %TEMP%\\PrintDesk-error.log", "PrintDesk", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}