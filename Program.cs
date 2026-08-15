using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PrintCalc
{
    public class Form1 : Form
    {
        private readonly CheckedListBox files = new();
        private readonly Label totalLabel = new();
        private readonly Label itemCountLabel = new();
        private readonly Label statusLabel = new();
        private readonly Label pathLabel = new();
        private readonly NumericUpDown simplexRate = new();
        private readonly NumericUpDown duplexRate = new();
        private readonly NumericUpDown copies = new();
        private readonly ComboBox mode = new();
        private FileSystemWatcher? watcher;
        private readonly Dictionary<string, int> pageCounts = new(StringComparer.OrdinalIgnoreCase);

        private static readonly string WhatsAppTransfersPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Packages",
            "5319275A.WhatsAppDesktop_cv1g1gvanyjgm",
            "LocalState",
            "sessions",
            "A81DBB7DFA7274350BA54C86662D0C28B6F2998D",
            "transfers");

        private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif"
        };

        public Form1()
        {
            Text = "PrintDesk • WhatsApp Print Calculator";
            Width = 1040;
            Height = 760;
            MinimumSize = new Size(900, 680);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(245, 247, 250);
            Font = new Font("Segoe UI", 10F);

            BuildUi();
            StartWhatsAppWatcher();
        }

        private void BuildUi()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 92,
                BackColor = Color.FromArgb(32, 41, 55),
                Padding = new Padding(24, 16, 24, 12)
            };

            var title = new Label
            {
                Text = "PrintDesk",
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 22F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(24, 12)
            };
            header.Controls.Add(title);

            var subtitle = new Label
            {
                Text = "WhatsApp PDF & photo print calculator",
                ForeColor = Color.FromArgb(205, 214, 226),
                Font = new Font("Segoe UI", 10F),
                AutoSize = true,
                Location = new Point(26, 50)
            };
            header.Controls.Add(subtitle);

            statusLabel.Text = "● Connecting to WhatsApp…";
            statusLabel.ForeColor = Color.FromArgb(255, 220, 120);
            statusLabel.AutoSize = true;
            statusLabel.Font = new Font("Segoe UI Semibold", 9F);
            statusLabel.Location = new Point(790, 34);
            header.Controls.Add(statusLabel);
            Controls.Add(header);

            var controlPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 116,
                BackColor = Color.White,
                Padding = new Padding(20, 14, 20, 12)
            };

            AddLabel(controlPanel, "PRINT MODE", 20, 10);
            mode.Items.AddRange(new object[] { "Simplex", "Duplex" });
            mode.SelectedIndex = 0;
            mode.DropDownStyle = ComboBoxStyle.DropDownList;
            mode.Location = new Point(20, 34);
            mode.Width = 120;
            mode.Height = 34;
            controlPanel.Controls.Add(mode);

            AddLabel(controlPanel, "COPIES", 160, 10);
            copies.Minimum = 1;
            copies.Maximum = 9999;
            copies.Value = 1;
            copies.Location = new Point(160, 34);
            copies.Width = 80;
            controlPanel.Controls.Add(copies);

            AddLabel(controlPanel, "SIMPLEX / PAGE", 260, 10);
            simplexRate.Minimum = 0;
            simplexRate.Maximum = 100000;
            simplexRate.Value = 10;
            simplexRate.Location = new Point(260, 34);
            simplexRate.Width = 90;
            controlPanel.Controls.Add(simplexRate);

            AddLabel(controlPanel, "DUPLEX / SHEET", 370, 10);
            duplexRate.Minimum = 0;
            duplexRate.Maximum = 100000;
            duplexRate.Value = 15;
            duplexRate.Location = new Point(370, 34);
            duplexRate.Width = 90;
            controlPanel.Controls.Add(duplexRate);

            var clearButton = MakeButton("Clear selected", 485, 32, 120, 36, Color.FromArgb(235, 239, 245), Color.FromArgb(44, 55, 72));
            clearButton.Click += (_, _) => ClearSelected();
            controlPanel.Controls.Add(clearButton);

            var addButton = MakeButton("Add files", 618, 32, 104, 36, Color.FromArgb(45, 116, 255), Color.White);
            addButton.Click += (_, _) => AddFilesFromDialog();
            controlPanel.Controls.Add(addButton);

            pathLabel.Text = "WhatsApp auto folder: " + WhatsAppTransfersPath;
            pathLabel.ForeColor = Color.FromArgb(107, 116, 130);
            pathLabel.Font = new Font("Segoe UI", 8.5F);
            pathLabel.AutoEllipsis = true;
            pathLabel.Location = new Point(20, 80);
            pathLabel.Width = 950;
            controlPanel.Controls.Add(pathLabel);

            Controls.Add(controlPanel);

            var body = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 18, 20, 14),
                BackColor = Color.FromArgb(245, 247, 250)
            };

            var listCard = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(16, 14, 16, 10)
            };

            var listTitle = new Label
            {
                Text = "Incoming files",
                Font = new Font("Segoe UI Semibold", 13F, FontStyle.Bold),
                ForeColor = Color.FromArgb(36, 44, 58),
                AutoSize = true,
                Location = new Point(16, 12)
            };
            listCard.Controls.Add(listTitle);

            itemCountLabel.Text = "0 files";
            itemCountLabel.ForeColor = Color.FromArgb(120, 128, 140);
            itemCountLabel.AutoSize = true;
            itemCountLabel.Location = new Point(160, 16);
            listCard.Controls.Add(itemCountLabel);

            files.Dock = DockStyle.Fill;
            files.Top = 42;
            files.CheckOnClick = true;
            files.HorizontalScrollbar = true;
            files.BorderStyle = BorderStyle.None;
            files.BackColor = Color.White;
            files.ForeColor = Color.FromArgb(45, 55, 72);
            files.Font = new Font("Segoe UI", 10.5F);
            files.IntegralHeight = false;
            files.ItemCheck += (_, _) => BeginInvoke(new Action(Recalculate));
            listCard.Controls.Add(files);

            body.Controls.Add(listCard);
            Controls.Add(body);

            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 98,
                BackColor = Color.FromArgb(32, 41, 55),
                Padding = new Padding(24, 12, 24, 12)
            };

            totalLabel.Text = "TOTAL  Rs. 0.00";
            totalLabel.ForeColor = Color.White;
            totalLabel.Font = new Font("Segoe UI Semibold", 24F, FontStyle.Bold);
            totalLabel.AutoSize = true;
            totalLabel.Location = new Point(24, 26);
            footer.Controls.Add(totalLabel);

            var rule = new Label
            {
                Text = "Simplex: 1 page = Rs.10   •   Duplex: odd page count gets +1 blank side, then 2 pages = Rs.15",
                ForeColor = Color.FromArgb(190, 200, 214),
                Font = new Font("Segoe UI", 9F),
                AutoSize = true,
                Location = new Point(520, 38)
            };
            footer.Controls.Add(rule);

            Controls.Add(footer);

            mode.SelectedIndexChanged += (_, _) => Recalculate();
            copies.ValueChanged += (_, _) => Recalculate();
            simplexRate.ValueChanged += (_, _) => Recalculate();
            duplexRate.ValueChanged += (_, _) => Recalculate();
        }

        private static void AddLabel(Control parent, string text, int x, int y)
        {
            parent.Controls.Add(new Label
            {
                Text = text,
                ForeColor = Color.FromArgb(115, 123, 136),
                Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(x, y)
            });
        }

        private static Button MakeButton(string text, int x, int y, int width, int height, Color back, Color fore)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Width = width,
                Height = height,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                BackColor = back,
                ForeColor = fore,
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
        }

        private void StartWhatsAppWatcher()
        {
            if (!Directory.Exists(WhatsAppTransfersPath))
            {
                statusLabel.Text = "● WhatsApp folder not found";
                statusLabel.ForeColor = Color.FromArgb(255, 160, 100);
                return;
            }

            watcher?.Dispose();
            watcher = new FileSystemWatcher(WhatsAppTransfersPath)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                IncludeSubdirectories = true,
                Filter = "*.*",
                EnableRaisingEvents = true
            };

            watcher.Created += (_, e) => QueueAddFile(e.FullPath);
            watcher.Renamed += (_, e) => QueueAddFile(e.FullPath);

            statusLabel.Text = "● WhatsApp monitoring ON";
            statusLabel.ForeColor = Color.FromArgb(125, 235, 160);

            // Pick up files that already arrived before the app was opened.
            foreach (var file in SafeEnumerateFiles(WhatsAppTransfersPath))
                QueueAddFile(file);
        }

        private void QueueAddFile(string path)
        {
            if (IsDisposed) return;
            BeginInvoke(new Action(() => AddFile(path)));
        }

        private void AddFilesFromDialog()
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Print files|*.pdf;*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.gif|PDF|*.pdf|Images|*.jpg;*.jpeg;*.png;*.webp;*.bmp;*.gif",
                Multiselect = true,
                Title = "Select PDF or photo files"
            };

            if (dialog.ShowDialog() != DialogResult.OK) return;
            foreach (var file in dialog.FileNames) AddFile(file);
        }

        private void AddFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;

            string ext = Path.GetExtension(path);
            if (!ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) && !ImageExtensions.Contains(ext)) return;

            try
            {
                if (pageCounts.ContainsKey(path)) return;

                int pages = ImageExtensions.Contains(ext) ? 1 : CountPdfPages(path);
                pageCounts[path] = pages;

                string type = ImageExtensions.Contains(ext) ? "PHOTO" : "PDF";
                string text = $"{type}   {Path.GetFileName(path)}   •   {pages} page" + (pages == 1 ? "" : "s");

                files.Items.Insert(0, text);
                files.SetItemChecked(0, true);
                Recalculate();
            }
            catch
            {
                // Ignore files that are still being copied or are not readable yet.
            }
        }

        private void ClearSelected()
        {
            for (int i = files.Items.Count - 1; i >= 0; i--)
            {
                if (!files.GetItemChecked(i)) continue;

                string display = files.Items[i]?.ToString() ?? "";
                var match = pageCounts.Keys.FirstOrDefault(k => display.Contains(Path.GetFileName(k), StringComparison.OrdinalIgnoreCase));
                if (match != null) pageCounts.Remove(match);
                files.Items.RemoveAt(i);
            }

            Recalculate();
        }

        private void Recalculate()
        {
            decimal total = 0;
            int selected = 0;
            int i = 0;

            foreach (var item in files.Items.Cast<object>())
            {
                if (!files.GetItemChecked(i++)) continue;
                selected++;

                string display = item?.ToString() ?? "";
                var key = pageCounts.Keys.FirstOrDefault(k => display.Contains(Path.GetFileName(k), StringComparison.OrdinalIgnoreCase));
                if (key == null) continue;

                int pages = pageCounts[key];
                decimal pricePerCopy;

                if (mode.SelectedIndex == 1)
                {
                    // Duplex: if page count is odd, add one blank side first.
                    int adjustedPages = pages % 2 == 0 ? pages : pages + 1;
                    int sheets = adjustedPages / 2;
                    pricePerCopy = sheets * duplexRate.Value;
                }
                else
                {
                    pricePerCopy = pages * simplexRate.Value;
                }

                total += pricePerCopy * copies.Value;
            }

            itemCountLabel.Text = $"{selected} selected • {files.Items.Count} total";
            totalLabel.Text = $"TOTAL  Rs. {total:N2}";
        }

        private static IEnumerable<string> SafeEnumerateFiles(string root)
        {
            try
            {
                return Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
                    .Where(IsSupportedFile)
                    .OrderByDescending(File.GetCreationTimeUtc)
                    .Take(250)
                    .ToArray();
            }
            catch
            {
                return Array.Empty<string>();
            }
        }

        private static bool IsSupportedFile(string file)
        {
            string ext = Path.GetExtension(file);
            return ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) || ImageExtensions.Contains(ext);
        }

        private static int CountPdfPages(string path)
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return CountPdfPages(stream);
        }

        private static int CountPdfPages(Stream stream)
        {
            using var reader = new StreamReader(stream);
            string text = reader.ReadToEnd();
            int count = Regex.Matches(text, @"/Type\s*/Page\b").Count;
            return Math.Max(1, count);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            watcher?.Dispose();
            base.OnFormClosed(e);
        }

        [STAThread]
        public static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}
