using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PrintCalc
{
    public class Form1 : Form
    {
        private readonly CheckedListBox files = new();
        private readonly Label totalLabel = new();
        private readonly NumericUpDown simplexRate = new();
        private readonly NumericUpDown duplexRate = new();
        private readonly NumericUpDown copies = new();
        private readonly ComboBox mode = new();
        private FileSystemWatcher? watcher;

        public Form1()
        {
            Text = "Shop Print Price Calculator - PDF Pages";
            Width = 780;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            var top = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 180,
                Padding = new Padding(10),
                AutoSize = false
            };

            var folderButton = new Button { Text = "Select WhatsApp PDF Folder", Width = 210, Height = 35 };
            folderButton.Click += (_, _) => SelectFolder();
            top.Controls.Add(folderButton);

            top.Controls.Add(new Label { Text = "Mode", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
            mode.Items.AddRange(new object[] { "Simplex", "Duplex" });
            mode.SelectedIndex = 0;
            mode.Width = 100;
            top.Controls.Add(mode);

            top.Controls.Add(new Label { Text = "Copies", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
            copies.Minimum = 1;
            copies.Maximum = 9999;
            copies.Value = 1;
            copies.Width = 70;
            top.Controls.Add(copies);

            top.Controls.Add(new Label { Text = "Simplex Rs./page", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
            simplexRate.Minimum = 0;
            simplexRate.Maximum = 100000;
            simplexRate.Value = 10;
            simplexRate.Width = 70;
            top.Controls.Add(simplexRate);

            top.Controls.Add(new Label { Text = "Duplex Rs./sheet (2 pages)", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
            duplexRate.Minimum = 0;
            duplexRate.Maximum = 100000;
            duplexRate.Value = 15;
            duplexRate.Width = 80;
            top.Controls.Add(duplexRate);

            var addPdf = new Button { Text = "Add PDF", Width = 100, Height = 35 };
            addPdf.Click += (_, _) => AddPdfFromDialog();
            top.Controls.Add(addPdf);

            var help = new Label
            {
                Text = "Duplex: odd page count +1, then count every 2 pages as 1 sheet",
                AutoSize = true,
                Padding = new Padding(10, 8, 0, 0)
            };
            top.Controls.Add(help);

            mode.SelectedIndexChanged += (_, _) => Recalculate();
            copies.ValueChanged += (_, _) => Recalculate();
            simplexRate.ValueChanged += (_, _) => Recalculate();
            duplexRate.ValueChanged += (_, _) => Recalculate();

            files.Dock = DockStyle.Fill;
            files.CheckOnClick = true;
            files.HorizontalScrollbar = true;
            files.ItemCheck += (_, e) => BeginInvoke(new Action(Recalculate));

            totalLabel.Dock = DockStyle.Bottom;
            totalLabel.Height = 90;
            totalLabel.Text = "Total: Rs. 0";
            totalLabel.TextAlign = ContentAlignment.MiddleCenter;
            totalLabel.Font = new Font("Segoe UI", 28, FontStyle.Bold);

            Controls.Add(files);
            Controls.Add(totalLabel);
            Controls.Add(top);
        }

        private void SelectFolder()
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "Select the folder where WhatsApp Desktop saves PDFs"
            };

            if (dialog.ShowDialog() != DialogResult.OK) return;

            watcher?.Dispose();
            watcher = new FileSystemWatcher(dialog.SelectedPath, "*.pdf")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                EnableRaisingEvents = true
            };
            watcher.Created += (_, e) => BeginInvoke(new Action(() => AddFile(e.FullPath)));
            watcher.Renamed += (_, e) => BeginInvoke(new Action(() => AddFile(e.FullPath)));

            MessageBox.Show("Watching started. New PDF files in this folder will be added automatically.");
        }

        private void AddPdfFromDialog()
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Multiselect = true
            };

            if (dialog.ShowDialog() != DialogResult.OK) return;
            foreach (var file in dialog.FileNames) AddFile(file);
        }

        private void AddFile(string path)
        {
            if (!File.Exists(path) || Path.GetExtension(path).ToLowerInvariant() != ".pdf") return;

            for (int attempt = 0; attempt < 10; attempt++)
            {
                try
                {
                    using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int pages = CountPdfPages(stream);
                    string key = $"{Path.GetFileName(path)} | {pages} pages";

                    bool exists = false;
                    foreach (var item in files.Items)
                    {
                        if (string.Equals(item?.ToString(), key, StringComparison.OrdinalIgnoreCase))
                        {
                            exists = true;
                            break;
                        }
                    }

                    if (!exists)
                    {
                        files.Items.Insert(0, key);
                        files.SetItemChecked(0, true);
                        Recalculate();
                    }
                    return;
                }
                catch
                {
                    System.Threading.Thread.Sleep(300);
                }
            }
        }

        private void Recalculate()
        {
            decimal total = 0;
            decimal copyCount = copies.Value;

            for (int i = 0; i < files.Items.Count; i++)
            {
                string key = files.Items[i]?.ToString() ?? "";
                var match = Regex.Match(key, @"\|\s*(\d+)\s+pages$");
                if (!match.Success || !files.GetItemChecked(i)) continue;

                int pages = int.Parse(match.Groups[1].Value);
                decimal pricePerCopy;

                if (mode.SelectedIndex == 1)
                {
                    // Duplex rule: if page count is odd, add one page for pricing.
                    // Then every 2 pages count as 1 sheet at the duplex rate.
                    int billablePages = (pages % 2 == 0) ? pages : pages + 1;
                    int sheets = billablePages / 2;
                    pricePerCopy = sheets * duplexRate.Value;
                }
                else
                {
                    // Simplex: every PDF page is one printed sheet at Rs.10.
                    pricePerCopy = pages * simplexRate.Value;
                }

                total += pricePerCopy * copyCount;
            }

            totalLabel.Text = $"Total: Rs. {total:N2}";
        }

        private static int CountPdfPages(Stream stream)
        {
            using var reader = new StreamReader(stream);
            string text = reader.ReadToEnd();
            int count = Regex.Matches(text, @"/Type\s*/Page\b").Count;
            return Math.Max(1, count);
        }

        [STAThread]
        public static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}
