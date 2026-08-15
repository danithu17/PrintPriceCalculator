using System;
using System.Collections.Generic;
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
        private readonly Dictionary<string, decimal> prices = new();

        public Form1()
        {
            Text = "Shop Print Price Calculator";
            Width = 760;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            var top = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 150,
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

            top.Controls.Add(new Label { Text = "Simplex Rs.", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
            simplexRate.Minimum = 0;
            simplexRate.Maximum = 100000;
            simplexRate.Value = 10;
            simplexRate.Width = 70;
            top.Controls.Add(simplexRate);

            top.Controls.Add(new Label { Text = "Duplex Rs.", AutoSize = true, Padding = new Padding(10, 8, 0, 0) });
            duplexRate.Minimum = 0;
            duplexRate.Maximum = 100000;
            duplexRate.Value = 15;
            duplexRate.Width = 70;
            top.Controls.Add(duplexRate);

            var addPdf = new Button { Text = "Add PDF", Width = 100, Height = 35 };
            addPdf.Click += (_, _) => AddPdfFromDialog();
            top.Controls.Add(addPdf);

            foreach (var control in new Control[] { mode, copies, simplexRate, duplexRate })
                control.TextChanged += (_, _) => Recalculate();
            copies.ValueChanged += (_, _) => Recalculate();
            mode.SelectedIndexChanged += (_, _) => Recalculate();

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
            using var dialog = new FolderBrowserDialog { Description = "Select the folder where WhatsApp Desktop saves PDFs" };
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
            using var dialog = new OpenFileDialog { Filter = "PDF files (*.pdf)|*.pdf", Multiselect = true };
            if (dialog.ShowDialog() != DialogResult.OK) return;
            foreach (var file in dialog.FileNames) AddFile(file);
        }

        private void AddFile(string path)
        {
            if (!File.Exists(path) || Path.GetExtension(path).ToLowerInvariant() != ".pdf") return;
            try
            {
                for (int i = 0; i < 10; i++)
                {
                    try
                    {
                        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        int pages = CountPdfPages(stream);
                        string key = $"{Path.GetFileName(path)} | {pages} pages";
                        if (!prices.ContainsKey(key))
                        {
                            files.Items.Insert(0, key);
                            files.SetItemChecked(0, true);
                            prices[key] = 0;
                            UpdateItemPrice(key, pages);
                        }
                        return;
                    }
                    catch { System.Threading.Thread.Sleep(300); }
                }
            }
            catch { }
        }

        private void UpdateItemPrice(string key, int pages)
        {
            decimal rate = mode.SelectedIndex == 1 ? duplexRate.Value : simplexRate.Value;
            decimal price = pages * copies.Value * rate;
            prices[key] = price;
            Recalculate();
        }

        private void Recalculate()
        {
            decimal total = 0;
            var rate = mode.SelectedIndex == 1 ? duplexRate.Value : simplexRate.Value;
            for (int i = 0; i < files.Items.Count; i++)
            {
                string key = files.Items[i].ToString() ?? "";
                var match = Regex.Match(key, @"\|\s*(\d+)\s+pages$");
                if (!match.Success) continue;
                int pages = int.Parse(match.Groups[1].Value);
                decimal price = pages * copies.Value * rate;
                prices[key] = price;
                if (files.GetItemChecked(i)) total += price;
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
