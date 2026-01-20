using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;
using Spire.Doc;
using Spire.Xls;

namespace PrintCalc {
    public class Form1 : Form {
        private CheckedListBox clbFiles;
        private Label lblTotal;
        private string watchPath = "";
        private FileSystemWatcher watcher;
        private Dictionary<string, double> filePrices = new Dictionary<string, double>();

        public Form1() {
            this.Text = "Ape Kade Pro - WhatsApp Live Sync";
            this.Size = new Size(650, 750);
            this.BackColor = Color.FromArgb(20, 20, 20);
            this.ForeColor = Color.White;

            // Step 1: Folder Selection Header
            Panel pnlHeader = new Panel() { Dock = DockStyle.Top, Height = 100, BackColor = Color.FromArgb(40, 40, 40) };
            Button btnSelectFolder = new Button() { 
                Text = "Step 1: Select WhatsApp Download Folder", 
                Size = new Size(350, 40), Location = new Point(20, 30), 
                BackColor = Color.FromArgb(0, 150, 136), FlatStyle = FlatStyle.Flat 
            };
            btnSelectFolder.Click += SelectFolder;
            pnlHeader.Controls.Add(btnSelectFolder);
            this.Controls.Add(pnlHeader);

            // Step 2: List of Files
            clbFiles = new CheckedListBox() { 
                Dock = DockStyle.Fill, BackColor = Color.Black, ForeColor = Color.Lime, 
                BorderStyle = BorderStyle.None, Font = new Font("Consolas", 10), CheckOnClick = true 
            };
            clbFiles.ItemCheck += (s, e) => BeginInvoke(new Action(CalculateTotal));
            this.Controls.Add(clbFiles);

            // Step 3: Bill Display
            lblTotal = new Label() { 
                Text = "Bill: Rs. 0", Font = new Font("Segoe UI", 36, FontStyle.Bold), 
                Dock = DockStyle.Bottom, Height = 100, TextAlign = ContentAlignment.MiddleCenter, 
                ForeColor = Color.Yellow, BackColor = Color.FromArgb(30, 30, 30) 
            };
            this.Controls.Add(lblTotal);
        }

        private void SelectFolder(object sender, EventArgs e) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                if (fbd.ShowDialog() == DialogResult.OK) {
                    watchPath = fbd.SelectedPath;
                    StartWatching();
                    MessageBox.Show("Syncing Started! Dan WhatsApp eke file ekak download unama methanata auto eyi.");
                }
            }
        }

        private void StartWatching() {
            if (watcher != null) watcher.Dispose();
            watcher = new FileSystemWatcher(watchPath);
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            watcher.Created += (s, e) => {
                System.Threading.Thread.Sleep(2000); // File eka save wenna welawa denawa
                this.Invoke(new Action(() => AddFileToList(e.FullPath)));
            };
            watcher.EnableRaisingEvents = true;
        }

        private void AddFileToList(string path) {
            try {
                int p = 1; string ext = Path.GetExtension(path).ToLower();
                if (ext == ".pdf") p = GetPdfPages(path);
                else if (ext == ".docx" || ext == ".doc") { Document d = new Document(path); p = d.PageCount; }
                else if (ext == ".xlsx" || ext == ".xls") { Workbook w = new Workbook(); w.LoadFromFile(path); p = w.Worksheets.Count; }
                
                double price = ((p / 2) * 15) + ((p % 2) * 10);
                string itemText = $"[{DateTime.Now:HH:mm}] {Path.GetFileName(path)} -> Rs.{price}";
                clbFiles.Items.Insert(0, itemText); // Aluth file eka udatama enawa
                clbFiles.SetItemChecked(0, true);
                filePrices[itemText] = price;
                CalculateTotal();
            } catch { }
        }

        private void CalculateTotal() {
            double total = 0;
            foreach (var item in clbFiles.CheckedItems) {
                if (filePrices.ContainsKey(item.ToString())) total += filePrices[item.ToString()];
            }
            lblTotal.Text = $"Bill: Rs. {total}";
        }

        static int GetPdfPages(string path) {
            using StreamReader sr = new StreamReader(path);
            return System.Text.RegularExpressions.Regex.Matches(sr.ReadToEnd(), @"/Type\s*/Page[^s]").Count;
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.Run(new Form1()); }
    }
}
