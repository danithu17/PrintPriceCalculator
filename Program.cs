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
        private ListBox lstDetails;
        private Label lblTotal;
        private TextBox txtPhone;
        private DateTimePicker dtStart;
        private double totalAmount = 0;
        private FileSystemWatcher watcher;
        private string downloadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

        public Form1() {
            this.Text = "Ape Kade Smart Print Manager - WhatsApp Auto-Scanner";
            this.Size = new Size(650, 750);
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Search Panel
            Panel pnlHeader = new Panel() { Dock = DockStyle.Top, Height = 180, BackColor = Color.FromArgb(45, 45, 48), Padding = new Padding(15) };
            
            pnlHeader.Controls.Add(new Label() { Text = "WhatsApp Number (Last 3):", Location = new Point(20, 25), Width = 180 });
            txtPhone = new TextBox() { Location = new Point(200, 25), Width = 100, BackColor = Color.Black, ForeColor = Color.Yellow, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            
            pnlHeader.Controls.Add(new Label() { Text = "Selected Date:", Location = new Point(20, 65), Width = 180 });
            dtStart = new DateTimePicker() { Location = new Point(200, 65), Width = 130, Format = DateTimePickerFormat.Short };

            Button btnScan = new Button() { Text = "SCAN ALL FILES", Location = new Point(20, 110), Size = new Size(280, 40), BackColor = Color.SeaGreen, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            btnScan.Click += (s, e) => RunManualScan();

            pnlHeader.Controls.AddRange(new Control[] { txtPhone, dtStart, btnScan });
            this.Controls.Add(pnlHeader);

            // Result Area
            lstDetails = new ListBox() { Location = new Point(20, 200), Size = new Size(590, 380), BackColor = Color.Black, ForeColor = Color.Lime, BorderStyle = BorderStyle.None, Font = new Font("Consolas", 10) };
            this.Controls.Add(lstDetails);

            lblTotal = new Label() { Text = "Total: Rs. 0", Font = new Font("Segoe UI", 36, FontStyle.Bold), Dock = DockStyle.Bottom, Height = 100, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Yellow };
            this.Controls.Add(lblTotal);

            // Auto-Scanner Setup
            SetupAutoWatcher();
        }

        private void SetupAutoWatcher() {
            if (!Directory.Exists(downloadPath)) return;
            watcher = new FileSystemWatcher(downloadPath);
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            watcher.Created += (s, e) => {
                System.Threading.Thread.Sleep(2000); // Wait for file to fully land
                this.Invoke(new Action(() => ProcessSingleFile(e.FullPath)));
            };
            watcher.EnableRaisingEvents = true;
        }

        private void RunManualScan() {
            totalAmount = 0; lstDetails.Items.Clear();
            string phone = txtPhone.Text.Trim();
            
            var files = Directory.GetFiles(downloadPath)
                .Select(f => new FileInfo(f))
                .Where(f => f.CreationTime.Date == dtStart.Value.Date)
                .Where(f => string.IsNullOrEmpty(phone) || f.Name.Contains(phone))
                .ToList();

            foreach (var file in files) ProcessSingleFile(file.FullName);
            lblTotal.Text = $"Total: Rs. {totalAmount}";
        }

        private void ProcessSingleFile(string path) {
            try {
                int p = 1; string ext = Path.GetExtension(path).ToLower();
                if (ext == ".pdf") p = GetPdfPages(path);
                else if (ext == ".docx" || ext == ".doc") { Document d = new Document(path); p = d.PageCount; }
                else if (ext == ".xlsx" || ext == ".xls") { Workbook w = new Workbook(); w.LoadFromFile(path); p = w.Worksheets.Count; }
                
                double price = ((p / 2) * 15) + ((p % 2) * 10);
                totalAmount += price;
                lstDetails.Items.Add($"{Path.GetFileName(path)} ({p} pgs) -> Rs.{price}");
                lblTotal.Text = $"Total: Rs. {totalAmount}";
            } catch { }
        }

        static int GetPdfPages(string path) {
            using StreamReader sr = new StreamReader(path);
            return System.Text.RegularExpressions.Regex.Matches(sr.ReadToEnd(), @"/Type\s*/Page[^s]").Count;
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.Run(new Form1()); }
    }
}
