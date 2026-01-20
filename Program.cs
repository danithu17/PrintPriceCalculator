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
        private TextBox txtPhone;
        private DateTimePicker dtStart;
        private Dictionary<string, double> filePrices = new Dictionary<string, double>();
        private string downloadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

        public Form1() {
            this.Text = "Ape Kade Pro - Search & Select Print";
            this.Size = new Size(700, 800);
            this.BackColor = Color.FromArgb(25, 25, 25);
            this.ForeColor = Color.White;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Search Panel
            Panel pnlHeader = new Panel() { Dock = DockStyle.Top, Height = 180, BackColor = Color.FromArgb(40, 40, 40), Padding = new Padding(15) };
            
            pnlHeader.Controls.Add(new Label() { Text = "Phone (Last 3):", Location = new Point(20, 25), Width = 120 });
            txtPhone = new TextBox() { Location = new Point(150, 25), Width = 100, BackColor = Color.Black, ForeColor = Color.Yellow, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            
            pnlHeader.Controls.Add(new Label() { Text = "Select Date:", Location = new Point(20, 65), Width = 120 });
            dtStart = new DateTimePicker() { Location = new Point(150, 65), Width = 150, Format = DateTimePickerFormat.Short };

            Button btnSearch = new Button() { Text = "FIND CUSTOMER FILES", Location = new Point(150, 110), Size = new Size(250, 40), BackColor = Color.FromArgb(0, 122, 204), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            btnSearch.Click += SearchFiles;

            pnlHeader.Controls.AddRange(new Control[] { txtPhone, dtStart, btnSearch });
            this.Controls.Add(pnlHeader);

            // Selectable List
            Label lblInstr = new Label() { Text = "Print karanna ona files select karanna:", Dock = DockStyle.Top, Height = 30, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.Gray };
            this.Controls.Add(lblInstr);

            clbFiles = new CheckedListBox() { Dock = DockStyle.Fill, BackColor = Color.Black, ForeColor = Color.Lime, BorderStyle = BorderStyle.None, Font = new Font("Consolas", 10), CheckOnClick = true };
            clbFiles.ItemCheck += (s, e) => BeginInvoke(new Action(CalculateTotal));
            this.Controls.Add(clbFiles);

            // Bottom Total
            lblTotal = new Label() { Text = "Total: Rs. 0", Font = new Font("Segoe UI", 36, FontStyle.Bold), Dock = DockStyle.Bottom, Height = 120, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Yellow, BackColor = Color.FromArgb(35, 35, 35) };
            this.Controls.Add(lblTotal);
        }

        private void SearchFiles(object sender, EventArgs e) {
            clbFiles.Items.Clear();
            filePrices.Clear();
            string phone = txtPhone.Text.Trim();
            
            if (!Directory.Exists(downloadPath)) return;

            var files = Directory.GetFiles(downloadPath)
                .Select(f => new FileInfo(f))
                .Where(f => f.CreationTime.Date == dtStart.Value.Date)
                .Where(f => string.IsNullOrEmpty(phone) || f.Name.Contains(phone))
                .OrderByDescending(f => f.CreationTime)
                .ToList();

            foreach (var file in files) {
                double price = GetFilePrice(file.FullName);
                string itemText = $"{file.Name} -> Rs.{price}";
                clbFiles.Items.Add(itemText, true); // Default okkoma select wela enne
                filePrices[itemText] = price;
            }
            CalculateTotal();
        }

        private double GetFilePrice(string path) {
            try {
                int p = 1; string ext = Path.GetExtension(path).ToLower();
                if (ext == ".pdf") p = GetPdfPages(path);
                else if (ext == ".docx" || ext == ".doc") { Document d = new Document(path); p = d.PageCount; }
                else if (ext == ".xlsx" || ext == ".xls") { Workbook w = new Workbook(); w.LoadFromFile(path); p = w.Worksheets.Count; }
                
                return ((p / 2) * 15) + ((p % 2) * 10);
            } catch { return 10; } // Error unoth default Rs.10
        }

        private void CalculateTotal() {
            double total = 0;
            foreach (var item in clbFiles.CheckedItems) {
                if (filePrices.ContainsKey(item.ToString())) total += filePrices[item.ToString()];
            }
            lblTotal.Text = $"Total: Rs. {total}";
        }

        static int GetPdfPages(string path) {
            using StreamReader sr = new StreamReader(path);
            return System.Text.RegularExpressions.Regex.Matches(sr.ReadToEnd(), @"/Type\s*/Page[^s]").Count;
        }

        [STAThread] static void Main() { Application.EnableVisualStyles(); Application.Run(new Form1()); }
    }
}
