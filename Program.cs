using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace PrintCalc {
    public class Form1 : Form {
        private RadioButton rbBW, rbColor;
        private ListBox lstDetails;
        private Label lblTotal;

        public Form1() {
            this.Text = "Ape Kade Print Master";
            this.Size = new Size(450, 500);
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;

            Label title = new Label() { Text = "Print Cost Calculator", Font = new Font("Arial", 16, FontStyle.Bold), Dock = DockStyle.Top, Height = 50, TextAlign = ContentAlignment.MiddleCenter };
            
            rbBW = new RadioButton() { Text = "B&W (Depatta Logic)", Location = new Point(50, 70), Checked = true, Width = 150 };
            rbColor = new RadioButton() { Text = "Color (Rs. 50)", Location = new Point(220, 70), Width = 150 };

            Button btnSelect = new Button() { Text = "Select Files (PDF/Photos)", Location = new Point(50, 110), Size = new Size(320, 40), BackColor = Color.ForestGreen, FlatStyle = FlatStyle.Flat };
            btnSelect.Click += ProcessFiles;

            lstDetails = new ListBox() { Location = new Point(50, 170), Size = new Size(320, 180), BackColor = Color.Black, ForeColor = Color.Lime };
            lblTotal = new Label() { Text = "Total: Rs. 0", Font = new Font("Arial", 20, FontStyle.Bold), Location = new Point(50, 370), Size = new Size(320, 50), TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Gold };

            this.Controls.AddRange(new Control[] { title, rbBW, rbColor, btnSelect, lstDetails, lblTotal });
        }

        private void ProcessFiles(object? sender, EventArgs e) {
            OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "Printable|*.pdf;*.jpg;*.png;*.jpeg" };
            if (ofd.ShowDialog() == DialogResult.OK) {
                double total = 0; lstDetails.Items.Clear();
                foreach (string file in ofd.FileNames) {
                    int p = Path.GetExtension(file).ToLower() == ".pdf" ? GetPdfPageCount(file) : 1;
                    double price = rbBW.Checked ? ((p / 2) * 15) + ((p % 2) * 10) : p * 50;
                    total += price;
                    lstDetails.Items.Add($"{Path.GetFileName(file)} ({p} pgs) -> Rs.{price}");
                }
                lblTotal.Text = $"Total: Rs. {total}";
            }
        }

        static int GetPdfPageCount(string path) {
            try {
                using StreamReader sr = new StreamReader(path);
                return System.Text.RegularExpressions.Regex.Matches(sr.ReadToEnd(), @"/Type\s*/Page[^s]").Count;
            } catch { return 1; }
        }

        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.Run(new Form1());
        }
    }
}
