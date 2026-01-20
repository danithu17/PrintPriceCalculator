using System;
using System.IO;
using System.Windows.Forms;
using System.Linq;

namespace PrintCalc {
    class Program {
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "PDF & Images|*.pdf;*.jpg;*.png;*.jpeg" };
            
            if (ofd.ShowDialog() == DialogResult.OK) {
                double totalBill = 0;
                string summary = "Bill Summary:\n\n";

                foreach (string file in ofd.FileNames) {
                    int pCount = 1; // Default photo ekakata 1i
                    if (Path.GetExtension(file).ToLower() == ".pdf") {
                        // PDF ekak nam pages check karanna library ekak ona (Meka setup ekedi wenawa)
                        pCount = GetPdfPageCount(file); 
                    }

                    // Oyage Logic: (Pairs x 15) + (Ithuru x 10)
                    double price = ((pCount / 2) * 15) + ((pCount % 2) * 10);
                    totalBill += price;
                    summary += $"{Path.GetFileName(file)} ({pCount} pgs) - Rs.{price}\n";
                }

                MessageBox.Show($"{summary}\n------------------\nTotal: Rs.{totalBill}", "Ape Kade Bill");
            }
        }

        static int GetPdfPageCount(string path) {
            using (StreamReader sr = new StreamReader(path)) {
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"/Type\s*/Page[^s]");
                return regex.Matches(sr.ReadToEnd()).Count;
            }
        }
    }
}
