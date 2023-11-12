using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Shapes;

namespace CsvToExcelTransformer_Win
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectFile(object sender, EventArgs e)
        {
            string filePath = PickFile(); ;
            if (!string.IsNullOrEmpty(filePath))
            {
                EntryBox.Text = filePath;
            }
        }

        private string PickFile()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "CSV files (*.csv)|*.csv";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            var path = String.Empty;
            var dialog = openFileDialog1.ShowDialog();

            if (dialog == true)
            {
                path = openFileDialog1.FileName;
            }

            return path;
        }

        public async Task DisplayAlertAsync(string title, string message, string cancel)
        {
            MessageBox.Show(message, title);
        }

        private async void GenerateExcelFile(object sender, EventArgs e)
        {
            var filePath = EntryBox.Text;

            if (String.IsNullOrEmpty(filePath))
            {
                await DisplayAlertAsync("Błąd", "Proszę wskazać plik CSV", "OK");
                return;
            }
            if (!File.Exists(filePath))
            {
                await DisplayAlertAsync("Błąd", "Plik nie istnieje", "OK");
                return;
            }
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Read CSV data
                string[] csvLines = File.ReadAllLines(filePath);
                worksheet.Cells[1, 1].Value = "ID zgłoszenia";
                worksheet.Cells[1, 2].Value = "Data";
                worksheet.Cells[1, 3].Value = "Status";
                worksheet.Cells[1, 4].Value = "Imię";
                worksheet.Cells[1, 5].Value = "Nazwisko";
                worksheet.Cells[1, 6].Value = "Email";
                worksheet.Cells[1, 7].Value = "Telefon";
                worksheet.Cells[1, 8].Value = "Zainteresowanie";
                worksheet.Cells[1, 9].Value = "I zgoda";
                worksheet.Cells[1, 10].Value = "II zgoda";
                worksheet.Cells[1, 11].Value = "III zgoda";
                worksheet.Cells[1, 12].Value = "Status";
                worksheet.Cells[1, 13].Value = "Notatki";
                for (int i = 1; i < csvLines.Length; i++)
                {
                    CsvLine csv = CsvLine.FromCsv(csvLines[i]);
                    worksheet.Cells[i + 1, 1].Value = csv.Id;
                    worksheet.Cells[i + 1, 2].Value = csv.Date;
                    worksheet.Cells[i + 1, 3].Value = csv.Status;
                    worksheet.Cells[i + 1, 4].Value = csv.Name.Trim();
                    worksheet.Cells[i + 1, 5].Value = csv.Lastname.Trim();
                    worksheet.Cells[i + 1, 6].Value = csv.Email;
                    string phoneNumber;
                    try
                    {
                        if (csv.Tel.StartsWith('+'))
                        {
                            phoneNumber = String.Format("{0:+## ###-###-###}", csv.Tel);
                        }
                        else
                        {
                            phoneNumber = String.Format("{0:###-###-###}", csv.Tel);
                        }
                    }
                    catch
                    {
                        phoneNumber = csv.Tel;
                    }
                    worksheet.Cells[i + 1, 7].Value = phoneNumber;
                    worksheet.Cells[i + 1, 8].Value = csv.Interest;
                    worksheet.Cells[i + 1, 9].Value = csv.FirstAgr;
                    worksheet.Cells[i + 1, 10].Value = csv.SecondAgr;
                    worksheet.Cells[i + 1, 11].Value = csv.ThirdAgr;
                    worksheet.Cells[i + 1, 12].Value = string.Empty;
                    worksheet.Cells[i + 1, 13].Value = string.Empty;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Save Excel file
                var outPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), $"Baza  danych {DateTime.Now.ToString("yyyy-MM-dd HH-mm")}.xlsx");
                FileInfo excelFile = new FileInfo(outPath);
                excelPackage.SaveAs(excelFile);

                await DisplayAlertAsync("Sukces", $"Plik Excel został wygenerowany:\n{outPath}", "OK");
                return;
            }
        }
    }
}
