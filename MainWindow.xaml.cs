
using Microsoft.Win32;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

using OfficeHelper.Models;

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using Forms = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Controls;

namespace OfficeHelper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<WordValue> values;
        private bool accepted;

        public List<Column> Columns { get; set; } = new();
        public string WordPath { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new();
            dialog.Filter = "Excel fayllar|*.xlsx";
            if (dialog.ShowDialog(this) == true)
            {
                DataTable dtTable = new DataTable();
                List<string> rowList = new List<string>();
                ISheet sheet;
                using (var stream = File.OpenRead(dialog.FileName))
                {
                    stream.Position = 0;
                    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                    sheet = xssWorkbook.GetSheetAt(0);
                    IRow headerRow1 = sheet.GetRow(3);
                    IRow headerRow2 = sheet.GetRow(4);

                    int cellCount = headerRow1.LastCellNum;
                    Columns.Clear();
                    dgMain.Columns.Clear();
                    for (int j = 0; j < cellCount; j++)
                    {
                        string name = headerRow2.GetCell(j).ToString();
                        if (name == null || string.IsNullOrWhiteSpace(name))
                        {
                            name = headerRow1.GetCell(j).ToString();
                        }
                        dgMain.Columns.Add(new DataGridTextColumn() { Header = name, Binding = new Binding($"[{j}]") });
                        Columns.Add(new() { Name = $"[{name}]", IsChecked = false });
                    }
                    lbColumns.Items.Refresh();
                    dgMain.Items.Clear();
                    for (int i = (headerRow2.RowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                                {
                                    rowList.Add(row.GetCell(j).ToString());
                                }
                            }
                        }
                        if (rowList.Count > 0)
                        {
                            dgMain.Items.Add(rowList.ToArray());
                        }
                        dgMain.Items.Refresh();
                        rowList.Clear();
                    }
                }
                //dgMain.ItemsSource = dtTable.AsDataView();
            }
        }

        private void btnSelectWord_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new();
            dialog.Filter = "Word fayllar|*.docx";
            if (dialog.ShowDialog(this) == true)
            {
                WordPath = dialog.FileName;
                using (FileStream sw = File.OpenRead(WordPath))
                {
                    XWPFDocument doc = new XWPFDocument(sw);
                    values = new();
                    for (int h = 0; h < Columns.Count; h++)
                    {
                        for (int i = 0; i < doc.Paragraphs.Count; i++)
                        {
                            for (int j = 0; j < doc.Paragraphs[i].Runs.Count; j++)
                            {
                                string text = doc.Paragraphs[i].Runs[j].GetText(0);
                                if (text != null && text.Contains(Columns[h].Name))
                                {
                                    Columns[h].IsChecked = true;
                                    values.Add(new() { ColumnIndex = h, ParagraphIndex = i, RunIndex = j });
                                }
                            }
                        }
                    }
                    lbColumns.Items.Refresh();
                }
            }
        }

        private async void btnStart_Click(object sender, RoutedEventArgs e)
        {
            //PrintDialog printDialog = new PrintDialog();
            //DialogResult result = printDialog.ShowDialog();
            //if (result == System.Windows.Forms.DialogResult.OK)
            //{
            accepted = false;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\winword.exe");
            string exepath = key.GetValue("").ToString();
            ProcessStartInfo info = new ProcessStartInfo()
            {
                FileName = exepath,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            pbMain.Maximum = dgMain.Items.Count;
            for (int i = 0; i < dgMain.Items.Count; i++)
            {
                using (FileStream sw = File.OpenRead(WordPath))
                {
                    XWPFDocument doc = new XWPFDocument(sw);
                    foreach (var value in values)
                    {
                        var text = doc.Paragraphs[value.ParagraphIndex].Runs[value.RunIndex].GetText(0);
                        text = text.Replace(Columns[value.ColumnIndex].Name, ((DataRowView)dgMain.Items[i])[value.ColumnIndex].ToString());
                        doc.Paragraphs[value.ParagraphIndex].Runs[value.RunIndex].SetText(text);
                    }
                    var fileName = "";
                    using (var file = File.Create(GetFileName(System.IO.Path.GetFileName(WordPath))))
                    {
                        doc.Write(file);
                        fileName = file.Name;
                    }
                    info.Arguments = $"\"{fileName}\" /mFilePrintDefault /mFileExit /q /n";
                    if (accepted || MessageBox.Show("Birinchi fayl tayyorlandi, Ko'rib tekshiring!", "Ogohlantirish", MessageBoxButton.OK) == MessageBoxResult.OK)
                    {
                        if (!accepted)
                        {
                            var word = Process.Start(new ProcessStartInfo(exepath, $"\"{fileName}\""));
                            word.WaitForExit();
                        }
                        if (accepted || MessageBox.Show("Sizga ma'qulmi? Chop etilsinmi?", "Ogohlantirish", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                        {
                            accepted = true;
                            var proc = Process.Start(info);
                            while (proc.MainWindowHandle == default(IntPtr))
                            {
                                await Task.Delay(1000);
                            }
                            proc.Kill();
                        }
                        else break;
                    }
                    else if (!accepted) break;
                }

                pbMain.Value++;
                ShowProcess($"{i + 1}-qator chop etishga berildi!");
            }
            //}
        }

        private string GetFileName(string name)
        {
            var dirPath = AppDomain.CurrentDomain.BaseDirectory + "Temp";
            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }
            return $"{dirPath}/{DateTime.Now.Ticks}_{name}";
        }

        private void btnShowProcess_Click(object sender, RoutedEventArgs e)
        {
            if (rowProgress.Height == new GridLength(0))
            {
                rowProgress.Height = new GridLength(200);
                btnShowProcess.Content = "Jarayonni yashirish";
            }
            else
            {
                rowProgress.Height = new GridLength(0);
                btnShowProcess.Content = "Jarayonni ko'rsatish";
            }
        }
        public void ShowProcess(string message)
        {
            Dispatcher.Invoke(() =>
            {
                icProcess.Items.Add(message);
            });
        }
        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://t.me/Programmer1718");
        }

        private void dgMain_AutoGeneratingColumn(object sender, System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs e)
        {
            //var dGrid = (sender as DataGrid);
            //if (dGrid == null) return;
            //var view = dGrid.ItemsSource as DataView;
            //if (view == null) return;
            //var table = view.Table;
            //e.Column.Header = table.Columns[e.Column.Header as String].Caption;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var text = ((Column)((Button)sender).DataContext).Name;
            Clipboard.SetText(text);
        }
    }
}
