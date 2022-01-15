
using Microsoft.Win32;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

using OfficeHelper.Models;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Controls;
using System.Collections.ObjectModel;

namespace OfficeHelper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string dataPath = AppDomain.CurrentDomain.BaseDirectory + "Files/Data.xlsx";
        private static string selectedExcelPath;
        private List<Data> keyValues = new();
        private List<WordValue> valueIndexes;
        private ObservableCollection<string[]> data = new();
        private bool accepted;
        private int editingIndex;

        public List<Column> Columns { get; set; } = new();
        public string WordPath { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            ISheet sheet;
            using (var stream = File.OpenRead(dataPath))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);

                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    string name = headerRow.GetCell(j).ToString();
                    keyValues.Add(new Data { Name = name, Values = new() });
                }
                for (int i = (headerRow.RowNum + 1); i <= sheet.LastRowNum; i++)
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
                                keyValues.First(x => x.Name == headerRow.GetCell(j).StringCellValue).Values.Add(row.GetCell(j).StringCellValue);
                            }
                        }
                    }
                }
            }
            dgMain.ItemsSource = data;
        }

        private void btnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new();
            dialog.Filter = "Excel fayllar|*.xlsx";
            if (dialog.ShowDialog(this) == true)
            {
                selectedExcelPath = dialog.FileName;
                List<string> rowList = new List<string>();
                ISheet sheet;
                using (var stream = File.OpenRead(selectedExcelPath))
                {
                    stream.Position = 0;
                    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                    sheet = xssWorkbook.GetSheetAt(0);
                    IRow headerRow1 = sheet.GetRow(3);
                    IRow headerRow2 = sheet.GetRow(4);

                    int cellCount = headerRow1.LastCellNum;
                    Columns.Clear();
                    var temp = dgMain.Columns[0];
                    dgMain.Columns.Clear();
                    dgMain.Columns.Add(temp);
                    for (int j = 0; j < cellCount; j++)
                    {
                        string name = headerRow2.GetCell(j).ToString();
                        if (name == null || string.IsNullOrWhiteSpace(name))
                        {
                            name = headerRow1.GetCell(j).ToString();
                        }
                        if (keyValues.Any(x => x.Name == name))
                        {
                            var column = new DataGridComboBoxColumn()
                            {
                                Header = name,
                                ItemsSource = keyValues.First(x => x.Name == name).Values.ToList(),
                                IsReadOnly = false,
                                TextBinding = new Binding($"[{j}]") { ValidatesOnExceptions = false, ValidatesOnNotifyDataErrors = false, ValidatesOnDataErrors = false, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                                EditingElementStyle = (Style)FindResource("ComboBoxEditingStyle"),
                                ElementStyle = (Style)FindResource("TextBlockComboBoxStyle"),
                            };
                            dgMain.Columns.Add(column);
                        }
                        else
                        {
                            dgMain.Columns.Add(new DataGridTextColumn() { Binding = new Binding($"[{j}]"), Header = name });
                        }
                        Columns.Add(new() { Name = $"[{name}]", IsChecked = false });
                    }
                    lbColumns.Items.Refresh();
                    data.Clear();
                    for (int i = (headerRow2.RowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                rowList.Add(row.GetCell(j).ToString());
                            }
                        }
                        if (rowList.Count > 0)
                        {
                            data.Add(rowList.ToArray());
                        }
                        dgMain.Items.Refresh();
                        rowList.Clear();
                    }
                }
            }
            btnExportToExcel.IsEnabled = true;
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
                    valueIndexes = new();
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
                                    valueIndexes.Add(new() { ColumnIndex = h, ParagraphIndex = i, RunIndex = j });
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
            accepted = false;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\winword.exe");
            string exepath = key.GetValue("").ToString();
            ProcessStartInfo info = new ProcessStartInfo()
            {
                FileName = exepath,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            pbMain.Maximum = dgMain.SelectedItems.Count;
            foreach (string[] item in dgMain.SelectedItems)
            {
                using (FileStream sw = File.OpenRead(WordPath))
                {
                    XWPFDocument doc = new XWPFDocument(sw);
                    foreach (var value in valueIndexes)
                    {
                        var text = doc.Paragraphs[value.ParagraphIndex].Runs[value.RunIndex].GetText(0);
                        text = text.Replace(Columns[value.ColumnIndex].Name, item[value.ColumnIndex].ToString());
                        doc.Paragraphs[value.ParagraphIndex].Runs[value.RunIndex].SetText(text, 0);
                    }
                    var fileName = "";
                    using (var file = File.Create(GetFileName(Path.GetFileName(WordPath))))
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
                ShowProcess($"{pbMain.Value + 1}-qator chop etishga berildi!");
            }
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var text = ((Column)((Button)sender).DataContext).Name;
            Clipboard.SetText(text);
        }

        private void btnNewRow_Click(object sender, RoutedEventArgs e)
        {
            data.Add(new string[dgMain.Columns.Count]);
            dgMain.Items.Refresh();
            dgMain.ScrollIntoView(data.Last());
        }

        private void btnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel fayl|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                ISheet sheet;
                using (var stream = File.OpenRead(selectedExcelPath))
                {
                    stream.Position = 0;
                    XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                    sheet = xssWorkbook.GetSheetAt(0);
                    for (int i = 6; i < sheet.LastRowNum + 1; i++)
                    {
                        sheet.RemoveRow(sheet.GetRow(i));
                    }
                    foreach (var item in data)
                    {
                        var row = sheet.CreateRow(sheet.LastRowNum + 1);
                        foreach (var value in item)
                        {
                            row.CreateCell(row.LastCellNum != -1 ? row.LastCellNum : 0).SetCellValue(value);
                        }
                    }
                    using (var fileToSave = File.Open(saveFileDialog.FileName, FileMode.OpenOrCreate))
                    {
                        fileToSave.Flush();
                        xssWorkbook.Write(fileToSave);
                    }
                }
            }
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            dhMain.IsOpen = false;
            data[editingIndex] = (string[])ugDialogHost.DataContext;
            dgMain.Items.Refresh();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            editingIndex = data.IndexOf((string[])((Button)sender).DataContext);
            GenerateEditor(((Button)sender).DataContext);
            dhMain.IsOpen = true;
        }

        private void GenerateEditor(object values)
        {
            ugDialogHost.Height = 0;
            ugDialogHost.Children.Clear();
            ugDialogHost.DataContext = values;
            for (int i = 0; i < Columns.Count; i++)
            {
                string name = Columns[i].Name.Trim('[', ']');
                var tb = new TextBlock
                {
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = 500,
                    Text = name,
                };
                ugDialogHost.Children.Add(tb);
                if (keyValues.Any(x => x.Name == name))
                {
                    var cb = new ComboBox()
                    {
                        ItemsSource = keyValues.First(x => x.Name == name).Values.ToList(),
                        IsEditable = true,
                    };
                    cb.SetBinding(ComboBox.TextProperty, $"[{i}]");
                    ugDialogHost.Children.Add(cb);
                }
                else
                {
                    var tbx = new TextBox()
                    {
                        VerticalContentAlignment = System.Windows.VerticalAlignment.Center
                    };
                    tbx.SetBinding(TextBox.TextProperty, $"[{i}]");
                    ugDialogHost.Children.Add(tbx);
                }
                ugDialogHost.Height += 40;
            }
        }
    }
}
