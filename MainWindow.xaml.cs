using Microsoft.Win32;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

using OfficeHelper.Models;

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OfficeHelper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<WordValue> Values;
        public Dictionary<string, bool> Columns;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new();
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
                    IRow headerRow = sheet.GetRow(5);

                    Columns = new();
                    foreach (var item in headerRow.Cells.Where(x=>!string.IsNullOrWhiteSpace(x.StringCellValue)))
                    {
                        var key = item.StringCellValue;
                        while (Columns.ContainsKey(key))
                        {
                            key += "*";
                        }
                        Columns.Add($"[{item.StringCellValue}]", false);
                    }
                    lbColumns.Items.Refresh();

                    int cellCount = headerRow.LastCellNum;
                    for (int j = 0; j < cellCount; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                        {
                            dtTable.Columns.Add(cell.ToString());
                        }
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
                                    rowList.Add(row.GetCell(j).ToString());
                                }
                            }
                        }
                        if (rowList.Count > 0)
                            dtTable.Rows.Add(rowList.ToArray());
                        rowList.Clear();
                    }
                }
                dgMain.ItemsSource = dtTable.DefaultView;
            }
        }

        private void btnSelectWord_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new();
            dialog.Filter = "Word fayllar|*.docx";
            if (dialog.ShowDialog(this) == true)
            {
                using (FileStream sw = File.OpenRead(dialog.FileName))
                {
                    XWPFDocument doc = new XWPFDocument(sw);
                    Values = new();
                    foreach (var column in Columns)
                    {
                        for (int i = 0; i < doc.Paragraphs.Count; i++)
                        {
                            for (int j = 0; j < doc.Paragraphs[i].Runs.Count; j++)
                            {
                                string text = doc.Paragraphs[i].Runs[j].GetText(0);
                                if (text != null && text.Contains(column.Key))
                                {
                                    Columns[column.Key] = true;
                                    Values.Add(new() { ColumnKey = column.Key, ParagraphIndex = i, RunIndex = j });
                                }
                            }
                        }
                    }
                    lbColumns.Items.Refresh();
                }
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in dgMain.Items)
            {
                foreach (var value in Values)
                {

                }
            }
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
    }
}
