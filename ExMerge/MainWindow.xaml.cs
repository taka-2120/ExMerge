using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExMerge
{
    public partial class MainWindow : Window
    {
        private List<string> files = [];
        private const int RowLimit = 44;
        private string? outputDirectory = null;

        public MainWindow()
        {
            InitializeComponent();
            issueMonthTextBox.Text = (DateTime.Now.Month - 1).ToString();
            outputFileNameTextBox.Text = $"支払明細書{issueMonthTextBox.Text}月";
        }

        private void mergeButton_click(object sender, RoutedEventArgs e)
        {
            if (files.Count == 0)
            {
                MessageBox.Show(this, "ファイルを選択してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrEmpty(outputFileNameTextBox.Text))
            {
                MessageBox.Show(this, "出力ファイル名を入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrEmpty(issueMonthTextBox.Text))
            {
                MessageBox.Show(this, "請求月を入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrEmpty(outputDirectory))
            {
                MessageBox.Show(this, "ファイルの保存場所を選択してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }


            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var payments = new List<Payment>();

            foreach (var file in files)
            {
                var package = new ExcelPackage(new FileInfo(file));
                var sheet = package.Workbook.Worksheets[0];
                var rowCount = 4;
                while (true)
                {
                    var code = sheet.Cells[rowCount, 1].GetValue<string>();
                    var name = sheet.Cells[rowCount, 2].GetValue<string>();
                    var amount = sheet.Cells[rowCount, 5].GetValue<int?>();
                    if (string.IsNullOrEmpty(code) || string.IsNullOrEmpty(name))
                    {
                        break;
                    }

                    var payment = new Payment
                    {
                        Code = code,
                        Name = name,
                        Amount = amount ?? 0,
                    };
                    payments.Add(payment);
                    rowCount++;
                }
            }

            // New file
            var outputPath = Path.Combine(outputDirectory, $"{outputFileNameTextBox.Text}.xlsx");
            if (File.Exists(outputPath))
            {
                try
                {
                    File.Delete(outputPath);
                }
                catch (IOException error)
                {
                    MessageBox.Show(this, $"ファイルの削除に失敗しました。\n{error.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            /*try
            {*/
            var outputFile = new FileInfo(outputPath);
            var outputPackage = new ExcelPackage(outputFile);
            outputPackage.Workbook.Worksheets.Add("Sheet1");

            // Sort payments by code
            payments.Sort((lhs, rhs) => lhs.Code.CompareTo(rhs.Code));

            // Output
            var codeHistory = new List<string>();
            var pageProgress = 1;
            var totalRowProgress = 0;
            var totalAmount = 0;
            while (payments.Count > 0)
            {
                var sheet = createFormat(outputPackage, int.Parse(issueMonthTextBox.Text), totalRowProgress);

                // Page number
                var pageCell = sheet.Cells[1 + totalRowProgress, 10];
                pageCell.Value = $"No. {pageProgress}";

                // Increment for header rows
                totalRowProgress += 4;

                var computedPaymentCount = 0;
                var subTotal = 0;
                foreach (var payment in payments)
                {
                    if (computedPaymentCount >= RowLimit)
                    {
                        break;
                    }

                    // Operation per Page
                    computedPaymentCount += 1;
                    totalRowProgress += 1;

                    int matchedCount = payments.FindAll(p => p.Code == payment.Code).Count - 1;
                    if (!codeHistory.Contains(payment.Code) && matchedCount > 0)
                    {
                        // If matchedCount over row limit, skip merging
                        if (matchedCount + computedPaymentCount > RowLimit)
                        {
                            totalRowProgress += RowLimit - computedPaymentCount;
                            break;
                        }

                        sheet.Cells[$"B{totalRowProgress}:B{totalRowProgress + matchedCount}"].Merge = true;

                        if (payment.Amount != 0)
                        {
                            sheet.Cells[totalRowProgress, 4].Value = payment.Amount;
                        }
                    }

                    if (!codeHistory.Contains(payment.Code))
                    {
                        var codeCell = sheet.Cells[totalRowProgress, 2];
                        codeCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        codeCell.Value = payment.Code;
                        codeCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    var nameCell = sheet.Cells[totalRowProgress, 3];
                    nameCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    nameCell.Value = Utils.fullToHalf(payment.Name);
                    nameCell.Style.ShrinkToFit = true;

                    if (payment.Amount != 0)
                    {
                        sheet.Cells[totalRowProgress, 4].Value = payment.Amount;
                    }

                    for (int j = 0; j < 9; j++)
                    {
                        // sheet.Cells[rowProgress, j + 1].Style.Border = 
                    }

                    codeHistory.Add(payment.Code);
                    subTotal += payment.Amount;
                }

                // Remove processed payments
                payments.RemoveAll(payment => codeHistory.Contains(payment.Code));

                totalRowProgress += 1;

                // Add Sub Total
                sheet.Cells[totalRowProgress, 3].Value = "小計";
                sheet.Cells[totalRowProgress, 4].Value = subTotal;
                totalAmount += subTotal;

                totalRowProgress += 1;

                if (payments.Count == 0)
                {
                    // Add Total
                    sheet.Cells[totalRowProgress, 3].Value = "合計";
                    sheet.Cells[totalRowProgress, 4].Value = totalAmount;

                    totalRowProgress += 1;
                }
                pageProgress += 1;
            }

            outputPackage.Save();

            MessageBox.Show(this, $"ファイルの結合に成功しました。\nファイルの保存場所: {outputDirectory}", "メッセージ", MessageBoxButton.OK, MessageBoxImage.Information);
            /*} catch (Exception error)
            {
                MessageBox.Show(this, $"Excelファイルの操作に失敗しました。\n{error.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }*/
        }

        private void chooseButton_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excelファイル|*.xlsx",
                Title = "結合するファイルの選択",
                DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments),
            };
            if (fileDialog.ShowDialog() == false)
            {
                return;
            }
            if (fileDialog.FileNames.Length == 0)
            {
                return;
            }
            var selectedFiles = fileDialog.FileNames;
            fileNameTextBlock.Text = "選択中のファイル: ";
            foreach (var data in selectedFiles.Select((file, index) => new { index, file }))
            {
                if (data.index != 0)
                {
                    fileNameTextBlock.Text += ", ";
                }
                this.files.Add(data.file);
                fileNameTextBlock.Text += data.file;
            }
        }

        private void chooseOutputButton_Click(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog
            {
                DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Title = "保存先のフォルダを選択してください。",
                Multiselect = false,
            };
            if (folderDialog.ShowDialog() == false)
            {
                return;
            }
            outputDirectory = folderDialog.FolderName;
            outputDirectoryTextBox.Text = outputDirectory;
        }


        #region //Format Worksheet
        private ExcelWorksheet createFormat(ExcelPackage package, int month, int initRow)
        {
            var sheet = package.Workbook.Worksheets[0];
            sheet.Cells.Style.Font.Size = 10;
            sheet.Columns.Width = 8;
            sheet.Rows.Height = 15;

            // Leading Column
            sheet.Column(1).Width = 2.5;

            // Title
            sheet.Cells[$"C{1 + initRow}:H{1 + initRow}"].Merge = true;
            var titleCell = sheet.Cells[1 + initRow, 3];
            titleCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            titleCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            titleCell.Value = $"{month}月分支払明細書";

            // Header
            sheet.Cells[$"B{3 + initRow}:B{4 + initRow}"].Merge = true;
            var codeCell = sheet.Cells[3 + initRow, 2];
            codeCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            codeCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            codeCell.RichText.Add("支払先\r\nコード");
            sheet.Column(2).Width = 9;

            sheet.Cells[$"C{3 + initRow}:C{4 + initRow}"].Merge = true;
            var nameCell = sheet.Cells[3 + initRow, 3];
            nameCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            nameCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            nameCell.Value = "支払先名";
            sheet.Column(3).Width = 15;

            sheet.Cells[$"D{3 + initRow}:D{4 + initRow}"].Merge = true;
            var amountCell = sheet.Cells[3 + initRow, 4];
            amountCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            amountCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            amountCell.Value = "請求金額";
            sheet.Column(4).Width = 10;

            sheet.Cells[$"E{3 + initRow}:J{3 + initRow}"].Merge = true;
            var largeHeaderCell = sheet.Cells[3 + initRow, 5];
            largeHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            largeHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            largeHeaderCell.Value = "支払金額内訳";

            sheet.Cells[4 + initRow, 5].Value = "相殺";
            sheet.Cells[4 + initRow, 6].Value = "手形";
            sheet.Cells[4 + initRow, 7].Value = "期日";
            sheet.Cells[4 + initRow, 8].Value = "小切手";
            sheet.Cells[4 + initRow, 9].Value = "振込";
            sheet.Cells[4 + initRow, 10].Value = "備考";

            // Sheet format
            sheet.PrinterSettings.PaperSize = ePaperSize.A4;
            sheet.PrinterSettings.Orientation = eOrientation.Portrait;
            sheet.PrinterSettings.TopMargin = 0.75M;
            sheet.PrinterSettings.BottomMargin = 0.75M;
            sheet.PrinterSettings.LeftMargin = 0.7M;
            sheet.PrinterSettings.RightMargin = 0.7M;

            return sheet;
        }
        #endregion

        private void issueMonthTextBox_TextChanged(object sender, EventArgs e)
        {
            var currentPosition = issueMonthTextBox.SelectionStart - 1;
            var text = ((TextBox)sender).Text;

            var regex = new Regex("^[0-9]*$");

            if (!regex.IsMatch(text))
            {
                var foundChar = Regex.Match(issueMonthTextBox.Text, @"[^0-9]");
                if (foundChar.Success)
                {
                    issueMonthTextBox.Text = issueMonthTextBox.Text.Remove(foundChar.Index, 1);
                }

                issueMonthTextBox.Select(currentPosition, 0);
                return;
            }

            if (text.Length == 0)
            {
                return;
            }
            var month = int.Parse(text);
            if (month < 1 || 12 < month)
            {
                issueMonthTextBox.Text = issueMonthTextBox.Text.Remove(text.Length - 1, 1);
                issueMonthTextBox.Select(currentPosition, 0);
            }
        }
    }
}