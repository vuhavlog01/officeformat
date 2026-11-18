using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using VFormat.WordEngine;

namespace VFormat.UI.Common;

public partial class PdfConverterDialog : System.Windows.Window
{
    private readonly Microsoft.Office.Interop.Word.Application _wordApp;
    private string _lastOutputDir = "";

    public bool EnableOcr { get; set; }

    public PdfConverterDialog(Microsoft.Office.Interop.Word.Application wordApp, string initialPdfPath = null)
    {
        InitializeComponent();
        _wordApp = wordApp;
        PdfEngine.EnsurePdfReflowAlertSuppressed();

        UpdateEngineBadge();

        if (!string.IsNullOrEmpty(initialPdfPath) && File.Exists(initialPdfPath))
        {
            lstFiles.Items.Add(initialPdfPath);
        }
    }

    private void UpdateEngineBadge()
    {
        try
        {
            if (SmartOcrEngine.IsSmartOcrAvailable())
            {
                txtEngineBadge.Text = "⚡ Chuẩn thể thức văn bản & Bảng biểu";
            }
            else
            {
                txtEngineBadge.Text = "⚡ Microsoft Word Reflow";
            }
        }
        catch
        {
            txtEngineBadge.Text = "⚡ Chuẩn thể thức văn bản";
        }
    }

    private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog
        {
            Filter = "Tệp PDF (*.pdf)|*.pdf|Tất cả tệp (*.*)|*.*",
            Multiselect = true,
            Title = "Chọn tệp PDF để chuyển sang Word"
        };

        if (ofd.ShowDialog() == true)
        {
            foreach (string file in ofd.FileNames)
            {
                if (!lstFiles.Items.Contains(file))
                {
                    lstFiles.Items.Add(file);
                }
            }
        }
    }

    private void BtnClearFiles_Click(object sender, RoutedEventArgs e)
    {
        lstFiles.Items.Clear();
        txtStatus.Text = "";
        progressBar.Visibility = Visibility.Collapsed;
    }

    private void LstFiles_DragOver(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
        }
        else
        {
            e.Effects = System.Windows.DragDropEffects.None;
        }
        e.Handled = true;
    }

    private void LstFiles_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
            if (files != null)
            {
                foreach (string f in files)
                {
                    if (string.Equals(Path.GetExtension(f), ".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!lstFiles.Items.Contains(f))
                        {
                            lstFiles.Items.Add(f);
                        }
                    }
                }
            }
        }
    }

    private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
    {
        using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
        {
            fbd.Description = "Chọn thư mục lưu file Word (.docx) kết quả:";
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtCustomFolder.Text = fbd.SelectedPath;
            }
        }
    }

    private async void BtnStartConvert_Click(object sender, RoutedEventArgs e)
    {
        if (lstFiles.Items.Count == 0)
        {
            System.Windows.MessageBox.Show("Vui lòng chọn ít nhất 1 file PDF để chuyển đổi.", "Chưa chọn file", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var files = new List<string>();
        foreach (var item in lstFiles.Items)
        {
            files.Add(item.ToString());
        }

        bool openAfter = chkOpenAfterConvert.IsChecked == true;
        string customOut = rbCustomFolder.IsChecked == true ? (txtCustomFolder.Text ?? "").Trim() : "";

        btnStartConvert.IsEnabled = false;
        progressBar.Visibility = Visibility.Visible;
        progressBar.IsIndeterminate = true;

        string firstConvertedDocx = null;
        int successCount = 0;

        for (int i = 0; i < files.Count; i++)
        {
            string pdf = files[i];
            string outDir = string.IsNullOrEmpty(customOut) ? (Path.GetDirectoryName(pdf) ?? "") : customOut;
            _lastOutputDir = outDir;
            string docx = Path.Combine(outDir, Path.GetFileNameWithoutExtension(pdf) + ".docx");

            txtStatus.Text = $"Đang xử lý ({i + 1}/{files.Count}): {Path.GetFileName(pdf)}...";

            try
            {
                // 1. Nhận diện chữ và chuyển đổi sang Word bằng động cơ Smart OCR (PaddleOCR & ONNX)
                if (SmartOcrEngine.IsSmartOcrAvailable())
                {
                    Dispatcher.Invoke(() => txtStatus.Text = $"Đang nhận diện chữ và chuyển đổi bằng Smart OCR: {Path.GetFileName(pdf)}...");
                    bool converted = await System.Threading.Tasks.Task.Run(() =>
                    {
                        return SmartOcrEngine.ConvertPdfToDocx(pdf, docx, msg =>
                        {
                            Dispatcher.Invoke(() => txtStatus.Text = msg);
                        });
                    });

                    if (converted && File.Exists(docx))
                    {
                        successCount++;
                        if (firstConvertedDocx == null) firstConvertedDocx = docx;
                        continue;
                    }
                }

                // 2. Dự phòng: Dùng Word Reflow trên STA thread nếu máy chưa cài Smart OCR
                Dispatcher.Invoke(() => txtStatus.Text = $"Đang chuyển đổi bằng Word Reflow: {Path.GetFileName(pdf)}...");
                PdfEngine.ConvertPdfToDocx(_wordApp, pdf, docx, false, msg => txtStatus.Text = msg);
                successCount++;
                if (firstConvertedDocx == null) firstConvertedDocx = docx;
            }
            catch (Exception ex)
            {
                txtStatus.Text = $"Lỗi {Path.GetFileName(pdf)}: {ex.Message}";
            }
        }

        progressBar.Visibility = Visibility.Collapsed;
        btnStartConvert.IsEnabled = true;
        txtStatus.Text = $"Hoàn tất chuyển đổi {successCount}/{files.Count} tệp sang Word (.docx)!";

        if (successCount > 0 && openAfter && !string.IsNullOrEmpty(firstConvertedDocx) && File.Exists(firstConvertedDocx))
        {
            try
            {
                _wordApp.Documents.Open(firstConvertedDocx);
            }
            catch
            {
                Process.Start(new ProcessStartInfo(firstConvertedDocx) { UseShellExecute = true });
            }
        }

        System.Windows.MessageBox.Show($"Đã chuyển đổi thành công {successCount}/{files.Count} file PDF sang Word (.docx) bằng công nghệ Smart OCR!", "Chuyển đổi hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        string dir = _lastOutputDir;
        if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
        {
            if (lstFiles.Items.Count > 0)
            {
                dir = Path.GetDirectoryName(lstFiles.Items[0].ToString());
            }
        }

        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
        {
            Process.Start("explorer.exe", dir);
        }
        else
        {
            System.Windows.MessageBox.Show("Chưa có thư mục kết quả để mở.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
