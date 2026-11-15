using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using VFormat.Core.Pdf;

namespace VFormat.UI.Common
{
    public partial class LatPdfDialog : Window
    {
        public class PdfFileItem : INotifyPropertyChanged
        {
            public int Index { get; set; }
            public string FilePath { get; set; }
            public string FileName => Path.GetFileName(FilePath);
            public int PageCount { get; set; }
            public string PageCountText => PageCount > 0 ? $"{PageCount} trang" : "...";
            public long FileSizeBytes { get; set; }
            public string SizeText
            {
                get
                {
                    if (FileSizeBytes <= 0) return "0 KB";
                    if (FileSizeBytes < 1024 * 1024)
                        return $"{FileSizeBytes / 1024.0:F1} KB";
                    return $"{FileSizeBytes / (1024.0 * 1024.0):F2} MB";
                }
            }

            private string _status = "Sẵn sàng";
            public string Status
            {
                get => _status;
                set
                {
                    _status = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Status)));
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private readonly ObservableCollection<PdfFileItem> _files = new ObservableCollection<PdfFileItem>();
        private string _lastOutputDir = string.Empty;

        public LatPdfDialog()
        {
            InitializeComponent();
            lvFiles.ItemsSource = _files;
            AppendLog("Động cơ thuần C# OfficeFormat đã sẵn sàng. Hãy thêm tệp PDF để bắt đầu.");
        }

        private void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                string time = DateTime.Now.ToString("HH:mm:ss");
                txtLog.AppendText($"[{time}] {message}\n");
                txtLog.ScrollToEnd();
            });
        }

        private void RefreshFileCounts()
        {
            txtFileCount.Text = $" ({_files.Count} tệp)";
            for (int i = 0; i < _files.Count; i++)
            {
                _files[i].Index = i + 1;
            }
        }

        private void AddFilePaths(IEnumerable<string> paths)
        {
            foreach (var p in paths)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;

                if (File.Exists(p) && Path.GetExtension(p).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    if (_files.Any(f => f.FilePath.Equals(p, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var fi = new FileInfo(p);
                    var item = new PdfFileItem
                    {
                        Index = _files.Count + 1,
                        FilePath = p,
                        FileSizeBytes = fi.Length,
                        PageCount = 0,
                        Status = "Sẵn sàng"
                    };
                    _files.Add(item);

                    // Đọc nhanh số trang PDF
                    Task.Run(() =>
                    {
                        int pc = PdfFlipEngine.GetPageCount(p);
                        Dispatcher.Invoke(() =>
                        {
                            item.PageCount = pc;
                            lvFiles.Items.Refresh();
                        });
                    });
                }
                else if (Directory.Exists(p))
                {
                    try
                    {
                        var pdfs = Directory.GetFiles(p, "*.pdf", SearchOption.AllDirectories);
                        AddFilePaths(pdfs);
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"Không thể duyệt thư mục {p}: {ex.Message}");
                    }
                }
            }

            RefreshFileCounts();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Tệp PDF (*.pdf)|*.pdf|Mọi tệp (*.*)|*.*",
                Multiselect = true,
                Title = "Chọn các tệp PDF cần lật / sắp xếp trang"
            };

            if (ofd.ShowDialog() == true)
            {
                AddFilePaths(ofd.FileNames);
            }
        }

        private void BtnAddFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                fbd.Description = "Chọn thư mục chứa các tệp PDF:";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    AddFilePaths(new[] { fbd.SelectedPath });
                }
            }
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            var selected = lvFiles.SelectedItems.Cast<PdfFileItem>().ToList();
            foreach (var item in selected)
            {
                _files.Remove(item);
            }
            RefreshFileCounts();
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            _files.Clear();
            RefreshFileCounts();
        }

        private void LvFiles_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void LvFiles_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (dropped != null && dropped.Length > 0)
                {
                    AddFilePaths(dropped);
                }
            }
        }

        private void BtnBrowseCustomDir_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                fbd.Description = "Chọn thư mục để lưu các tệp PDF sau khi xử lý:";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtCustomDir.Text = fbd.SelectedPath;
                    radCustomDir.IsChecked = true;
                }
            }
        }

        private void BtnPageOrganizer_Click(object sender, RoutedEventArgs e)
        {
            var item = lvFiles.SelectedItem as PdfFileItem ?? _files.FirstOrDefault();
            if (item == null || !File.Exists(item.FilePath))
            {
                ModernMessageDialog.ShowWarning("Vui lòng thêm và chọn một tệp PDF trong danh sách để mở trình sắp xếp trực quan.", "THÔNG BÁO");
                return;
            }

            var organizer = new PdfPageOrganizerDialog(item.FilePath)
            {
                Owner = this
            };

            if (organizer.ShowDialog() == true && !string.IsNullOrEmpty(organizer.ExportedPath))
            {
                _lastOutputDir = Path.GetDirectoryName(organizer.ExportedPath);
                btnOpenOutDir.IsEnabled = true;
                AppendLog($"[Hoàn tất sắp xếp trực quan] Tệp đã xuất: {organizer.ExportedPath}");
            }
        }

        private async void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (_files.Count == 0)
            {
                ModernMessageDialog.ShowWarning("Vui lòng thêm ít nhất một tệp PDF vào danh sách để xử lý.", "THÔNG BÁO");
                return;
            }

            bool reverse = chkReverse.IsChecked == true;
            int rotAngle = 0;
            switch (cboRotation.SelectedIndex)
            {
                case 1: rotAngle = 90; break;
                case 2: rotAngle = 180; break;
                case 3: rotAngle = 270; break;
                default: rotAngle = 0; break;
            }

            bool inplace = radInplace.IsChecked == true;
            string outDir = null;
            if (radCustomDir.IsChecked == true)
            {
                outDir = txtCustomDir.Text.Trim();
                if (string.IsNullOrEmpty(outDir) || !Directory.Exists(outDir))
                {
                    ModernMessageDialog.ShowError("Thư mục lưu riêng không hợp lệ hoặc chưa được chọn. Vui lòng kiểm tra lại.", "LỖI");
                    return;
                }
            }

            string suffix = txtSuffix.Text.Trim();
            if (string.IsNullOrEmpty(suffix)) suffix = "_lat";

            if (inplace)
            {
                bool proceed = ModernMessageDialog.ShowConfirm(
                    "CẢNH BÁO: Bạn đang chọn ghi đè trực tiếp lên tệp PDF gốc.\nNội dung file gốc sẽ bị thay thế vĩnh viễn.\n\nBạn có chắc chắn muốn tiếp tục?",
                    "XÁC NHẬN GHI ĐÈ");

                if (!proceed)
                {
                    AppendLog("Người dùng đã hủy thao tác ghi đè trực tiếp.");
                    return;
                }
            }

            btnRun.IsEnabled = false;
            btnOpenOrganizer.IsEnabled = false;
            pbProgress.Value = 0;
            txtProgressPercent.Text = "0%";

            AppendLog($"Bắt đầu xử lý {_files.Count} tệp PDF bằng động cơ thuần C# siêu tốc...");

            int total = _files.Count;
            int successCount = 0;
            int failCount = 0;

            for (int i = 0; i < total; i++)
            {
                var fileItem = _files[i];
                fileItem.Status = "Đang xử lý...";
                lvFiles.Items.Refresh();

                AppendLog($"[{i + 1}/{total}] Đang xử lý: {fileItem.FileName}...");

                try
                {
                    var res = await Task.Run(() =>
                        PdfFlipEngine.ProcessPdfNative(
                            fileItem.FilePath,
                            outDir: outDir,
                            suffix: suffix,
                            inplace: inplace,
                            overwrite: true,
                            reverseOrder: reverse,
                            globalRotation: rotAngle));

                    if (res.Success)
                    {
                        fileItem.Status = "✓ Thành công";
                        successCount++;
                        _lastOutputDir = Path.GetDirectoryName(res.OutputPath);
                        AppendLog($"   ✓ Thành công! Lưu tại: {res.OutputPath} ({res.FinalPageCount} trang) [{res.ElapsedTime.TotalMilliseconds:F0} ms]");
                    }
                    else
                    {
                        fileItem.Status = "✗ Lỗi";
                        failCount++;
                        AppendLog($"   ✗ Thất bại: {res.ErrorMessage}");
                    }
                }
                catch (Exception ex)
                {
                    fileItem.Status = "✗ Lỗi";
                    failCount++;
                    AppendLog($"   ✗ Ngoại lệ: {ex.Message}");
                }

                int percent = (int)(((i + 1) / (double)total) * 100);
                pbProgress.Value = percent;
                txtProgressPercent.Text = $"{percent}%";
                lvFiles.Items.Refresh();
            }

            btnRun.IsEnabled = true;
            btnOpenOrganizer.IsEnabled = true;
            btnOpenOutDir.IsEnabled = !string.IsNullOrEmpty(_lastOutputDir) && Directory.Exists(_lastOutputDir);

            AppendLog($"========================================================");
            AppendLog($"HOÀN TẤT: {successCount}/{total} tệp thành công, {failCount} lỗi.");

            string summaryMsg = $"Đã hoàn tất xử lý {total} tệp PDF!\n\n" +
                $"✓ Thành công: {successCount} tệp\n" +
                $"✗ Thất bại: {failCount} tệp" +
                (!string.IsNullOrEmpty(_lastOutputDir) ? $"\n\nThư mục kết quả: {_lastOutputDir}" : "");

            if (failCount == 0)
            {
                ModernMessageDialog.ShowSuccess(summaryMsg, "HOÀN TẤT XỬ LÝ PDF");
            }
            else
            {
                ModernMessageDialog.ShowWarning(summaryMsg, "HOÀN TẤT XỬ LÝ PDF");
            }
        }

        private void BtnOpenOutDir_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_lastOutputDir) && Directory.Exists(_lastOutputDir))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = _lastOutputDir,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    AppendLog($"Không thể mở thư mục: {ex.Message}");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
