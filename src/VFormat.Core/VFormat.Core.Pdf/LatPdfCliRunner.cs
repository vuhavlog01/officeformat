using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace VFormat.Core.Pdf
{
    /// <summary>
    /// Điều phối và thực thi công cụ LatPdf (Reverse Engineer từ LatPdf.exe)
    /// Hỗ trợ AI RapidOCR phát hiện hướng chữ, xóa trang trắng scan và sắp xếp hàng loạt.
    /// </summary>
    public static class LatPdfCliRunner
    {
        private static string _customExecutablePath;

        public static void SetExecutablePath(string path)
        {
            _customExecutablePath = path;
        }

        /// <summary>
        /// Tìm kiếm vị trí thực thi của LatPdf.exe.
        /// </summary>
        public static string FindLatPdfExecutable()
        {
            if (!string.IsNullOrEmpty(_customExecutablePath) && File.Exists(_customExecutablePath))
                return _customExecutablePath;

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string localTools = Path.Combine(baseDir, "tools", "LatPdf.exe");
            if (File.Exists(localTools))
                return localTools;

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appDataTools = Path.Combine(appData, "OfficeFormat", "tools", "LatPdf.exe");
            if (File.Exists(appDataTools))
                return appDataTools;

            // Kiểm tra thư mục Downloads thường dùng của tác giả
            string downloadPath = @"D:\Downloads\LatPdf.exe";
            if (File.Exists(downloadPath))
                return downloadPath;

            string userDownloads = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "LatPdf.exe");
            if (File.Exists(userDownloads))
                return userDownloads;

            return null;
        }

        /// <summary>
        /// Kiểm tra tính sẵn sàng của công nghệ AI LatPdf (RapidOCR / OpenCV).
        /// </summary>
        public static bool IsLatPdfAvailable()
        {
            return !string.IsNullOrEmpty(FindLatPdfExecutable());
        }

        /// <summary>
        /// Chạy tiến trình LatPdf CLI để xử lý một hoặc nhiều file PDF.
        /// </summary>
        public static async Task<PdfProcessResult> RunAsync(
            string srcFile,
            string outDir = null,
            string suffix = "_lat",
            bool inplace = false,
            bool overwrite = false,
            bool reverseOrder = true,
            bool autoOrient = true,
            bool removeBlank = true,
            string pageOrderSpec = null,
            Action<string> logCallback = null)
        {
            var sw = Stopwatch.StartNew();
            var res = new PdfProcessResult { SourcePath = srcFile };

            string exePath = FindLatPdfExecutable();
            if (string.IsNullOrEmpty(exePath))
            {
                // Fallback về động cơ thuần C# Native PDF Engine
                logCallback?.Invoke("[Thông báo] Không tìm thấy LatPdf.exe AI. Đang sử dụng Động cơ thuần C# (Native C# PDF Engine)...");
                List<int> pageOrder = null;
                if (!string.IsNullOrEmpty(pageOrderSpec))
                {
                    int totalP = PdfFlipEngine.GetPageCount(srcFile);
                    try { pageOrder = PdfFlipEngine.ParsePageOrder(pageOrderSpec, totalP); } catch { }
                }

                return await Task.Run(() =>
                    PdfFlipEngine.ProcessPdfNative(
                        srcFile,
                        outDir: outDir,
                        suffix: suffix,
                        inplace: inplace,
                        overwrite: overwrite,
                        reverseOrder: reverseOrder,
                        pageOrder: pageOrder));
            }

            // Chuẩn bị tham số dòng lệnh cho LatPdf.exe
            var args = new StringBuilder();
            args.Append($"\"{srcFile}\"");

            if (!string.IsNullOrEmpty(outDir))
            {
                args.Append($" -o \"{outDir}\"");
            }

            if (!string.IsNullOrEmpty(suffix))
            {
                args.Append($" -s \"{suffix}\"");
            }

            if (inplace)
            {
                args.Append(" --inplace -y");
            }

            if (overwrite)
            {
                args.Append(" --overwrite");
            }

            if (!reverseOrder && string.IsNullOrEmpty(pageOrderSpec))
            {
                args.Append(" --no-reverse");
            }

            if (autoOrient)
            {
                args.Append(" --auto-orient");
            }

            if (removeBlank)
            {
                args.Append(" --remove-blank");
            }

            if (!string.IsNullOrEmpty(pageOrderSpec))
            {
                args.Append($" --order \"{pageOrderSpec}\"");
            }

            logCallback?.Invoke($"[Khởi chạy LatPdf] {Path.GetFileName(srcFile)}...");

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = args.ToString(),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                using (var process = new Process { StartInfo = psi })
                {
                    process.OutputDataReceived += (s, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            logCallback?.Invoke(e.Data);
                        }
                    };

                    process.ErrorDataReceived += (s, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            logCallback?.Invoke($"[Lỗi/Cảnh báo] {e.Data}");
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    await Task.Run(() => process.WaitForExit());

                    res.Success = (process.ExitCode == 0);
                    if (!res.Success)
                    {
                        res.ErrorMessage = $"Tiến trình kết thúc với mã {process.ExitCode}";
                    }
                }

                // Xác định file kết quả
                string targetDir = string.IsNullOrEmpty(outDir) ? (Path.GetDirectoryName(srcFile) ?? "") : outDir;
                string srcName = Path.GetFileNameWithoutExtension(srcFile);
                string ext = Path.GetExtension(srcFile);
                string expectedOut = inplace ? srcFile : Path.Combine(targetDir, $"{srcName}{suffix}{ext}");
                
                if (File.Exists(expectedOut))
                {
                    res.OutputPath = expectedOut;
                    res.FinalPageCount = PdfFlipEngine.GetPageCount(expectedOut);
                    res.OriginalPageCount = PdfFlipEngine.GetPageCount(srcFile);
                    res.Success = true;
                }
            }
            catch (Exception ex)
            {
                res.Success = false;
                res.ErrorMessage = ex.Message;
                logCallback?.Invoke($"[Lỗi ngoại lệ] {ex.Message}");
            }

            sw.Stop();
            res.ElapsedTime = sw.Elapsed;
            return res;
        }
    }
}
