using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace VFormat.WordEngine;

/// <summary>
/// Động cơ chuyển đổi PDF sang Word bằng trí tuệ nhân tạo Smart OCR (PaddleOCR & ONNX Engine).
/// Nhận diện chữ tiếng Việt chính xác cao từ cả PDF quét (scanned) và PDF điện tử.
/// </summary>
public static class SmartOcrEngine
{
    private static string _pythonExeCache = null;
    private static bool? _isSmartOcrAvailableCache = null;

    public static string FindPythonExe()
    {
        if (!string.IsNullOrEmpty(_pythonExeCache) && File.Exists(_pythonExeCache))
            return _pythonExeCache;

        string[] candidates = new[]
        {
            @"C:\Python314\python.exe",
            @"C:\Python313\python.exe",
            @"C:\Python312\python.exe",
            @"C:\Python311\python.exe",
            @"C:\Python310\python.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python314\python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python313\python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python312\python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python311\python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Python310\python.exe")
        };

        foreach (var c in candidates)
        {
            if (File.Exists(c))
            {
                _pythonExeCache = c;
                return c;
            }
        }

        try
        {
            var psi = new ProcessStartInfo("where", "python")
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var p = Process.Start(psi))
            {
                string outTxt = p.StandardOutput.ReadToEnd();
                p.WaitForExit(2000);
                foreach (var line in outTxt.Split('\n'))
                {
                    string tr = line.Trim();
                    if (File.Exists(tr) && !tr.Contains("WindowsApps"))
                    {
                        _pythonExeCache = tr;
                        return tr;
                    }
                }
            }
        }
        catch { }

        return "python";
    }

    public static bool IsSmartOcrAvailable()
    {
        if (_isSmartOcrAvailableCache.HasValue)
            return _isSmartOcrAvailableCache.Value;

        string py = FindPythonExe();
        try
        {
            var psi = new ProcessStartInfo(py, "-c \"import pymupdf, docx, PIL; print('OK')\"")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit(8000);
                _isSmartOcrAvailableCache = output.Contains("OK");
            }
        }
        catch
        {
            _isSmartOcrAvailableCache = false;
        }

        return _isSmartOcrAvailableCache.Value;
    }

    private static string FindConverterScriptPath()
    {
        string[] candidates = new[]
        {
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "smart_ocr_converter.py"),
            Path.Combine(Path.GetDirectoryName(typeof(SmartOcrEngine).Assembly.Location) ?? "", "smart_ocr_converter.py"),
            @"D:\PM\ADD_IN\v-format\src\VFormat.WordEngine\smart_ocr_converter.py",
            @"D:\PM\ADD_IN\v-format\src\VFormat.WordEngine\VFormat.WordEngine\smart_ocr_converter.py"
        };

        foreach (var c in candidates)
        {
            if (!string.IsNullOrEmpty(c) && File.Exists(c))
                return c;
        }
        return null;
    }

    /// <summary>
    /// Chuyển đổi PDF sang Word (.docx) bằng công nghệ Smart OCR chuyên sâu chuẩn văn bản hành chính Việt Nam.
    /// </summary>
    public static bool ConvertPdfToDocx(string pdfPath, string docxPath, Action<string> statusCallback = null)
    {
        string py = FindPythonExe();
        if (!IsSmartOcrAvailable()) return false;

        statusCallback?.Invoke("Đang khởi động động cơ nhận diện Smart OCR tiếng Việt...");

        string scriptPath = FindConverterScriptPath();
        if (string.IsNullOrEmpty(scriptPath) || !File.Exists(scriptPath))
        {
            statusCallback?.Invoke("Không tìm thấy tệp smart_ocr_converter.py!");
            return false;
        }

        try
        {
            var psi = new ProcessStartInfo(py, $"\"{scriptPath}\" \"{pdfPath}\" \"{docxPath}\"")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using (var p = Process.Start(psi))
            {
                p.OutputDataReceived += (s, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data)) return;
                    string line = e.Data;
                    if (line.StartsWith("PAGE_PROGRESS:"))
                    {
                        string progress = line.Substring("PAGE_PROGRESS:".Length);
                        statusCallback?.Invoke($"Đang nhận diện chữ tiếng Việt & cấu trúc văn bản: Trang {progress}...");
                    }
                    else if (line.StartsWith("TOTAL_PAGES:"))
                    {
                        string total = line.Substring("TOTAL_PAGES:".Length);
                        statusCallback?.Invoke($"Tệp PDF có {total} trang, đang phân tích định dạng...");
                    }
                };

                p.BeginOutputReadLine();
                string stderr = p.StandardError.ReadToEnd();
                p.WaitForExit(300000); // Tối đa 5 phút cho tài liệu lớn

                if (File.Exists(docxPath))
                {
                    statusCallback?.Invoke("Chuyển đổi PDF sang Word thành công bằng động cơ Smart OCR!");
                    return true;
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            statusCallback?.Invoke($"Lỗi Smart OCR: {ex.Message}");
            return false;
        }
    }
}
