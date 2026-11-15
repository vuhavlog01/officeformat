using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace VFormat.Core.Pdf
{
    /// <summary>
    /// Lõi xử lý đảo ngược thứ tự trang, sắp xếp và xoay trang PDF thuần C# (Native PDF Engine).
    /// Dịch và mở rộng từ lõi xử lý latpdf/core.py.
    /// </summary>
    public static class PdfFlipEngine
    {
        /// <summary>
        /// Tạo đường dẫn file kết quả không bị trùng bằng cách thêm (1), (2), ... nếu đã tồn tại.
        /// </summary>
        public static string UniquePath(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return path;

            string dir = Path.GetDirectoryName(path) ?? string.Empty;
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            int counter = 1;
            while (true)
            {
                string candidate = Path.Combine(dir, $"{name} ({counter}){ext}");
                if (!File.Exists(candidate))
                    return candidate;
                counter++;
            }
        }

        /// <summary>
        /// Lấy nhanh tổng số trang của file PDF.
        /// </summary>
        public static int GetPageCount(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return 0;

            try
            {
                using (var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly))
                {
                    return doc.PageCount;
                }
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Phân tích chuỗi thứ tự trang (1-based) thành danh sách chỉ số trang 0-based.
        /// Hỗ trợ:
        /// - Danh sách rời rạc: "1, 3, 2" hoặc "1 3 2"
        /// - Dải trang tăng hoặc giảm: "1-3", "5-1"
        /// - Kết hợp: "1, 4-2, 5"
        /// </summary>
        public static List<int> ParsePageOrder(string spec, int totalPages)
        {
            if (string.IsNullOrWhiteSpace(spec))
                throw new ArgumentException("Chưa nhập thứ tự trang.");

            if (totalPages <= 0)
                throw new ArgumentException("Tổng số trang không hợp lệ.");

            string raw = spec.Trim().Replace(";", ",").Replace(" ", ",");
            var tokens = raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(t => t.Trim())
                            .Where(t => !string.IsNullOrEmpty(t))
                            .ToList();

            if (tokens.Count == 0)
                throw new ArgumentException("Thứ tự trang không hợp lệ.");

            var indices = new List<int>();
            foreach (var tok in tokens)
            {
                if (tok.Contains("-"))
                {
                    var parts = tok.Split('-');
                    if (parts.Length != 2 || !int.TryParse(parts[0], out int startP) || !int.TryParse(parts[1], out int endP))
                        throw new ArgumentException($"Dải trang không hợp lệ: '{tok}'");

                    if (startP < 1 || startP > totalPages || endP < 1 || endP > totalPages)
                        throw new ArgumentException($"Số trang trong '{tok}' vượt quá tổng số trang ({totalPages}).");

                    int step = (startP <= endP) ? 1 : -1;
                    for (int p = startP; step > 0 ? p <= endP : p >= endP; p += step)
                    {
                        indices.Add(p - 1);
                    }
                }
                else
                {
                    if (!int.TryParse(tok, out int p))
                        throw new ArgumentException($"Số trang không hợp lệ: '{tok}'");

                    if (p < 1 || p > totalPages)
                        throw new ArgumentException($"Trang {p} không tồn tại (file chỉ có {totalPages} trang).");

                    indices.Add(p - 1);
                }
            }
            return indices;
        }

        /// <summary>
        /// Xử lý đảo ngược, sắp xếp lại hoặc xoay trang file PDF bằng động cơ thuần C#.
        /// </summary>
        public static PdfProcessResult ProcessPdfNative(
            string srcPath,
            string outDir = null,
            string suffix = "_lat",
            bool inplace = false,
            bool overwrite = false,
            bool reverseOrder = true,
            List<int> pageOrder = null,
            Dictionary<int, int> manualRotations = null,
            Action<int, int> onPageProgress = null)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var result = new PdfProcessResult
            {
                SourcePath = srcPath
            };

            if (string.IsNullOrEmpty(srcPath) || !File.Exists(srcPath))
            {
                result.Success = false;
                result.ErrorMessage = "Tệp PDF nguồn không tồn tại.";
                return result;
            }

            try
            {
                // Xác định đường dẫn file đầu ra
                string targetDir = string.IsNullOrEmpty(outDir) ? (Path.GetDirectoryName(srcPath) ?? "") : outDir;
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                string srcName = Path.GetFileNameWithoutExtension(srcPath);
                string ext = Path.GetExtension(srcPath);

                string outPath;
                if (inplace)
                {
                    outPath = srcPath;
                }
                else
                {
                    string baseOut = Path.Combine(targetDir, $"{srcName}{suffix}{ext}");
                    outPath = overwrite ? baseOut : UniquePath(baseOut);
                }

                result.OutputPath = outPath;

                // Mở file nguồn
                using (var inputDoc = PdfReader.Open(srcPath, PdfDocumentOpenMode.Import))
                {
                    int totalPages = inputDoc.PageCount;
                    result.OriginalPageCount = totalPages;

                    if (totalPages == 0)
                    {
                        result.Success = false;
                        result.ErrorMessage = "Tệp PDF không có trang nào.";
                        return result;
                    }

                    // Xác định thứ tự trang
                    List<int> targetOrder;
                    if (pageOrder != null && pageOrder.Count > 0)
                    {
                        targetOrder = pageOrder;
                    }
                    else if (reverseOrder)
                    {
                        targetOrder = Enumerable.Range(0, totalPages).Reverse().ToList();
                    }
                    else
                    {
                        targetOrder = Enumerable.Range(0, totalPages).ToList();
                    }

                    using (var outputDoc = new PdfDocument())
                    {
                        int rotCount = 0;
                        int processed = 0;

                        for (int i = 0; i < targetOrder.Count; i++)
                        {
                            int origIdx = targetOrder[i];
                            if (origIdx < 0 || origIdx >= totalPages)
                                continue;

                            var page = inputDoc.Pages[origIdx];

                            // Áp dụng góc xoay bổ sung nếu có
                            if (manualRotations != null && manualRotations.TryGetValue(origIdx, out int addAngle))
                            {
                                int curAngle = page.Rotate;
                                page.Rotate = (curAngle + addAngle) % 360;
                                if (addAngle % 360 != 0)
                                {
                                    rotCount++;
                                }
                            }

                            outputDoc.AddPage(page);
                            processed++;
                            onPageProgress?.Invoke(processed, targetOrder.Count);
                        }

                        // Ghi ra file tạm an toàn
                        string tempFile = Path.Combine(Path.GetTempPath(), $"of_{Guid.NewGuid():N}.pdf");
                        try
                        {
                            outputDoc.Save(tempFile);

                            if (File.Exists(outPath))
                            {
                                File.Delete(outPath);
                            }
                            File.Move(tempFile, outPath);
                        }
                        finally
                        {
                            if (File.Exists(tempFile))
                            {
                                try { File.Delete(tempFile); } catch { }
                            }
                        }

                        result.FinalPageCount = outputDoc.PageCount;
                        result.RotatedPagesCount = rotCount;
                        result.Success = true;
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }

            sw.Stop();
            result.ElapsedTime = sw.Elapsed;
            return result;
        }
    }
}
