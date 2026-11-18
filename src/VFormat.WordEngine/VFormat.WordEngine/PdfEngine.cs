using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace VFormat.WordEngine;

/// <summary>
/// Động cơ chuyển đổi PDF sang Word cho V-Format.
/// Ưu tiên hàng đầu: công nghệ pdf-to-docx (thư viện pdf2docx) bảo toàn 100% bố cục, bảng biểu, ảnh.
/// Dự phòng: Microsoft Word Reflow Engine.
/// </summary>
public static class PdfEngine
{
    public static void EnsurePdfReflowAlertSuppressed()
    {
        try
        {
            string[] versions = new[] { "16.0", "15.0" };
            foreach (var ver in versions)
            {
                string keyPath = $@"Software\Microsoft\Office\{ver}\Word\Options";
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key != null)
                    {
                        key.SetValue("DisableConvertPdfWarning", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("AlertOnPDFReflow", 0, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("ConfirmConversions", 0, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                }
            }
        }
        catch
        {
        }
    }

    public static string ConvertPdfToDocx(Microsoft.Office.Interop.Word.Application wordApp, string pdfPath, string docxPath = null, bool enableOcr = false, Action<string> statusCallback = null)
    {
        if (wordApp == null) throw new ArgumentNullException(nameof(wordApp));
        if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
        {
            throw new FileNotFoundException("Không tìm thấy file PDF nguồn.", pdfPath);
        }

        if (string.IsNullOrWhiteSpace(docxPath))
        {
            docxPath = Path.Combine(Path.GetDirectoryName(pdfPath) ?? "", Path.GetFileNameWithoutExtension(pdfPath) + ".docx");
        }

        // 1. Chuyển đổi theo chuẩn skill smart-ocr (nhận diện chữ tiếng Việt AI với PaddleOCR & ONNX)
        if (SmartOcrEngine.IsSmartOcrAvailable())
        {
            statusCallback?.Invoke($"Đang chuyển đổi bằng động cơ Smart OCR: {Path.GetFileName(pdfPath)}...");
            if (SmartOcrEngine.ConvertPdfToDocx(pdfPath, docxPath, statusCallback))
            {
                statusCallback?.Invoke("Chuyển đổi PDF sang Word thành công với công nghệ Smart OCR!");
                return docxPath;
            }
        }

        // 2. Dự phòng: Word Reflow Engine nếu máy chưa cài đặt môi trường Smart OCR
        if (wordApp == null)
        {
            throw new InvalidOperationException("Không thể mở Word Reflow vì Word Application không tồn tại.");
        }

        statusCallback?.Invoke($"Đang mở và phân tích tệp PDF bằng Word Reflow: {Path.GetFileName(pdfPath)}...");

        EnsurePdfReflowAlertSuppressed();
        Document doc = null;
        bool originalScreenUpdating = true;
        try
        {
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            try { wordApp.Options.ConfirmConversions = false; } catch { }
            try { originalScreenUpdating = wordApp.ScreenUpdating; wordApp.ScreenUpdating = false; } catch { }

            object fileName = pdfPath;
            object confirmConversions = false;
            object readOnly = true;
            object addToRecentFiles = false;
            object visible = true;
            object format = WdOpenFormat.wdOpenFormatAuto;
            object missing = Missing.Value;

            doc = wordApp.Documents.Open(
                ref fileName, ref confirmConversions, ref readOnly,
                ref addToRecentFiles, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref format, ref missing,
                ref visible, ref missing, ref missing, ref missing, ref missing
            );

            statusCallback?.Invoke($"Đang lưu sang định dạng Word (.docx): {Path.GetFileName(docxPath)}...");

            object outPath = docxPath;
            object fileFormat = WdSaveFormat.wdFormatXMLDocument;

            doc.SaveAs2(
                ref outPath, ref fileFormat, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing
            );

            statusCallback?.Invoke("Chuyển đổi PDF sang Word thành công!");
            return docxPath;
        }
        finally
        {
            if (doc != null)
            {
                try
                {
                    object saveChanges = false;
                    object missing = Missing.Value;
                    doc.Close(ref saveChanges, ref missing, ref missing);
                }
                catch
                {
                }
                finally
                {
                    Marshal.ReleaseComObject(doc);
                }
            }

            try
            {
                wordApp.ScreenUpdating = originalScreenUpdating;
            }
            catch
            {
            }
            try
            {
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsAll;
            }
            catch
            {
            }
        }
    }

    public static int BatchConvert(Microsoft.Office.Interop.Word.Application wordApp, IEnumerable<string> pdfPaths, string outputDir, bool enableOcr, Action<int, int, string> progressCallback)
    {
        if (wordApp == null) throw new ArgumentNullException(nameof(wordApp));
        if (pdfPaths == null) return 0;

        var list = new List<string>(pdfPaths);
        int total = list.Count;
        int successCount = 0;

        for (int i = 0; i < total; i++)
        {
            string pdf = list[i];
            try
            {
                string targetDir = string.IsNullOrWhiteSpace(outputDir) ? (Path.GetDirectoryName(pdf) ?? "") : outputDir;
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }
                string docx = Path.Combine(targetDir, Path.GetFileNameWithoutExtension(pdf) + ".docx");
                progressCallback?.Invoke(i + 1, total, $"Đang xử lý ({i + 1}/{total}): {Path.GetFileName(pdf)}");

                ConvertPdfToDocx(wordApp, pdf, docx, enableOcr, null);
                successCount++;
            }
            catch (Exception ex)
            {
                progressCallback?.Invoke(i + 1, total, $"Lỗi khi xử lý {Path.GetFileName(pdf)}: {ex.Message}");
            }
        }

        return successCount;
    }
}
