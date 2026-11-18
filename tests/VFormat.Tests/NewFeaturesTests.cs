using System;
using System.IO;
using System.Collections.Generic;
using Xunit;
using VFormat.Core.Common;
using VFormat.WordEngine;

namespace VFormat.Tests;

public class NewFeaturesTests
{
    [Fact]
    public void QrCodeHelper_GenerateQrPngBytes_ValidContent_ReturnsNonEmptyBytes()
    {
        byte[] bytes = QrCodeHelper.GenerateQrPngBytes("https://dichvucong.gov.vn", 5);
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 50);

        // Check PNG signature: 137, 80, 78, 71, 13, 10, 26, 10
        Assert.Equal(0x89, bytes[0]);
        Assert.Equal(0x50, bytes[1]);
        Assert.Equal(0x4E, bytes[2]);
        Assert.Equal(0x47, bytes[3]);
    }

    [Fact]
    public void QrCodeHelper_BuildDocumentMetadataString_FormatsCorrectly()
    {
        string meta = QrCodeHelper.BuildDocumentMetadataString(
            "UBND XÃ QUẾ SƠN",
            "Số: 15/KH-UBND",
            "Quế Sơn, ngày 18 tháng 8 năm 2026",
            "Kế hoạch Chuyển đổi số giai đoạn 2026 - 2030",
            "CHỦ TỊCH - Nguyễn Văn A",
            "https://queson.gov.vn/vb/15"
        );

        Assert.Contains("CQ: UBND XÃ QUẾ SƠN", meta);
        Assert.Contains("SỐ: Số: 15/KH-UBND", meta);
        Assert.Contains("TRÍCH YẾU: Kế hoạch Chuyển đổi số", meta);
        Assert.Contains("NGƯỜI KÝ: CHỦ TỊCH - Nguyễn Văn A", meta);
        Assert.Contains("TRA CỨU: https://queson.gov.vn/vb/15", meta);
    }

    [Fact]
    public void QrCodeHelper_BuildVietQrEmvCoPayload_GeneratesValidFormat()
    {
        string payload = QrCodeHelper.BuildVietQrEmvCoPayload(
            "970436", // Vietcombank
            "0011001234567",
            500000m,
            "Nop le phi cong chung",
            "NGUYEN VAN A"
        );

        Assert.NotNull(payload);
        Assert.StartsWith("000201", payload); // Tag 00 length 02 value 01
        Assert.Contains("A000000727", payload); // NAPAS GUID
        Assert.Contains("970436", payload); // Bank BIN
        Assert.Contains("0011001234567", payload); // Account number
        Assert.Contains("5406500000", payload); // Amount 500000
        Assert.Contains("5802VN", payload); // Country VN
        Assert.Contains("5303704", payload); // Currency VND (704)
        Assert.EndsWith(payload.Substring(payload.Length - 4), payload); // Has 4-char CRC
    }

    [Fact]
    public void PrintEngine_GetInstalledPrinters_ReturnsList()
    {
        var printers = PrintEngine.GetInstalledPrinters(out string defaultPrinter);
        Assert.NotNull(printers);
    }

    [Fact]
    public void PrintEngine_PrintPages_NullDoc_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() =>
        {
            PrintEngine.PrintPages(null, Microsoft.Office.Interop.Word.WdPrintOutPages.wdPrintOddPagesOnly);
        });
    }

    [Fact]
    public void PdfEngine_ConvertPdfToDocx_NullApp_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() =>
        {
            PdfEngine.ConvertPdfToDocx(null, "test.pdf");
        });
    }

    [Fact]
    public void PdfEngine_BatchConvert_NullApp_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() =>
        {
            PdfEngine.BatchConvert(null, new[] { "test.pdf" }, "out", false, null);
        });
    }

    [Fact]
    public void PdfEngine_EnsurePdfReflowAlertSuppressed_ExecutesAndSetsRegistry()
    {
        PdfEngine.EnsurePdfReflowAlertSuppressed();

        using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Word\Options"))
        {
            if (key != null)
            {
                object val = key.GetValue("DisableConvertPdfWarning");
                Assert.NotNull(val);
                Assert.Equal(1, Convert.ToInt32(val));

                object reflow = key.GetValue("AlertOnPDFReflow");
                Assert.NotNull(reflow);
                Assert.Equal(0, Convert.ToInt32(reflow));
            }
        }
    }

    [Fact]
    public void SmartOcrEngine_Detection_WorksProperly()
    {
        string py = SmartOcrEngine.FindPythonExe();
        Assert.False(string.IsNullOrEmpty(py));
        bool available = SmartOcrEngine.IsSmartOcrAvailable();
        Assert.True(available, "Smart OCR (PaddleOCR & ONNX) should be available in Python environment");
    }
}
