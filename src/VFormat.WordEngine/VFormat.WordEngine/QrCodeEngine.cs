using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine;

public enum QrInsertLocation
{
    Cursor,
    HeaderTopRight,
    DocumentEnd
}

public enum QrFormatMode
{
    WordNativeField,
    ImageFile
}

public class DocumentMetadata
{
    public string Agency { get; set; } = "";
    public string DocumentNumber { get; set; } = "";
    public string DateText { get; set; } = "";
    public string Excerpt { get; set; } = "";
    public string Signer { get; set; } = "";
    public string LookupUrl { get; set; } = "";
}

public static class QrCodeEngine
{
    public static DocumentMetadata ExtractMetadata(Document doc)
    {
        var meta = new DocumentMetadata();
        if (doc == null) return meta;

        try
        {
            // 1. Quét từ bảng đầu tiên (Bảng Quốc hiệu / Tên cơ quan / Số hiệu / Ngày tháng)
            if (doc.Tables.Count > 0)
            {
                Table headerTable = doc.Tables[1];
                try
                {
                    int rowCount = headerTable.Rows.Count;
                    int colCount = headerTable.Columns.Count;
                    if (rowCount >= 1 && colCount >= 2)
                    {
                        // Cột 1 (thường là Cơ quan ban hành, Số ký hiệu)
                        string c1Text = CleanCellText(headerTable.Cell(1, 1).Range.Text);
                        var c1Lines = c1Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        if (c1Lines.Length > 0)
                        {
                            meta.Agency = c1Lines[0].Trim();
                            for (int i = 1; i < c1Lines.Length; i++)
                            {
                                string line = c1Lines[i].Trim();
                                if (line.StartsWith("Số:", StringComparison.OrdinalIgnoreCase) || line.StartsWith("Số", StringComparison.OrdinalIgnoreCase))
                                {
                                    meta.DocumentNumber = line;
                                }
                            }
                        }

                        // Cột 2 (Quốc hiệu, Địa danh ngày tháng)
                        string c2Text = CleanCellText(headerTable.Cell(1, 2).Range.Text);
                        var dateMatch = Regex.Match(c2Text, @"(?:ngày|tháng|năm|\d{1,2}\s*tháng\s*\d{1,2}\s*năm\s*\d{4})", RegexOptions.IgnoreCase);
                        if (dateMatch.Success)
                        {
                            var c2Lines = c2Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var line in c2Lines)
                            {
                                if (Regex.IsMatch(line, @"\bngày\b.*\bnăm\b", RegexOptions.IgnoreCase))
                                {
                                    meta.DateText = line.Trim();
                                    break;
                                }
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(headerTable);
                }
            }

            // 2. Quét tiêu đề văn bản / Trích yếu nội dung từ 25 đoạn đầu
            int scanMax = Math.Min(doc.Paragraphs.Count, 25);
            for (int i = 1; i <= scanMax; i++)
            {
                Paragraph p = doc.Paragraphs[i];
                try
                {
                    if (p.Range.Tables.Count > 0) continue;
                    string t = (p.Range.Text ?? "").Trim();
                    if (string.IsNullOrEmpty(t)) continue;

                    // Nếu chưa có Số/KH
                    if (string.IsNullOrEmpty(meta.DocumentNumber) && (t.StartsWith("Số:", StringComparison.OrdinalIgnoreCase) || t.StartsWith("Số ", StringComparison.OrdinalIgnoreCase)))
                    {
                        meta.DocumentNumber = t;
                        continue;
                    }

                    // Nếu chưa có ngày tháng
                    if (string.IsNullOrEmpty(meta.DateText) && Regex.IsMatch(t, @"\bngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}\b", RegexOptions.IgnoreCase))
                    {
                        meta.DateText = t;
                        continue;
                    }

                    // Nhận diện Tên loại văn bản / Trích yếu
                    if (string.IsNullOrEmpty(meta.Excerpt))
                    {
                        if (Regex.IsMatch(t, @"^(QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|CÔNG VĂN|CHỈ THỊ|BÁO CÁO|NGHỊ QUYẾT|BIÊN BẢN)", RegexOptions.IgnoreCase))
                        {
                            // Lấy đoạn này và đoạn trích yếu tiếp theo nếu có
                            meta.Excerpt = t;
                            if (i + 1 <= scanMax)
                            {
                                string nextT = (doc.Paragraphs[i + 1].Range.Text ?? "").Trim();
                                if (!string.IsNullOrEmpty(nextT) && !nextT.StartsWith("Điều", StringComparison.OrdinalIgnoreCase) && nextT.Length > 5)
                                {
                                    meta.Excerpt += " - " + nextT;
                                }
                            }
                            break;
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(p);
                }
            }

            // 3. Quét người ký từ bảng cuối (Bảng ký tên)
            if (doc.Tables.Count > 0)
            {
                Table lastTable = doc.Tables[doc.Tables.Count];
                try
                {
                    int lastRow = lastTable.Rows.Count;
                    int lastCol = lastTable.Columns.Count;
                    if (lastRow >= 1 && lastCol >= 2)
                    {
                        string signerCell = CleanCellText(lastTable.Cell(lastRow, lastCol).Range.Text);
                        var sLines = signerCell.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        if (sLines.Length > 0)
                        {
                            meta.Signer = sLines[sLines.Length - 1].Trim();
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(lastTable);
                }
            }
        }
        catch
        {
        }

        return meta;
    }

    public static void InsertQrCode(Document doc, string qrContent, QrInsertLocation location, QrFormatMode mode, float sizeCm = 2.5f)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (string.IsNullOrWhiteSpace(qrContent)) throw new ArgumentException("Nội dung mã QR không được để trống.", nameof(qrContent));

        float points = sizeCm * 28.3465f; // 1 cm = 28.3465 pt

        using (new ScreenFreezeScope(doc.Application))
        {
            using (new SafeUndoScope(doc.Application, "OfficeFormat Tạo Mã QR"))
            {
                Range targetRange = null;

                switch (location)
                {
                    case QrInsertLocation.Cursor:
                        targetRange = doc.Application?.Selection?.Range ?? doc.Content;
                        break;

                    case QrInsertLocation.DocumentEnd:
                        object endPos = Math.Max(0, doc.Content.End - 1);
                        targetRange = doc.Range(ref endPos, ref endPos);
                        targetRange.InsertParagraphAfter();
                        endPos = Math.Max(0, doc.Content.End - 1);
                        targetRange = doc.Range(ref endPos, ref endPos);
                        break;

                    case QrInsertLocation.HeaderTopRight:
                        // Chèn vào góc phải trên trang 1 (khung Shape nổi hoặc đầu tài liệu)
                        InsertTopRightQr(doc, qrContent, mode, points);
                        return;
                }

                if (targetRange == null)
                {
                    targetRange = doc.Content;
                }

                if (mode == QrFormatMode.WordNativeField)
                {
                    // Chèn trường DISPLAYBARCODE trực tiếp của Word
                    InsertNativeFieldQr(doc, targetRange, qrContent, (int)(sizeCm * 40));
                }
                else
                {
                    // Chèn ảnh PNG chất lượng cao
                    InsertImageQr(doc, targetRange, qrContent, points, points);
                }
            }
        }
    }

    private static void InsertTopRightQr(Document doc, string qrContent, QrFormatMode mode, float points)
    {
        string tempPng = Path.Combine(Path.GetTempPath(), $"OF_QR_{Guid.NewGuid():N}.png");
        try
        {
            QrCodeHelper.SaveQrPng(qrContent, tempPng, 12);
            object linkToFile = false;
            object saveWithDocument = true;
            object rangeObj = doc.Paragraphs[1].Range;

            Shape shape = doc.Shapes.AddPicture(
                tempPng, ref linkToFile, ref saveWithDocument,
                Left: doc.PageSetup.PageWidth - doc.PageSetup.RightMargin - points,
                Top: doc.PageSetup.TopMargin,
                Width: points,
                Height: points,
                Anchor: ref rangeObj
            );

            try
            {
                shape.WrapFormat.Type = WdWrapType.wdWrapSquare;
            }
            finally
            {
                Marshal.ReleaseComObject(shape);
            }
        }
        finally
        {
            try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
        }
    }

    private static void InsertNativeFieldQr(Document doc, Range range, string qrContent, int scale = 100)
    {
        // Word DISPLAYBARCODE format: DISPLAYBARCODE "text" QR \q 3 \s 100
        string cleanContent = qrContent.Replace("\"", "\\\"").Replace("\r\n", "\n").Replace("\r", "\n");
        object fieldText = $"DISPLAYBARCODE \"{cleanContent}\" QR \\q 3 \\s {Math.Max(50, Math.Min(300, scale))}";
        object fieldType = WdFieldType.wdFieldEmpty;
        object preserveFormatting = true;
        Field fld = doc.Fields.Add(range, ref fieldType, ref fieldText, ref preserveFormatting);
        try
        {
            fld.Update();
        }
        finally
        {
            Marshal.ReleaseComObject(fld);
        }
    }

    private static void InsertImageQr(Document doc, Range range, string qrContent, float widthPt, float heightPt)
    {
        string tempPng = Path.Combine(Path.GetTempPath(), $"OF_QR_{Guid.NewGuid():N}.png");
        try
        {
            QrCodeHelper.SaveQrPng(qrContent, tempPng, 12);
            object linkToFile = false;
            object saveWithDoc = true;
            object rangeObj = range;
            InlineShape inlineShape = range.InlineShapes.AddPicture(tempPng, ref linkToFile, ref saveWithDoc, ref rangeObj);
            try
            {
                inlineShape.Width = widthPt;
                inlineShape.Height = heightPt;
            }
            finally
            {
                Marshal.ReleaseComObject(inlineShape);
            }
        }
        finally
        {
            try { if (File.Exists(tempPng)) File.Delete(tempPng); } catch { }
        }
    }

    private static string CleanCellText(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        return text.Replace("\a", "").Replace("\r\a", "").Trim();
    }
}
