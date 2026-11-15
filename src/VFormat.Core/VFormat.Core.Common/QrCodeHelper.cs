using System;
using System.IO;
using System.Text;
using QRCoder;

namespace VFormat.Core.Common;

public static class QrCodeHelper
{
    public static byte[] GenerateQrPngBytes(string content, int pixelsPerModule = 10, QRCodeGenerator.ECCLevel ecc = QRCodeGenerator.ECCLevel.M)
    {
        if (string.IsNullOrEmpty(content))
        {
            content = " ";
        }

        using (var qrGenerator = new QRCodeGenerator())
        {
            using (var qrCodeData = qrGenerator.CreateQrCode(content, ecc))
            {
                var pngQrCode = new PngByteQRCode(qrCodeData);
                return pngQrCode.GetGraphic(pixelsPerModule);
            }
        }
    }

    public static void SaveQrPng(string content, string filePath, int pixelsPerModule = 10, QRCodeGenerator.ECCLevel ecc = QRCodeGenerator.ECCLevel.M)
    {
        byte[] bytes = GenerateQrPngBytes(content, pixelsPerModule, ecc);
        string dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }
        File.WriteAllBytes(filePath, bytes);
    }

    public static string BuildDocumentMetadataString(string agency, string docNumber, string dateText, string excerpt, string signer, string lookupUrl = null)
    {
        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(agency))
        {
            sb.AppendLine($"CQ: {agency.Trim()}");
        }
        if (!string.IsNullOrWhiteSpace(docNumber))
        {
            sb.AppendLine($"SỐ: {docNumber.Trim()}");
        }
        if (!string.IsNullOrWhiteSpace(dateText))
        {
            sb.AppendLine($"NGÀY: {dateText.Trim()}");
        }
        if (!string.IsNullOrWhiteSpace(excerpt))
        {
            sb.AppendLine($"TRÍCH YẾU: {excerpt.Trim()}");
        }
        if (!string.IsNullOrWhiteSpace(signer))
        {
            sb.AppendLine($"NGƯỜI KÝ: {signer.Trim()}");
        }
        if (!string.IsNullOrWhiteSpace(lookupUrl))
        {
            sb.AppendLine($"TRA CỨU: {lookupUrl.Trim()}");
        }
        return sb.ToString().TrimEnd();
    }

    /// <summary>
    /// Tạo payload mã VietQR chuẩn EMVCo (NAPAS)
    /// </summary>
    public static string BuildVietQrEmvCoPayload(string bankBin, string accountNo, decimal? amount, string memo, string accountName)
    {
        if (string.IsNullOrWhiteSpace(bankBin) || string.IsNullOrWhiteSpace(accountNo))
        {
            return "";
        }

        bankBin = bankBin.Trim();
        accountNo = accountNo.Trim();

        // Sub-elements của Tag 38 (Merchant Account Info - VietQR NAPAS)
        // 00: GUID A000000727 (NAPAS)
        // 01: Beneficiary Org & Account
        string napasGuid = FormatEmvCoTag("00", "A000000727");
        string benSub00 = FormatEmvCoTag("00", bankBin);
        string benSub01 = FormatEmvCoTag("01", accountNo);
        string benAccountInfo = FormatEmvCoTag("01", benSub00 + benSub01);
        string serviceCode = FormatEmvCoTag("02", "QRIBFTTA"); // Quick response interbank fund transfer to account
        string tag38Value = napasGuid + benAccountInfo + serviceCode;
        string tag38 = FormatEmvCoTag("38", tag38Value);

        // Tag 00: Payload Format Indicator (01)
        string tag00 = FormatEmvCoTag("00", "01");
        // Tag 01: Point of Initiation Method (11: Static, 12: Dynamic)
        string tag01 = FormatEmvCoTag("01", amount.HasValue && amount.Value > 0 ? "12" : "11");
        // Tag 53: Transaction Currency (704 = VND)
        string tag53 = FormatEmvCoTag("53", "704");
        // Tag 54: Transaction Amount
        string tag54 = "";
        if (amount.HasValue && amount.Value > 0)
        {
            tag54 = FormatEmvCoTag("54", Math.Round(amount.Value, 0).ToString("F0"));
        }
        // Tag 58: Country Code (VN)
        string tag58 = FormatEmvCoTag("58", "VN");
        // Tag 59: Merchant Name
        string tag59 = "";
        if (!string.IsNullOrWhiteSpace(accountName))
        {
            tag59 = FormatEmvCoTag("59", RemoveAccents(accountName.Trim().ToUpperInvariant()));
        }
        // Tag 60: Merchant City
        string tag60 = FormatEmvCoTag("60", "VIETNAM");
        // Tag 62: Additional Data Field (Memo)
        string tag62 = "";
        if (!string.IsNullOrWhiteSpace(memo))
        {
            string memoSub = FormatEmvCoTag("08", RemoveAccents(memo.Trim()));
            tag62 = FormatEmvCoTag("62", memoSub);
        }

        string rawPayloadWithoutCrc = tag00 + tag01 + tag38 + tag53 + tag54 + tag58 + tag59 + tag60 + tag62 + "6304";
        string crc = ComputeCrc16Ccitt(rawPayloadWithoutCrc);
        return rawPayloadWithoutCrc + crc;
    }

    private static string FormatEmvCoTag(string tag, string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        string len = value.Length.ToString("D2");
        return tag + len + value;
    }

    private static string ComputeCrc16Ccitt(string data)
    {
        byte[] bytes = Encoding.ASCII.GetBytes(data);
        ushort crc = 0xFFFF;
        for (int i = 0; i < bytes.Length; i++)
        {
            crc ^= (ushort)(bytes[i] << 8);
            for (int j = 0; j < 8; j++)
            {
                if ((crc & 0x8000) != 0)
                {
                    crc = (ushort)((crc << 1) ^ 0x1021);
                }
                else
                {
                    crc <<= 1;
                }
            }
        }
        return crc.ToString("X4");
    }

    private static string RemoveAccents(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        string normalized = text.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder();
        foreach (char c in normalized)
        {
            var category = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (category != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(c == 'Đ' ? 'D' : (c == 'đ' ? 'd' : c));
            }
        }
        return sb.ToString().Normalize(NormalizationForm.FormC);
    }
}
