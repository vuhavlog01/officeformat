using System;

namespace VFormat.Core.Pdf
{
    /// <summary>
    /// Kết quả xử lý lật / sắp xếp trang PDF.
    /// Dịch trực tiếp từ mô hình ProcessResult trong LatPdf.
    /// </summary>
    public class PdfProcessResult
    {
        public string SourcePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public int OriginalPageCount { get; set; }
        public int FinalPageCount { get; set; }
        public int RotatedPagesCount { get; set; }
        public int RemovedBlankPagesCount { get; set; }
        public TimeSpan ElapsedTime { get; set; }
    }
}
