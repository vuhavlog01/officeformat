using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonVFormat
    {
        public enum QuyChuanDinhDang
        {
            NghiDinh30,
            HuongDan05
        }

        private static HashSet<string> _tuDienTiengViet = null;

        private void RibbonVFormat_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnCanLeA4_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;

            if (doc != null)
            {
                // Thiết lập khổ giấy A4
                doc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;

                // Căn lề chuẩn: Trên 2cm, Dưới 2cm, Trái 3cm, Phải 1.5cm
                doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(2.0f);
                doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2.0f);
                doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(3.0f);
                doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.5f);

                System.Windows.Forms.MessageBox.Show("Đã căn lề chuẩn A4 thành công!", "Thông báo");
            }
        }

        private static readonly HashSet<string> DanhSachNgoaiLe = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Internet", "Email", "Online", "Website", "Server", "Data", "Scanner",
            "Nẵng", "UBND", "HĐND", "VPTW", "TW", "BCH"
        };

        private void KhoiTaoTuDien()
        {
            if (_tuDienTiengViet != null) return;

            _tuDienTiengViet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Lấy chính xác thư mục chứa file DLL của Add-in thay vì dùng BaseDirectory
            string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string assemblyDir = Path.GetDirectoryName(assemblyLocation);
            string dictPath = Path.Combine(assemblyDir, "vietnamese_words.txt");

            if (File.Exists(dictPath))
            {
                var lines = File.ReadAllLines(dictPath, Encoding.UTF8);
                foreach (var rawLine in lines)
                {
                    if (string.IsNullOrWhiteSpace(rawLine)) continue;

                    string lineNormalized = rawLine.Normalize(NormalizationForm.FormC).Trim();
                    string[] tokens = lineNormalized.Split(new[] { ' ', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var token in tokens)
                    {
                        string word = token.Trim();
                        if (!string.IsNullOrEmpty(word))
                        {
                            _tuDienTiengViet.Add(word);
                        }
                    }
                }
            }
        }

        private string GetConfigFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folder = Path.Combine(appData, "VFormatConfig");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return Path.Combine(folder, "loaiVanBan.json");
        }

        // Danh sách mặc định ban đầu
        private readonly List<string> danhSachMacDinh = new List<string>
        {
            "THÔNG BÁO", "BÁO CÁO", "QUYẾT ĐỊNH", "KẾ HOẠCH",
            "TỜ TRÌNH", "CHỈ THỊ", "NGHỊ QUYẾT", "CÔNG VĂN", "KẾT LUẬN"
        };

        /// <summary>
        /// Nạp toàn bộ danh sách loại văn bản (bao gồm mặc định + người dùng đã thêm)
        /// </summary>
        private HashSet<string> LayDanhSachLoaiVanBan()
        {
            var tapHop = new HashSet<string>(danhSachMacDinh, StringComparer.OrdinalIgnoreCase);
            string filePath = GetConfigFilePath();

            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
                foreach (string line in lines)
                {
                    string trimmed = line.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                    {
                        tapHop.Add(trimmed.ToUpperInvariant());
                    }
                }
            }
            return tapHop;
        }

        private void btnBoldNumbering_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;

            if (doc == null) return;

            wordApp.ScreenUpdating = false;
            int count = 0;

            try
            {
                // Regex tìm các mẫu: Số La Mã ("I.", "II."), Số tự nhiên ("1."), Phân cấp ("1.1.", "2.1", "1.1.1.")
                // Bắt đầu từ đầu dòng hoặc sau các khoảng trắng/tab
                Regex numberingRegex = new Regex(@"^\s*([IVXLCDM]+\.|\d+(\.\d+)*\.?)", RegexOptions.Multiline);

                foreach (Paragraph p in doc.Paragraphs)
                {
                    string text = p.Range.Text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    Match match = numberingRegex.Match(text);
                    if (match.Success && !string.IsNullOrWhiteSpace(match.Value.Trim()))
                    {
                        int startOffset = match.Index;
                        int length = match.Length;

                        Range numRange = doc.Range(p.Range.Start + startOffset, p.Range.Start + startOffset + length);

                        // Kiểm tra nếu chưa được bôi đậm hoàn toàn (1 = wdToggle/True)
                        if (numRange.Bold != 1)
                        {
                            numRange.Bold = 1;
                            count++;
                        }
                    }
                }

                MessageBox.Show($"Đã kiểm tra và bôi đậm {count} mục đánh số!", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }
        }

        private void btnCheckSpelling_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;

            if (doc == null) return;

            KhoiTaoTuDien();

            if (_tuDienTiengViet == null || _tuDienTiengViet.Count == 0)
            {
                MessageBox.Show("Chưa nạp được dữ liệu từ điển tiếng Việt!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            wordApp.ScreenUpdating = false;
            int errorCount = 0;

            try
            {
                // Regex nhận diện các từ đơn
                Regex wordRegex = new Regex(@"\b[\p{L}\p{M}]+\b", RegexOptions.Compiled);

                foreach (Paragraph p in doc.Paragraphs)
                {
                    string text = p.Range.Text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    // Chuẩn hóa văn bản đoạn về cùng chuẩn FormC
                    string normalizedParagraph = text.Normalize(NormalizationForm.FormC);

                    // Bỏ qua tiêu ngữ/quốc hiệu nếu cần
                    string upperText = normalizedParagraph.ToUpperInvariant();
                    if (upperText.Contains("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") ||
                        upperText.Contains("ĐỘC LẬP - TỰ DO - HẠNH PHÚC"))
                    {
                        continue;
                    }

                    MatchCollection matches = wordRegex.Matches(normalizedParagraph);

                    foreach (Match match in matches)
                    {
                        string word = match.Value.Trim();

                        // 1. Bỏ qua nếu là từ viết tắt hoàn toàn (hoa toàn bộ) hoặc độ dài 1 ký tự latin
                        if (word.Length > 1 && word == word.ToUpper()) continue;

                        // 2. Bỏ qua các từ nằm trong Whitelist ngoại lệ
                        if (DanhSachNgoaiLe.Contains(word)) continue;

                        // 3. Chuẩn hóa quy tắc bỏ dấu thay thế (hòa/hoà, hóa/hoá, thủy/thuỷ, hủy/huỷ)
                        string altWord = ChuyenDoiQuyTacDau(word);

                        // 4. Tra cứu đối chiếu
                        bool hopLe = _tuDienTiengViet.Contains(word) ||
                                     _tuDienTiengViet.Contains(altWord);

                        if (!hopLe)
                        {
                            int startOffset = p.Range.Start + match.Index;
                            int endOffset = startOffset + match.Length;

                            if (endOffset <= p.Range.End)
                            {
                                Range errorRange = doc.Range(startOffset, endOffset);
                                errorRange.HighlightColorIndex = WdColorIndex.wdYellow;
                                errorCount++;
                            }
                        }
                    }
                }

                MessageBox.Show($"Quét xong! Phát hiện {errorCount} vị trí nghi ngờ lỗi chính tả.", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }
        }

        private void btnCheckSpacing_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;

            if (doc == null) return;

            wordApp.ScreenUpdating = false;
            int spaceErrorCount = 0;

            try
            {
                // Regex phát hiện:
                // 1. Hai dấu cách trở lên: "  +"
                // 2. Dấu cách đứng trước dấu ngắt câu: " +[,.:;!?)]"
                // 3. Dấu cách đứng ngay sau dấu mở ngoặc: "[(“]\s+"
                Regex spaceRegex = new Regex(@"( {2,})|(\s+[,.:;!?\)])|([(\“]\s+)", RegexOptions.Compiled);

                foreach (Paragraph p in doc.Paragraphs)
                {
                    string text = p.Range.Text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    // Loại trừ Quốc hiệu và Tiêu ngữ
                    string upperText = text.ToUpperInvariant();
                    if (upperText.Contains("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") ||
                        upperText.Contains("ĐỘC LẬP - TỰ DO - HẠNH PHÚC") ||
                        upperText.Contains("ĐỘC LẬP – TỰ DO – HẠNH PHÚC"))
                    {
                        continue;
                    }

                    MatchCollection matches = spaceRegex.Matches(text);
                    foreach (Match m in matches)
                    {
                        // Tránh ký tự ngắt đoạn \r\n ở cuối dòng
                        int matchStart = p.Range.Start + m.Index;
                        int matchEnd = matchStart + m.Length;

                        if (matchEnd <= p.Range.End)
                        {
                            Range errorRange = doc.Range(matchStart, matchEnd);
                            errorRange.HighlightColorIndex = WdColorIndex.wdTurquoise; // Tô màu xanh ngọc lam
                            spaceErrorCount++;
                        }
                    }
                }

                MessageBox.Show($"Đã phát hiện và tô màu {spaceErrorCount} vị trí lỗi khoảng trắng (màu Xanh lơ).", "Kiểm tra khoảng trắng", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }
        }

        private void btnClearHighlights_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (doc != null)
            {
                doc.Content.HighlightColorIndex = WdColorIndex.wdNoHighlight;
                MessageBox.Show("Đã xóa toàn bộ màu tô đánh dấu trên tài liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string ChuyenDoiQuyTacDau(string input)
        {
            string s = input;
            // Nhóm oa, oe, uy
            s = s.Replace("oà", "òa").Replace("oá", "óa").Replace("oả", "ỏa").Replace("oã", "õa").Replace("oạ", "ọa")
                 .Replace("uỳ", "ủy").Replace("uý", "úy").Replace("uỷ", "ủy").Replace("uỹ", "ũy").Replace("uỵ", "ụy")
                 .Replace("oè", "òe").Replace("oé", "óe").Replace("oẻ", "ỏe").Replace("oẽ", "õe").Replace("oẹ", "ọe");

            if (s == input)
            {
                // Thử đảo ngược ngược lại nếu văn bản gốc ở dạng dấu mới
                s = s.Replace("òa", "oà").Replace("óa", "oá").Replace("ỏa", "oả").Replace("õa", "oã").Replace("ọa", "oạ")
                     .Replace("ủy", "uỷ").Replace("úy", "uý").Replace("ủy", "uỷ").Replace("ũy", "uỹ").Replace("ụy", "uỵ")
                     .Replace("òe", "oè").Replace("óe", "oé").Replace("ỏe", "oẻ").Replace("õe", "oẽ").Replace("ọe", "oẹ");
            }

            return s;
        }

        private void btnChuanHoaNoiNhan_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;

            if (doc == null) return;

            wordApp.ScreenUpdating = false;
            bool found = false;

            try
            {
                // 1. Nếu văn bản dùng Table (chân trang chia 2 cột: Nơi nhận bên trái, Ký bên phải)
                foreach (Table tbl in doc.Tables)
                {
                    foreach (Row row in tbl.Rows)
                    {
                        for (int c = 1; c <= row.Cells.Count; c++)
                        {
                            Cell cell = row.Cells[c];
                            string cellText = cell.Range.Text;

                            if (cellText.IndexOf("Nơi nhận", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // 1. Chuẩn hoá ô Nơi nhận bên trái
                                XuLyTrongCell(cell);

                                // 2. Format ô Chữ ký bên phải (nếu hàng có từ 2 ô trở lên)
                                if (c < row.Cells.Count)
                                {
                                    Cell rightCell = row.Cells[c + 1];
                                    XuLyCotKy(rightCell);
                                }

                                found = true;
                                break;
                            }
                        }
                        if (found) break;
                    }
                    if (found) break;
                }

                // 2. Nếu không nằm trong Table, quét các đoạn văn bản độc lập
                if (!found)
                {
                    ChuyenNoiNhanVaKyTenSangTable(doc);
                    MessageBox.Show("Đã thay đổi định dạng nơi nhận", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (found)
                {
                    MessageBox.Show("Đã chuẩn hóa thành công phần Nơi nhận!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy mục 'Nơi nhận' trong tài liệu!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xử lý: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }
        }

        private void XuLyTrongCell(Cell cell)
        {
            if (cell == null) return;

            // BƯỚC 1: Chuyển đổi toàn bộ Shift + Enter (^l / \v) thành Enter chuẩn (\r)
            // Giúp mỗi dòng thành phần trở thành một Paragraph riêng biệt, tránh bị ghi đè mất chữ
            Word.Find find = cell.Range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = "^l";         // Ký hiệu Manual Line Break (Shift + Enter) trong Word
            find.Replacement.Text = "^p"; // Ký hiệu Paragraph Mark (Enter chuẩn)
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchWildcards = false;
            find.Execute(Replace: WdReplace.wdReplaceAll);

            // BƯỚC 2: Thu thập các đoạn văn sau khi đã tách dòng chuẩn
            List<Paragraph> itemParas = new List<Paragraph>();
            Paragraph headerPara = null;

            for (int i = 1; i <= cell.Range.Paragraphs.Count; i++)
            {
                Paragraph p = cell.Range.Paragraphs[i];
                string cleanText = p.Range.Text.TrimEnd('\r', '\n', '\a', '\v', ' ', '\t');

                if (string.IsNullOrWhiteSpace(cleanText)) continue;

                if (cleanText.StartsWith("Nơi nhận", StringComparison.OrdinalIgnoreCase))
                {
                    headerPara = p;
                }
                else if (headerPara != null)
                {
                    // Các dòng sau dòng Nơi nhận
                    itemParas.Add(p);
                }
            }

            DinhDangNoiNhan(headerPara, itemParas);
        }

        

        /// <summary>
        /// Hàm cốt lõi: Gán định dạng chuẩn xác cho Header và từng dòng Nơi nhận
        /// </summary>
        private void DinhDangNoiNhan(Paragraph headerPara, List<Paragraph> itemParas)
        {
            // 1. Xử lý dòng tiêu đề "Nơi nhận:" (chỉ thay đổi chữ, không đụng vào dấu ngắt dòng)
            if (headerPara != null)
            {
                Range hRange = headerPara.Range.Duplicate;
                // Co bớt con trỏ để không chọn ký tự ngắt đoạn \r\a ở cuối
                if (hRange.End > hRange.Start)
                {
                    string endMarker = hRange.Text.EndsWith("\r\a") ? "\r\a" : (hRange.Text.EndsWith("\r") ? "\r" : "");
                    hRange.End = hRange.End - (endMarker.Length > 0 ? endMarker.Length : 1);

                    if (hRange.End >= hRange.Start)
                    {
                        hRange.Text = "Nơi nhận:";
                    }
                }

                headerPara.Range.Font.Name = "Times New Roman";
                headerPara.Range.Font.Size = 14f; // Theo HD 05, Nơi nhận thường cỡ 12
                headerPara.Range.Font.Bold = 1;   // In đậm
                headerPara.Range.Font.Italic = 0; // In nghiêng
                headerPara.Range.Font.Underline = WdUnderline.wdUnderlineSingle; // Gạch chân theo mẫu
                headerPara.Format.LeftIndent = 0;
                headerPara.Format.FirstLineIndent = 0;
            }

            // 2. Xử lý các dòng thành phần bên dưới
            int total = itemParas.Count;
            for (int i = 0; i < total; i++)
            {
                Paragraph p = itemParas[i];
                Range r = p.Range.Duplicate;
                string text = p.Range.Text.TrimEnd('\r', '\n', '\a', '\v', ' ', '\t');

                if (string.IsNullOrWhiteSpace(text)) continue;

                // Xóa các gạch đầu dòng, dấu cách, tab cũ ở đầu dòng
                text = Regex.Replace(text, @"^[\s\t\-–—\*\•]+", "").Trim();

                // Xóa toàn bộ dấu kết thúc câu cũ (., ;, :, khoảng trắng) ở cuối
                text = Regex.Replace(text, @"[\.,;\s\t]+$", "").Trim();

                // Dòng cuối kết thúc bằng '.', các dòng trên kết thúc bằng ';'
                bool isLast = (i == total - 1);
                string endingChar = isLast ? "." : ";";

                // Co ngắn Range để giữ lại ký hiệu ngắt đoạn của Cell/Paragraph
                string endMarker = p.Range.Text.EndsWith("\r\a") ? "\r\a" : (p.Range.Text.EndsWith("\r") ? "\r" : "");
                r.End = r.End - (endMarker.Length > 0 ? endMarker.Length : 1);

                if (r.End >= r.Start)
                {
                    r.Text = "- " + text + endingChar;
                }

                // Định dạng Font chữ
                p.Range.Font.Name = "Times New Roman";
                p.Range.Font.Size = 12f; // Chữ nơi nhận thường nhỏ hơn 1-2 size (11pt)
                p.Range.Font.Bold = 0;
                p.Range.Font.Italic = 0;
                p.Range.Font.Underline = WdUnderline.wdUnderlineNone;

                // Căn lề trái = 0
                p.Format.LeftIndent = 0;
                p.Format.FirstLineIndent = 0;
                p.Format.SpaceBefore = 0;
                p.Format.SpaceAfter = 0;
                p.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
        }

        private void btnConvertFooterToTable_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;
            ChuyenNoiNhanVaKyTenSangTable(doc);
        }

        /// <summary>
        /// Chuyển phần cuối văn bản từ dạng căn bằng Tab/khoảng trắng
        /// sang Table 2 cột:
        ///     Cột trái  = Nơi nhận
        ///     Cột phải  = Khối ký tên
        ///
        /// Xử lý được trường hợp:
        /// "Nơi nhận:                         T/M BAN THƯỜNG VỤ"
        /// nằm trên CÙNG MỘT paragraph.
        /// </summary>
        private void ChuyenNoiNhanVaKyTenSangTable(Word.Document doc)
        {
            Word.Paragraph paraBatDau = null;

            // ============================================================
            // 1. Tìm paragraph chứa "Nơi nhận:"
            // ============================================================
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                try
                {
                    // Nếu paragraph đã nằm trong Table thì không xử lý
                    if (p.Range.Information[
                        Word.WdInformation.wdWithInTable])
                    {
                        continue;
                    }

                    string text = CleanParagraphText(p.Range.Text);

                    if (text.StartsWith(
                        "Nơi nhận:",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        paraBatDau = p;
                        break;
                    }
                }
                catch
                {
                    // Bỏ qua paragraph lỗi
                }
            }

            if (paraBatDau == null)
                return;

            // ============================================================
            // 2. Thu thập toàn bộ các paragraph từ "Nơi nhận:"
            //    đến hết văn bản
            // ============================================================

            int startIndex = -1;

            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                if (doc.Paragraphs[i].Range.Start ==
                    paraBatDau.Range.Start)
                {
                    startIndex = i;
                    break;
                }
            }

            if (startIndex < 0)
                return;

            // ============================================================
            // 3. Tạo vùng dữ liệu trái/phải
            // ============================================================

            List<string> noiNhanLines = new List<string>();
            List<string> kyTenLines = new List<string>();

            bool coKhoiKyTen = false;

            int endParagraphIndex = doc.Paragraphs.Count;

            // ============================================================
            // 4. Phân tích từng dòng
            // ============================================================

            for (int i = startIndex;
                 i <= endParagraphIndex;
                 i++)
            {
                Word.Paragraph p = doc.Paragraphs[i];

                try
                {
                    // Nếu gặp Table thì dừng
                    if (p.Range.Information[
                        Word.WdInformation.wdWithInTable])
                    {
                        break;
                    }

                    string rawText = p.Range.Text;

                    string text = rawText
                        .Replace("\r", "")
                        .Replace("\a", "");

                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    // ====================================================
                    // Tách một dòng thành trái / phải
                    //
                    // Ưu tiên TAB.
                    // Nếu không có TAB thì tìm chuỗi >= 4 khoảng trắng.
                    // ====================================================

                    string leftText = "";
                    string rightText = "";

                    int separatorIndex = -1;
                    int separatorLength = 0;

                    // ----------------------------------------------------
                    // Trường hợp có TAB
                    // ----------------------------------------------------
                    int tabIndex = text.IndexOf('\t');

                    if (tabIndex >= 0)
                    {
                        separatorIndex = tabIndex;
                        separatorLength = 1;
                    }
                    else
                    {
                        // ------------------------------------------------
                        // Trường hợp dùng nhiều khoảng trắng để căn cột
                        // ------------------------------------------------
                        Match matchSpace = Regex.Match(
                            text,
                            @"\s{4,}");

                        if (matchSpace.Success)
                        {
                            separatorIndex = matchSpace.Index;
                            separatorLength = matchSpace.Length;
                        }
                    }

                    // ====================================================
                    // Có phân cách trái/phải
                    // ====================================================
                    if (separatorIndex >= 0)
                    {
                        leftText = text
                            .Substring(0, separatorIndex)
                            .Trim();

                        rightText = text
                            .Substring(
                                separatorIndex + separatorLength)
                            .Trim();

                        // Có nội dung bên phải => xác định đây là
                        // bố cục 2 cột.
                        if (!string.IsNullOrWhiteSpace(rightText))
                        {
                            coKhoiKyTen = true;
                        }
                    }
                    else
                    {
                        // =================================================
                        // Không có phân cách
                        //
                        // Nếu chưa vào khối ký tên => đưa sang trái.
                        // Nếu đã xác định khối ký tên và dòng này là
                        // chức danh => đưa sang phải.
                        // =================================================

                        if (LaDongKyTenTiepTheo(text))
                        {
                            rightText = text;
                            coKhoiKyTen = true;
                        }
                        else
                        {
                            leftText = text;
                        }
                    }

                    // ====================================================
                    // Bổ sung vào 2 danh sách
                    // ====================================================

                    if (!string.IsNullOrWhiteSpace(leftText))
                    {
                        noiNhanLines.Add(leftText);
                    }
                    else
                    {
                        // Giữ dòng trống tương ứng
                        noiNhanLines.Add("");
                    }

                    if (!string.IsNullOrWhiteSpace(rightText))
                    {
                        kyTenLines.Add(rightText);
                    }
                    else
                    {
                        kyTenLines.Add("");
                    }

                    // ====================================================
                    // Khi đã đi qua phần Nơi nhận và khối ký tên,
                    // nếu gặp nội dung mới thì có thể dừng.
                    //
                    // Trong trường hợp văn bản kết thúc ngay tại đây
                    // thì vòng lặp sẽ đi đến hết tài liệu.
                    // ====================================================
                }
                catch
                {
                    // Bỏ qua paragraph lỗi
                }
            }

            // ============================================================
            // 5. Kiểm tra có thực sự tìm thấy bố cục 2 cột hay không
            // ============================================================

            bool coNoiNhan = noiNhanLines.Any(
                x => x.StartsWith(
                    "Nơi nhận:",
                    StringComparison.OrdinalIgnoreCase));

            bool coKyTen = kyTenLines.Any(
                x => !string.IsNullOrWhiteSpace(x));

            if (!coNoiNhan || !coKyTen)
            {
                return;
            }

            // ============================================================
            // 6. Xác định Range cũ cần thay thế
            // ============================================================

            int startPosition = paraBatDau.Range.Start;

            int endPosition = doc.Content.End - 1;

            Word.Range oldRange = doc.Range(
                startPosition,
                endPosition);

            // ============================================================
            // 7. Xóa phần cũ
            // ============================================================

            oldRange.Delete();

            // ============================================================
            // 8. Tạo Table 1 hàng 2 cột
            // ============================================================

            Word.Range insertRange = doc.Range(
                startPosition,
                startPosition);

            Word.Table table = doc.Tables.Add(
                insertRange,
                1,
                2);

            // Không có đường viền
            table.Borders.Enable = 0;

            table.AllowAutoFit = false;

            // ============================================================
            // 9. Chiều rộng Table
            // ============================================================

            float pageWidth =
                doc.PageSetup.PageWidth;

            float leftMargin =
                doc.PageSetup.LeftMargin;

            float rightMargin =
                doc.PageSetup.RightMargin;

            float usableWidth =
                pageWidth -
                leftMargin -
                rightMargin;

            table.Columns[1].Width =
                usableWidth * 0.50f;

            table.Columns[2].Width =
                usableWidth * 0.50f;

            // ============================================================
            // 10. Ghép nội dung cột trái
            // ============================================================

            StringBuilder leftBuilder =
                new StringBuilder();

            foreach (string line in noiNhanLines)
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    leftBuilder.Append(line);
                }

                leftBuilder.Append("\r");
            }

            // ============================================================
            // 11. Ghép nội dung cột phải
            // ============================================================

            StringBuilder rightBuilder =
                new StringBuilder();

            foreach (string line in kyTenLines)
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    rightBuilder.Append(line);
                }

                rightBuilder.Append("\r");
            }

            // ============================================================
            // 12. Đưa nội dung vào Cell trái
            // ============================================================

            Word.Cell cellLeft = table.Cell(1, 1);

            Word.Range leftRange =
                cellLeft.Range;

            leftRange.End -= 1;

            leftRange.Text =
                leftBuilder.ToString();

            // ============================================================
            // 13. Đưa nội dung vào Cell phải
            // ============================================================

            Word.Cell cellRight = table.Cell(1, 2);

            Word.Range rightRange =
                cellRight.Range;

            rightRange.End -= 1;

            rightRange.Text =
                rightBuilder.ToString();

            // ============================================================
            // 14. Định dạng cột trái
            // ============================================================

            leftRange =
                cellLeft.Range.Duplicate;

            leftRange.End -= 1;

            leftRange.Font.Name =
                "Times New Roman";

            foreach (Word.Paragraph p in
                cellLeft.Range.Paragraphs)
            {
                p.Format.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphLeft;

                p.Format.FirstLineIndent = 0;
                p.Format.LeftIndent = 0;
                p.Format.RightIndent = 0;

                p.Format.SpaceBefore = 0;
                p.Format.SpaceAfter = 0;

                p.Format.LineSpacingRule =
                    Word.WdLineSpacing.wdLineSpaceSingle;

                string text =
                    CleanParagraphText(p.Range.Text);

                if (text.StartsWith(
                    "Nơi nhận:",
                    StringComparison.OrdinalIgnoreCase))
                {
                    p.Range.Font.Bold = 0;
                    p.Range.Font.Italic = 0;
                    p.Range.Font.Underline =
                        Word.WdUnderline.wdUnderlineSingle;
                    p.Range.Font.Size = 14;
                }
                else if (text.StartsWith("-"))
                {
                    p.Range.Font.Bold = 0;
                    p.Range.Font.Italic = 0;
                    p.Range.Font.Underline =
                        Word.WdUnderline.wdUnderlineNone;
                    p.Range.Font.Size = 12;
                }
            }

            // ============================================================
            // 15. Định dạng cột phải
            // ============================================================

            rightRange =
                cellRight.Range.Duplicate;

            rightRange.End -= 1;

            rightRange.Font.Name =
                "Times New Roman";

            foreach (Word.Paragraph p in
                cellRight.Range.Paragraphs)
            {
                p.Format.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;

                p.Format.FirstLineIndent = 0;
                p.Format.LeftIndent = 0;
                p.Format.RightIndent = 0;

                p.Format.SpaceBefore = 0;
                p.Format.SpaceAfter = 0;

                p.Format.LineSpacingRule =
                    Word.WdLineSpacing.wdLineSpaceSingle;

                string text =
                    CleanParagraphText(p.Range.Text);
                //bool isSignerName = (p == totalValid - 1);
                //XuLyTrongCell(Word.Cell);
                //MessageBox.Show(cellRight.Ran, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (text.StartsWith("T/M") ||
                    text.StartsWith("TM.") ||
                    text.StartsWith("TM "))
                {
                    p.Range.Font.Bold = 1;
                    p.Range.Font.Size = 14;
                }
                else
                {
                    p.Range.Font.Bold = 0;
                    p.Range.Font.Size = 14;
                }
            }
        }


        /// <summary>
        /// Nhận diện các dòng thuộc khối ký tên.
        /// Ví dụ: PHÓ BÍ THƯ, BÍ THƯ, CHỦ TỊCH...
        /// </summary>
        private bool LaDongKyTenTiepTheo(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string t = text.Trim().ToUpperInvariant();

            string[] danhSachChucDanh =
            {
        "BÍ THƯ",
        "PHÓ BÍ THƯ",
        "CHỦ TỊCH",
        "PHÓ CHỦ TỊCH",
        "CHÁNH VĂN PHÒNG",
        "PHÓ CHÁNH VĂN PHÒNG",
        "TRƯỞNG BAN",
        "PHÓ TRƯỞNG BAN",
        "KT.",
        "T/M",
        "TM.",
        "TM ",
        "TL.",
        "TL ",
        "TUQ.",
        "TUQ "
    };

            foreach (string chucDanh in danhSachChucDanh)
            {
                if (t.StartsWith(chucDanh))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Làm sạch nội dung Paragraph của Word.
        /// Loại bỏ ký tự kết thúc Paragraph và ký tự kết thúc Cell.
        /// </summary>
        private string CleanParagraphText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            return text
                .Replace("\r", "")
                .Replace("\a", "")
                .Trim();
        }


        private void XuLyCotKy(Cell rightCell)
        {
            if (rightCell == null) return;

            List<Paragraph> validParas = new List<Paragraph>();

            // Thu thập các dòng có chứa chữ (loại bỏ ký tự rác kết thúc ô \r, \a)
            for (int i = 1; i <= rightCell.Range.Paragraphs.Count; i++)
            {
                Paragraph p = rightCell.Range.Paragraphs[i];
                string text = p.Range.Text.TrimEnd('\r', '\n', '\a', ' ', '\t');

                if (!string.IsNullOrWhiteSpace(text))
                {
                    validParas.Add(p);
                }
            }

            if (validParas.Count == 0) return;

            int totalValid = validParas.Count;

            for (int i = 0; i < totalValid; i++)
            {
                Paragraph p = validParas[i];
                Range r = p.Range;
                string text = r.Text.TrimEnd('\r', '\n', '\a', ' ', '\t');

                // Định dạng cơ bản: Font Times New Roman, Size 14, Căn giữa
                r.Font.Name = "Times New Roman";
                r.Font.Size = 14f;
                r.Font.Underline = WdUnderline.wdUnderlineNone;
                p.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                p.Format.LeftIndent = 0;
                p.Format.RightIndent = 0;

                bool isSignerName = (i == totalValid - 1); // Dòng cuối cùng là Họ tên người ký

                // In đậm nếu là:
                // 1. Tên người ký (dòng cuối)
                // 2. Chức danh/thẩm quyền (TM., KT., BÍ THƯ, CHỦ TỊCH,...)
                if (isSignerName || !Regex.IsMatch(text, @"^(TM\.|T/L\.|KT\.|BÍ THƯ|PHÓ BÍ THƯ|CHỦ TỊCH|PHÓ CHỦ TỊCH)", RegexOptions.IgnoreCase))
                {
                    r.Font.Bold = 1;
                    r.Font.Italic = 0;
                }
                else
                {
                    // Các dòng ghi chú phụ khác (nếu có)
                    r.Font.Bold = 0;
                }
            }
        }

        /// <summary>
        /// Chỉ chuẩn hóa nội dung bên dưới phần tên loại văn bản (THÔNG BÁO, BÁO CÁO...) và trích yếu.
        /// Giữ nguyên vẹn toàn bộ phần bảng đầu trang, tiêu đề chính và trích yếu.
        /// </summary>
        private void XuLyNoiDungDeMucTheoHD05(Word.Document doc, Word.Application app)
        {
            float indent11cm = app.CentimetersToPoints(1.27f); // 1.1 cm trên Ruler
            int totalParas = doc.Paragraphs.Count;

            bool daGapTieuDeChinh = false;
            bool batDauChuanHoaThanVanBan = false;

            // Danh sách các tên loại văn bản phổ biến theo thể thức
            HashSet<string> cacLoaiVanBan = LayDanhSachLoaiVanBan();
        //    string[] cacLoaiVanBan = new string[]
        //    {
        //"THÔNG BÁO", "BÁO CÁO", "QUYẾT ĐỊNH", "KẾ HOẠCH",
        //"TỜ TRÌNH", "CHỈ THỊ", "NGHỊ QUYẾT", "CÔNG VĂN", "KẾT LUẬN"
        //    };

            for (int i = 1; i <= totalParas; i++)
            {
                try
                {
                    Word.Paragraph para = doc.Paragraphs[i];

                    // 1. Tuyệt đối không đụng vào Bảng (Quốc hiệu, cơ quan ban hành)
                    if (para.Range.Information[Word.WdInformation.wdWithInTable])
                        continue;

                    string text = LamSachText(para.Range.Text);
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    // 2. Tìm dòng tên loại văn bản chính (ví dụ: THÔNG BÁO, BÁO CÁO)
                    if (!daGapTieuDeChinh)
                    {
                        foreach (string loai in cacLoaiVanBan)
                        {
                            if (text.Equals(loai, StringComparison.OrdinalIgnoreCase))
                            {
                                daGapTieuDeChinh = true;
                                break;
                            }
                        }
                        // Nếu chưa tới tiêu đề chính, bỏ qua hoàn toàn không can thiệp
                        continue;
                    }

                    // 3. Đã gặp tiêu đề chính: Tiếp tục bỏ qua các dòng trích yếu và dấu gạch "----"
                    if (!batDauChuanHoaThanVanBan)
                    {
                        // Dấu gạch nối ngăn cách trích yếu
                        if (text.StartsWith("---") || text.StartsWith("–-"))
                        {
                            continue;
                        }

                        // Nếu dòng là trích yếu (thường viết chữ thường, in đậm, căn giữa)
                        // Ta nhận diện đoạn văn thân bài khi bắt đầu bằng các từ mở đầu hoặc có dấu kết thúc câu '.'
                        // Ví dụ: "Ngày tháng 10 năm 2026..." hoặc "Thực hiện Kế hoạch..."
                        bool laDoanMoDau = text.StartsWith("Ngày", StringComparison.OrdinalIgnoreCase) ||
                                          text.StartsWith("Thực hiện", StringComparison.OrdinalIgnoreCase) ||
                                          text.StartsWith("Căn cứ", StringComparison.OrdinalIgnoreCase) ||
                                          text.StartsWith("I.", StringComparison.OrdinalIgnoreCase) ||
                                          text.StartsWith("1.", StringComparison.OrdinalIgnoreCase) ||
                                          para.Format.Alignment == Word.WdParagraphAlignment.wdAlignParagraphJustify ||
                                          para.Format.Alignment == Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        // Nếu vẫn đang căn giữa (Center) và chưa có dấu chấm hết câu dài -> Vẫn là trích yếu, bỏ qua
                        if (para.Format.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter && !laDoanMoDau)
                        {
                            continue;
                        }

                        // Bắt đầu bước vào phần thân văn bản từ đoạn này
                        batDauChuanHoaThanVanBan = true;
                    }

                    // =========================================================================
                    // 4. CHỈ ÁP DỤNG ĐỊNH DẠNG TỪ ĐOẠN THÂN VĂN BẢN TRỞ ĐI
                    // =========================================================================

                    // Gỡ bỏ tính năng Bullet/Numbering tự động của Word để tránh giãn tab xa
                    try
                    {
                        if (para.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering)
                        {
                            para.Range.ListFormat.RemoveNumbers(Word.WdNumberType.wdNumberParagraph);
                        }
                    }
                    catch { }

                    // Cấu hình Ruler: Dòng đầu 1.1cm, hàng tiếp theo về 0cm
                    para.Format.LeftIndent = 0;
                    para.Format.RightIndent = 0;
                    para.Format.FirstLineIndent = indent11cm;
                    para.Format.TabStops.ClearAll();

                    // Dãn dòng và dãn đoạn
                    para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    para.Format.LineSpacing = app.LinesToPoints(1.2f);
                    para.Format.SpaceBefore = 0f;
                    para.Format.SpaceAfter = 6f;
                    para.Format.WidowControl = -1;

                    // --- ĐỀ MỤC CẤP 1: I. Nội dung ---
                    Match matchLaMa = Regex.Match(text, @"^([IVXLCDM]+)\s*[\.-]\s*(.+)$", RegexOptions.IgnoreCase);
                    if (matchLaMa.Success)
                    {
                        string laMa = matchLaMa.Groups[1].Value.ToUpperInvariant();
                        string noiDung = ChuanHoaKhoangTrang(matchLaMa.Groups[2].Value);

                        ThayNoiDungAnToan(para, laMa + ". " + noiDung);
                        para.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para.Range.Font.Bold = -1;
                        para.Range.Font.Italic = 0;
                        para.Range.Font.Size = 14;
                        continue;
                    }

                    // --- ĐỀ MỤC CẤP 3: 1.1. Nội dung (Kiểm tra trước Cấp 2) ---
                    Match matchCap3 = Regex.Match(text, @"^(\d+\.\d+)\.?[\s\t]*(.+)$");
                    if (matchCap3.Success)
                    {
                        string stt = matchCap3.Groups[1].Value;
                        string noiDung = ChuanHoaKhoangTrang(matchCap3.Groups[2].Value);

                        ThayNoiDungAnToan(para, stt + ". " + noiDung);
                        para.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para.Range.Font.Bold = -1;
                        para.Range.Font.Italic = -1; // In đậm nghiêng
                        para.Range.Font.Size = 14;
                        continue;
                    }

                    // --- ĐỀ MỤC CẤP 2: 1. Nội dung ---
                    Match matchCap2 = Regex.Match(text, @"^(\d+)\.[\s\t]*(.+)$");
                    if (matchCap2.Success)
                    {
                        string stt = matchCap2.Groups[1].Value;
                        string noiDung = ChuanHoaKhoangTrang(matchCap2.Groups[2].Value);

                        ThayNoiDungAnToan(para, stt + ". " + noiDung);
                        para.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para.Range.Font.Bold = -1;
                        para.Range.Font.Italic = 0;
                        para.Range.Font.Size = 14;
                        continue;
                    }

                    // --- DÒNG GẠCH ĐẦU DÒNG "-" ---
                    if (text.StartsWith("-"))
                    {
                        string noiDung = ChuanHoaKhoangTrang(text.Substring(1));

                        ThayNoiDungAnToan(para, "- " + noiDung);
                        para.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                        para.Range.Font.Bold = 0;
                        para.Range.Font.Italic = 0;
                        para.Range.Font.Size = 14;
                        continue;
                    }

                    // --- ĐOẠN VĂN MỞ ĐẦU HOẶC NỘI DUNG THƯỜNG ---
                    para.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    para.Range.Font.Bold = 0;
                    para.Range.Font.Italic = 0;
                    para.Range.Font.Size = 14;
                }
                catch { }
            }
        }

        /// <summary>
        /// Hàm dọn sạch ký tự Tab (\t) và các khoảng trắng liên tiếp thành đúng 1 dấu cách
        /// </summary>
        private string ChuanHoaKhoangTrang(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            // Thay thế tab thành dấu cách và rút gọn nhiều khoảng trắng liên tiếp về 1 dấu cách duy nhất
            return Regex.Replace(input.Replace("\t", " "), @"\s+", " ").Trim();
        }

        /// <summary>
        /// Thay thế nội dung một Paragraph mà không làm hỏng ký tự điều khiển ngắt dòng (\r)
        /// </summary>
        private void ThayNoiDungAnToan(Word.Paragraph para, string noiDungMoi)
        {
            Word.Range r = para.Range;
            if (r.End > r.Start)
            {
                Word.Range editRange = r.Duplicate;
                editRange.End = editRange.End - 1; // Giữ lại ký tự ngắt đoạn \r
                if (editRange.End >= editRange.Start)
                {
                    editRange.Text = noiDungMoi;
                }
            }
        }

        /// <summary>
        /// Loại bỏ các ký tự xuống dòng, ngắt ô bảng biểu và khoảng trắng thừa ở 2 đầu
        /// </summary>
        private string LamSachText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            return text
                .Replace("\r", "")
                .Replace("\n", "")
                .Replace("\a", "") // Ký tự kết thúc ô bảng biểu (cell marker)
                .Trim();
        }

        private void btnXuLyNoiDungDeMucTheo_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;

            // Kiểm tra xem người dùng có đang mở tài liệu nào không
            if (wordApp.Documents.Count == 0)
            {
                MessageBox.Show("Hiện không có tài liệu nào đang mở!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var doc = wordApp.ActiveDocument;

            try
            {
                // Tắt cập nhật màn hình tạm thời để tăng tốc xử lý và không bị giật lag
                wordApp.ScreenUpdating = false;

                // Truyền đủ cả 2 tham số: doc và wordApp
                XuLyNoiDungDeMucTheoHD05(doc, wordApp);

                MessageBox.Show("Đã chuẩn hóa đầu dòng và đầu mục thành công!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xử lý đề mục: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Bật lại cập nhật màn hình
                wordApp.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// Lưu thêm 1 loại văn bản mới vào file cấu hình
        /// </summary>
        private bool LuuLoaiVanBanMoi(string tenMoi)
        {
            if (string.IsNullOrWhiteSpace(tenMoi)) return false;

            tenMoi = tenMoi.Trim().ToUpperInvariant();
            var danhSachHienTai = LayDanhSachLoaiVanBan();

            if (danhSachHienTai.Contains(tenMoi))
            {
                return false; // Đã tồn tại trong danh sách
            }

            string filePath = GetConfigFilePath();
            // Ghi nối tiếp vào file UTF-8
            File.AppendAllLines(filePath, new[] { tenMoi }, Encoding.UTF8);
            return true;
        }
        private void btnThemLoaiVanBan_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;

            // 1. Kiểm tra xem có tài liệu nào đang mở không
            if (wordApp.Documents.Count == 0)
            {
                MessageBox.Show("Hiện không có tài liệu nào đang mở!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Lấy đoạn văn bản đang được người dùng bôi đen (Selection)
            Word.Selection selection = wordApp.Selection;
            if (selection == null || string.IsNullOrWhiteSpace(selection.Text))
            {
                MessageBox.Show("Vui lòng bôi đen cụm từ loại văn bản cần thêm (ví dụ: QUY CHẾ, HƯỚNG DẪN, ĐỀ ÁN...) trước khi bấm nút!",
                    "Chưa chọn văn bản", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Làm sạch chuỗi được chọn (bỏ ký tự ngắt dòng, khoảng trắng thừa)
            string tuDuocChon = LamSachText(selection.Text).ToUpperInvariant();

            // Ràng buộc an toàn: Không cho lưu cụm từ quá dài (thường tên loại văn bản dưới 50 ký tự)
            if (tuDuocChon.Length > 50)
            {
                MessageBox.Show("Cụm từ được chọn quá dài, vui lòng chỉ bôi đen đúng tên loại văn bản!",
                    "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 3. Tiến hành kiểm tra và lưu
            var danhSach = LayDanhSachLoaiVanBan();
            if (danhSach.Contains(tuDuocChon))
            {
                MessageBox.Show($"Loại văn bản \"{tuDuocChon}\" đã có sẵn trong danh sách nhận diện!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Xác nhận với người dùng
            var xacNhan = MessageBox.Show(
                $"Bạn có muốn thêm loại văn bản: \"{tuDuocChon}\" vào danh sách nhận diện chuẩn hóa không?",
                "Xác nhận thêm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (xacNhan == DialogResult.Yes)
            {
                LuuLoaiVanBanMoi(tuDuocChon);
                MessageBox.Show($"Đã thêm thành công \"{tuDuocChon}\"!\nTừ nay các văn bản có tiêu đề này sẽ được tự động nhận diện chính xác.",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Chuẩn hóa phần Kính gửi / Kính trình theo thể thức:
        /// - Xử lý trường hợp nằm trong Table 2 cột hoặc dạng Paragraph thông thường
        /// - Chuyển Shift+Enter thành Enter
        /// - Các dòng kết thúc bằng ';', dòng cuối cùng kết thúc bằng '.'
        /// </summary>
        private void ChuanHoaPhanKinhGui(Word.Document doc)
        {
            if (doc == null) return;

            // TRƯỜNG HỢP 1: Kính gửi nằm trong Table (như ảnh mẫu)
            foreach (Word.Table table in doc.Tables)
            {
                // Kiểm tra xem bảng có chứa chữ "Kính gửi" hoặc "Kính trình" không
                string tableText = table.Range.Text;
                if (tableText.IndexOf("Kính gửi", StringComparison.OrdinalIgnoreCase) < 0 &&
                    tableText.IndexOf("Kính trình", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                // Bỏ qua bảng tiêu ngữ ở đầu trang nếu bảng đó chỉ có Quốc hiệu/Cơ quan
                if (tableText.Contains("ĐẢNG CỘNG SẢN VIỆT NAM") || tableText.Contains("CỘNG HÒA XÃ HỘI"))
                {
                    // Nếu chữ "Kính gửi" không nằm ở bảng đầu thì bỏ qua
                    if (tableText.IndexOf("Kính gửi", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                }

                try
                {
                    table.Borders.Enable = 0; // Ẩn viền bảng

                    // Tìm ô chứa danh sách Kính gửi
                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        string cellRaw = cell.Range.Text;

                        // Tách Shift + Enter (^l) thành Enter (^p) nếu người dùng gõ ngắt dòng mềm
                        Word.Find find = cell.Range.Find;
                        find.ClearFormatting();
                        find.Replacement.ClearFormatting();
                        find.Text = "^l";
                        find.Replacement.Text = "^p";
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.Execute(Replace: Word.WdReplace.wdReplaceAll);

                        // Nếu ô này có chữ "Kính gửi:"
                        if (cellRaw.IndexOf("Kính gửi:", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            cellRaw.IndexOf("Kính trình:", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            cell.Range.Font.Name = "Times New Roman";
                            cell.Range.Font.Size = 14f;
                            cell.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                            cell.Range.ParagraphFormat.SpaceBefore = 0;
                            cell.Range.ParagraphFormat.SpaceAfter = 0;
                        }

                        // Nếu ô chứa các dòng gạch đầu dòng danh sách đơn vị nhận
                        if (cellRaw.Contains("-"))
                        {
                            XuLyDanhSachDauGachKinhGui(cell.Range);
                        }
                    }
                }
                catch { }
            }

            // TRƯỜNG HỢP 2: Kính gửi viết ở dạng Paragraph thông thường ngoài Bảng
            List<Word.Paragraph> listKinhGuiParas = new List<Word.Paragraph>();
            bool dangTrongKhoiKinhGui = false;

            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                Word.Paragraph p = doc.Paragraphs[i];
                if (p.Range.Information[Word.WdInformation.wdWithInTable])
                    continue;

                string text = LamSachText(p.Range.Text);
                if (string.IsNullOrWhiteSpace(text)) continue;

                if (text.StartsWith("Kính gửi:", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("Kính trình:", StringComparison.OrdinalIgnoreCase))
                {
                    dangTrongKhoiKinhGui = true;

                    // Nếu "Kính gửi: - Đơn vị A" viết trên cùng 1 dòng
                    if (text.Contains("-"))
                    {
                        listKinhGuiParas.Add(p);
                    }
                    continue;
                }

                if (dangTrongKhoiKinhGui)
                {
                    if (text.StartsWith("-"))
                    {
                        listKinhGuiParas.Add(p);
                    }
                    else
                    {
                        // Hết khối danh sách gạch đầu dòng
                        break;
                    }
                }
            }

            if (listKinhGuiParas.Count > 0)
            {
                DinhDangDanhSachKinhGui(listKinhGuiParas);
            }
        }

        /// <summary>
        /// Xử lý các dòng gạch đầu dòng bên trong một Range (Cell)
        /// </summary>
        private void XuLyDanhSachDauGachKinhGui(Word.Range range)
        {
            List<Word.Paragraph> paras = new List<Word.Paragraph>();
            for (int i = 1; i <= range.Paragraphs.Count; i++)
            {
                Word.Paragraph p = range.Paragraphs[i];
                string text = LamSachText(p.Range.Text);
                if (text.StartsWith("-"))
                {
                    paras.Add(p);
                }
            }

            if (paras.Count > 0)
            {
                DinhDangDanhSachKinhGui(paras);
            }
        }

        /// <summary>
        /// Định dạng chuẩn cho danh sách các dòng gạch đầu dòng Kính gửi
        /// </summary>
        private void DinhDangDanhSachKinhGui(List<Word.Paragraph> paras)
        {
            int total = paras.Count;
            for (int i = 0; i < total; i++)
            {
                Word.Paragraph p = paras[i];
                string text = LamSachText(p.Range.Text);
                if (string.IsNullOrWhiteSpace(text)) continue;

                // Tách chữ "Kính gửi:" nếu nằm chung một dòng
                string tienTo = "";
                if (text.StartsWith("Kính gửi:", StringComparison.OrdinalIgnoreCase))
                {
                    tienTo = "Kính gửi: ";
                    text = text.Substring("Kính gửi:".Length).Trim();
                }

                // Làm sạch gạch đầu dòng và khoảng trắng cũ
                text = Regex.Replace(text, @"^[\s\t\-–—]+", "").Trim();

                // Xóa dấu kết thúc câu cũ
                text = Regex.Replace(text, @"[\.,;\s\t]+$", "").Trim();

                // Dòng cuối là '.', các dòng trước là ';'
                bool laDongCuoi = (i == total - 1);
                string dauKetThuc = laDongCuoi ? "." : ";";

                ThayNoiDungAnToan(p, tienTo + "- " + text + dauKetThuc);

                p.Range.Font.Name = "Times New Roman";
                p.Range.Font.Size = 14f;
                p.Range.Font.Bold = 0;
                p.Range.Font.Italic = 0;
                p.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                p.Format.SpaceBefore = 0;
                p.Format.SpaceAfter = 2f;
                MessageBox.Show(text, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnChuanHoaKinhGui_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;
            if (wordApp.Documents.Count == 0)
            {
                MessageBox.Show("Hiện không có tài liệu nào đang mở!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var doc = wordApp.ActiveDocument;
            wordApp.ScreenUpdating = false;

            try
            {
                bool daXuLy = ChuanHoaPhanKinhGui(doc, wordApp);
                if (daXuLy)
                {
                    MessageBox.Show("Đã chuẩn hóa xong phần Kính gửi!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy mục 'Kính gửi' trong tài liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }
        }

        private bool ChuanHoaPhanKinhGui(Word.Document doc, Word.Application app)
        {
            if (doc == null) return false;
            bool daXuLy = false;
            float indent11cm = app.CentimetersToPoints(4f);

            // 1. TRƯỜNG HỢP: Kính gửi nằm trong Table (Gửi nhiều nơi dùng Table)
            foreach (Word.Table table in doc.Tables)
            {
                string tableText = table.Range.Text;
                if (tableText.IndexOf("Kính gửi", StringComparison.OrdinalIgnoreCase) < 0 &&
                    tableText.IndexOf("Kính trình", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (tableText.Contains("ĐẢNG CỘNG SẢN VIỆT NAM") && tableText.IndexOf("Kính gửi", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                try
                {
                    table.Borders.Enable = 0;

                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        Word.Find find = cell.Range.Find;
                        find.ClearFormatting();
                        find.Replacement.ClearFormatting();
                        find.Text = "^l";
                        find.Replacement.Text = "^p";
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.Execute(Replace: Word.WdReplace.wdReplaceAll);

                        string cellRaw = cell.Range.Text;
                        if (cellRaw.Contains("-"))
                        {
                            DinhDangDanhSachKinhGuiNhieuNoi(cell.Range);
                            daXuLy = true;
                        }
                    }
                }
                catch { }
            }

            // 2. TRƯỜNG HỢP: Kính gửi dạng văn bản thường (Paragraphs)
            int totalParas = doc.Paragraphs.Count;

            for (int i = 1; i <= totalParas; i++)
            {
                Word.Paragraph p = doc.Paragraphs[i];
                if (p.Range.Information[Word.WdInformation.wdWithInTable])
                    continue;

                string text = LamSachText(p.Range.Text);
                if (string.IsNullOrWhiteSpace(text)) continue;

                if (text.StartsWith("Kính gửi:", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("Kính gửi", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("Kính trình:", StringComparison.OrdinalIgnoreCase))
                {
                    // A. Kính gửi 1 nơi (như "Kính gửi: Văn phòng Thành ủy...")
                    if (!text.Contains("-"))
                    {
                        p.Format.LeftIndent = 0;
                        p.Format.RightIndent = 0;
                        p.Format.FirstLineIndent = indent11cm;
                        p.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;

                        p.Range.Font.Name = "Times New Roman";
                        p.Range.Font.Bold = 0;
                        p.Range.Font.Size = 14f;
                        MessageBox.Show(text, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // In đậm riêng cụm từ "Kính gửi:"
                        Word.Range r = p.Range.Duplicate;
                        Word.Find findKG = r.Find;
                        findKG.ClearFormatting();
                        findKG.Text = "Kính gửi:";
                        if (findKG.Execute())
                        {
                            r.Font.Bold = 0;
                        }

                        daXuLy = true;
                        break;
                    }
                    else
                    {
                        // B. Kính gửi nhiều nơi gạch đầu dòng ngoài Table
                        List<Word.Paragraph> listKinhGui = new List<Word.Paragraph> { p };

                        for (int j = i + 1; j <= totalParas; j++)
                        {
                            Word.Paragraph nextP = doc.Paragraphs[j];
                            string nextText = LamSachText(nextP.Range.Text);
                            if (nextText.StartsWith("-"))
                            {
                                listKinhGui.Add(nextP);
                            }
                            else if (!string.IsNullOrWhiteSpace(nextText))
                            {
                                break;
                            }
                        }

                        DinhDangDanhSachKinhGuiNhieuNoi(listKinhGui, indent11cm);
                        daXuLy = true;
                        break;
                    }
                }
            }

            return daXuLy;
        }

        private void DinhDangDanhSachKinhGuiNhieuNoi(List<Word.Paragraph> paras, float indent11cm)
        {
            int total = paras.Count;
            for (int i = 0; i < total; i++)
            {
                Word.Paragraph p = paras[i];
                string text = LamSachText(p.Range.Text);
                if (string.IsNullOrWhiteSpace(text)) continue;

                string tienTo = "";
                if (text.StartsWith("Kính gửi:", StringComparison.OrdinalIgnoreCase))
                {
                    tienTo = "Kính gửi: ";
                    text = text.Substring("Kính gửi:".Length).Trim();
                }

                text = Regex.Replace(text, @"^[\s\t\-–—]+", "").Trim();
                text = Regex.Replace(text, @"[\.,;\s\t]+$", "").Trim();
                

                bool laDongCuoi = (i == total - 1);
                string dauKetThuc = laDongCuoi ? "." : ";";

                ThayNoiDungAnToan(p, tienTo + "- " + text + dauKetThuc);

                p.Format.LeftIndent = 0;
                p.Format.FirstLineIndent = indent11cm;
                p.Range.Font.Name = "Times New Roman";
                p.Range.Font.Bold = 0;
                p.Range.Font.Size = 14f;
            }
        }

        private void DinhDangDanhSachKinhGuiNhieuNoi(Word.Range cellRange)
        {
            List<Word.Paragraph> paras = new List<Word.Paragraph>();
            for (int i = 1; i <= cellRange.Paragraphs.Count; i++)
            {
                Word.Paragraph p = cellRange.Paragraphs[i];
                string text = LamSachText(p.Range.Text);
                if (text.StartsWith("-")) paras.Add(p);
            }

            int total = paras.Count;
            for (int i = 0; i < total; i++)
            {
                Word.Paragraph p = paras[i];
                string text = LamSachText(p.Range.Text);
                text = Regex.Replace(text, @"^[\s\t\-–—]+", "").Trim();
                text = Regex.Replace(text, @"[\.,;\s\t]+$", "").Trim();

                bool laDongCuoi = (i == total - 1);
                string dauKetThuc = laDongCuoi ? "." : ";";
               
                ThayNoiDungAnToan(p, "- " + text + dauKetThuc);
                p.Range.Font.Name = "Times New Roman";
                p.Range.Font.Bold = 0;
                p.Range.Font.Size = 14f;
            }
        }
    }
}
