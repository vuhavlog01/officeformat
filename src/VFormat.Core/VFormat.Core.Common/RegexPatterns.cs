using System.Text.RegularExpressions;

namespace VFormat.Core.Common;

public static class RegexPatterns
{
	private const RegexOptions CommonOptions = RegexOptions.IgnoreCase | RegexOptions.Compiled;

	public static readonly Regex NationalHeader = new Regex("CỘNG\\s+HÒA\\s+XÃ\\s+HỘI\\s+CHỦ\\s+NGHĨA\\s+VIỆT\\s+NAM", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex Motto = new Regex("Độc\\s+lập\\s*[-–—]\\s*Tự\\s+do\\s*[-–—]\\s*Hạnh\\s+phúc", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex PartyHeader = new Regex("ĐẢNG\\s+CỘNG\\s+SẢN\\s+VIỆT\\s+NAM", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex YouthUnionHeader = new Regex("ĐOÀN\\s+(?:TNCS|THANH\\s+NIÊN\\s+CỘNG\\s+SẢN)\\s+HỒ\\s+CHÍ\\s+MINH", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex MilitaryAgency = new Regex("(BỘ\\s+QUỐC\\s+PHÒNG|BỘ\\s+TỔNG\\s+THAM\\s+MƯU|TỔNG\\s+CỤC\\s+(?:CHÍNH\\s+TRỊ|HẬU\\s+CẦN|KỸ\\s+THUẬT|CÔNG\\s+NGHIỆP\\s+QUỐC\\s+PHÒNG|TÌNH\\s+BÁO|2)|QUÂN\\s+ĐỘI\\s+NHÂN\\s+DÂN|QUÂN\\s+KHU|QUÂN\\s+ĐOÀN|QUÂN\\s+CHỦNG|BỘ\\s+TƯ\\s+LỆNH|CỤC\\s+(?:TÁC\\s+CHIẾN|QUÂN\\s+HUẤN|QUÂN\\s+LỰC|DÂN\\s+QUÂN\\s+TỰ\\s+VỆ|CÁN\\s+BỘ|CHÍNH\\s+SÁCH|TUYÊN\\s+HUẤN|DOANH\\s+TRẠI|QUÂN\\s+NHU|QUÂN\\s+Y|VẬN\\s+TẢI|QUÂN\\s+KHÍ|XE\\s*[-–—]\\s*MÁY|BẢN\\s+ĐỒ|CỨU\\s+HỘ\\s*[-–—]\\s*CỨU\\s+NẠN|KỸ\\s+THUẬT\\s+BINH\\s+CHỦNG|ĐIỀU\\s+TRA\\s+HÌNH\\s+SỰ\\s+BQP|TÀI\\s+CHÍNH\\s+BQP|KHOA\\s+HỌC\\s+QUÂN\\s+SỰ)|BỘ\\s+CHỈ\\s+HUY\\s+QUÂN\\s+SỰ|BỘ\\s+CHQS|BAN\\s+CHỈ\\s+HUY\\s+QUÂN\\s+SỰ|BAN\\s+CHQS|BCHQS|BCH\\s+QUÂN\\s+SỰ|CHQS|BỘ\\s+ĐỘI\\s+BIÊN\\s+PHÒNG|BỘ\\s+CHỈ\\s+HUY\\s+BIÊN\\s+PHÒNG|BCHBĐBP|BĐBP|ĐỒN\\s+BIÊN\\s+PHÒNG|SƯ\\s+ĐOÀN|LỮ\\s+ĐOÀN|TRUNG\\s+ĐOÀN|TIỂU\\s+ĐOÀN|ĐẠI\\s+ĐỘI|HỌC\\s+VIỆN\\s+QUÂN\\s+SỰ|TRƯỜNG\\s+QUÂN\\s+SỰ|BAN\\s+CHỈ\\s+HUY\\s+BIÊN\\s+PHÒNG|LLVT|LỰC\\s+LƯỢNG\\s+VŨ\\s+TRANG|DÂN\\s+QUÂN\\s+TỰ\\s+VỆ|DQTV)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex PartyAgency = new Regex("(ĐẢNG\\s+ỦY|TỈNH\\s+ỦY|THÀNH\\s+ỦY|HUYỆN\\s+ỦY|QUẬN\\s+ỦY|THỊ\\s+ỦY|ĐẢNG\\s+BỘ|CHI\\s+BỘ|CHI\\s+ỦY|BAN\\s+CHẤP\\s+HÀNH|BAN\\s+THƯỜNG\\s+VỤ|BAN\\s+BÍ\\s+THƯ|VĂN\\s+PHÒNG\\s+TRUNG\\s+ƯƠNG\\s+ĐẢNG|BAN\\s+TỔ\\s+CHỨC|BAN\\s+TUYÊN\\s+GIÁO|ỦY\\s+BAN\\s+KIỂM\\s+TRA)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex YouthUnionAgency = new Regex("(?:\\bHỘI\\s+ĐỒNG\\s+ĐỘI\\b|\\bLIÊN\\s+ĐỘI\\b|\\bCHI\\s+ĐỘI\\b|\\bBCH\\s+(?:LIÊN\\s+CHI\\s+)?ĐOÀN\\b|\\bBAN\\s+CHẤP\\s+HÀNH\\s+(?:LIÊN\\s+CHI\\s+)?ĐOÀN\\b|(?<![\\p{L}])(?:TỈNH|THÀNH|HUYỆN|QUẬN|THỊ)\\s+ĐOÀN\\b|\\b(?:LIÊN\\s+)?CHI\\s+ĐOÀN\\b|\\bPHÂN\\s+ĐOÀN\\b|\\bĐOÀN\\s+(?:TRƯỜNG|KHỐI|XÃ|PHƯỜNG|THỊ\\s+TRẤN|CƠ\\s+SỞ|TNCS|THANH\\s+NIÊN|TỔNG\\s+CÔNG\\s+TY|CÔNG\\s+TY|ĐẠI\\s+HỌC|HỌC\\s+VIỆN)\\b|\\bĐẠI\\s+HỘI\\s+(?:ĐẠI\\s+BIỂU\\s+(?:TOÀN\\s+QUỐC\\s+ĐOÀN|ĐOÀN)|ĐOÀN|CHI\\s+ĐOÀN)\\b)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex AdminAgency = new Regex("(ỦY\\s+BAN\\s+NHÂN\\s+DÂN|UBND|HỘI\\s+ĐỒNG\\s+NHÂN\\s+DÂN|HĐND|CHÍNH\\s+PHỦ|THỦ\\s+TƯỚNG|BỘ\\s+[\\p{L}\\s]+|TỔNG\\s+CỤC(?:\\s+[\\p{L}\\s]+)?|CỤC\\s+[\\p{L}\\s]+|SỞ\\s+[\\p{L}\\s]+|PHÒNG\\s+[\\p{L}\\s]+|TRƯỜNG\\s+[\\p{L}\\s]+|VIỆN\\s+[\\p{L}\\s]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex DocumentNumberGeneral = new Regex("^(?:Số|So)\\s*[:\\s]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex AdminNumber = new Regex("Số\\s*:\\s*(?:\\d{1,4}|\\.{2,}|_{2,}|\\s*)/[A-ZĐ0-9-]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex PartyNumber = new Regex("Số\\s*:?\\s*(?:\\d{1,4}|\\.{2,}|_{2,}|\\s*)-\\s*[A-ZĐ0-9-/]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex MilitaryNumber = new Regex("Số\\s*:\\s*(?:\\d{1,4}|\\.{2,}|_{2,}|\\s*)/[A-ZĐ0-9-]+(?:/[A-ZĐ0-9-]+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex YouthUnionNumberSymbol = new Regex("[/-]\\s*(?:TWĐTN|UBKTTWĐ|TĐTN|ĐTNK|ĐTNCA|BCSĐ|ĐTN|CĐCS)(?![\\p{L}\\d])", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex MilitaryNumberSymbol = new Regex("[/-](BCH|BQP|TM|CT|HC|KT|BTL|QK|SĐ|LĐ|e\\d+|BCHQS|BĐBP)\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex DocumentDate = new Regex("^[\\w\\s\\.\\-]+,\\s*ngày\\s+\\d{1,2}\\s+tháng\\s+\\d{1,2}\\s+năm\\s+\\d{4}\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex DocumentType = new Regex("^\\s*(NGHỊ\\s+QUYẾT|QUYẾT\\s+ĐỊNH|CHỈ\\s+THỊ|QUY\\s+CHẾ|QUY\\s+ĐỊNH|THÔNG\\s+CÁO|THÔNG\\s+BÁO|HƯỚNG\\s+DẪN|CHƯƠNG\\s+TRÌNH|KẾ\\s+HOẠCH|PHƯƠNG\\s+ÁN|ĐỀ\\s+ÁN|DỰ\\s+ÁN|BÁO\\s+CÁO|BIÊN\\s+BẢN|TỜ\\s+TRÌNH|HỢP\\s+ĐỒNG|CÔNG\\s+VĂN|CÔNG\\s+ĐIỆN|BẢN\\s+ĐĂNG\\s+KÝ|BẢN\\s+CAM\\s+KẾT|BẢN\\s+TỰ\\s+KIỂM\\s+ĐIỂM|BẢN\\s+KIỂM\\s+ĐIỂM|BẢN\\s+TỰ\\s+NHẬN\\s+XÉT|BẢN\\s+GHI\\s+NHỚ|BẢN\\s+THỎA\\s+THUẬN|BẢN\\s+TƯỜNG\\s+TRÌNH|BẢN\\s+KHAI|GIẤY\\s+ỦY\\s+QUYỀN|GIẤY\\s+MỜI|GIẤY\\s+GIỚI\\s+THIỆU|GIẤY\\s+XÁC\\s+NHẬN|GIẤY\\s+ĐI\\s+ĐƯỜNG|GIẤY\\s+NGHỈ\\s+PHÉP|PHIẾU\\s+ĐÁNH\\s+GIÁ|PHIẾU\\s+TRÌNH|PHIẾU\\s+XIN\\s+Ý\\s+KIẾN|PHIẾU\\s+GỬI|PHIẾU\\s+CHUYỂN|PHIẾU\\s+BÁO|THƯ\\s+CÔNG|THÔNG\\s+TRI|ĐƠN)(?::|\\s+.*)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex DispatchSubject = new Regex("^\\s*V/v\\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex LegalGround = new Regex("^\\s*Căn\\s+cứ\\s+(?:vào\\s+)?(Luật|Bộ\\s+luật|Pháp\\s+lệnh|Nghị\\s+định|Nghị\\s+quyết|Quyết\\s+định|Thông\\s+tư|Chỉ\\s+thị|Quy\\s+chế|Quy\\s+định|Kế\\s+hoạch|Phương\\s+án|Đề\\s+án|Dự\\s+án|Báo\\s+cáo|Tờ\\s+trình|Hướng\\s+dẫn|Thông\\s+báo|Công\\s+văn|Công\\s+điện|Kết\\s+luận|Biên\\s+bản|Ý\\s+kiến|Văn\\s+bản|Thỏa\\s+thuận|Hợp\\s+đồng|Điều\\s+lệ|Quy\\s+ước|chức\\s+năng|nhiệm\\s+vụ|quyền\\s+hạn|đề\\s+nghị|xét\\s+đề\\s+nghị)\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex PartHeader = new Regex("^\\s*PHẦN\\s+(THỨ\\s+)?[IVXLCDM\\d]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex ChapterHeader = new Regex("^\\s*Chương\\s+[IVXLCDM\\d]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex SectionHeader = new Regex("^\\s*(Tiểu\\s+mục|Mục)\\s+\\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex ArticleTitle = new Regex("^\\s*Điều\\s+\\d+\\.?\\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex ClauseNumber = new Regex("^\\s*\\d+\\.\\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex PointLetter = new Regex("^\\s*[a-zđ]\\)\\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex DirectRecipient = new Regex("^\\s*Kính\\s+gửi\\s*:?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex RecipientHeader = new Regex("^\\s*Nơi\\s+nhận\\s*:?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex RecipientItem = new Regex("^\\s*-\\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex SignAuthority = new Regex("^\\s*(TM\\.|KT\\.|TL\\.|TUQ\\.|Q\\.\\s+[A-ZÀ-Ỹ]|T/M|K/T|T/L)\\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex SignPosition = new Regex("^\\s*(CHỦ\\s+TỊCH|PHÓ\\s+CHỦ\\s+TỊCH|GIÁM\\s+ĐỐC|PHÓ\\s+GIÁM\\s+ĐỐC|TỔNG\\s+GIÁM\\s+ĐỐC|PHÓ\\s+TỔNG\\s+GIÁM\\s+ĐỐC|BỘ\\s+TRƯỞNG|THỨ\\s+TRƯỞNG|BÍ\\s+THƯ|PHÓ\\s+BÍ\\s+THƯ|CHÁNH\\s+VĂN\\s+PHÒNG|PHÓ\\s+CHÁNH\\s+VĂN\\s+PHÒNG|TRƯỞNG\\s+PHÒNG|PHÓ\\s+TRƯỞNG\\s+PHÒNG|TRƯỞNG\\s+BAN|PHÓ\\s+TRƯỞNG\\s+BAN|CHỦ\\s+NHIỆM|PHÓ\\s+CHỦ\\s+NHIỆM|TƯ\\s+LỆNH|PHÓ\\s+TƯ\\s+LỆNH|CHÍNH\\s+ỦY|PHÓ\\s+CHÍNH\\s+ỦY|CHỈ\\s+HUY\\s+TRƯỞNG|CHỈ\\s+HUY\\s+PHÓ|PHÓ\\s+CHỈ\\s+HUY\\s+TRƯỞNG|CHÍNH\\s+TRỊ\\s+VIÊN|PHÓ\\s+CHÍNH\\s+TRỊ\\s+VIÊN|THAM\\s+MƯU\\s+TRƯỞNG|PHÓ\\s+THAM\\s+MƯU\\s+TRƯỞNG|CỤC\\s+TRƯỞNG|VIỆN\\s+TRƯỞNG|HIỆU\\s+TRƯỞNG|TRƯỞNG\\s+ĐOÀN|PHÓ\\s+TRƯỞNG\\s+ĐOÀN)\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex MilitaryRank = new Regex("\\b(Đại\\s+tướng|Thượng\\s+tướng|Trung\\s+tướng|Thiếu\\s+tướng|Đại\\s+tá|Thượng\\s+tá|Trung\\s+tá|Thiếu\\s+tá|Đại\\s+úy|Thượng\\s+úy|Trung\\s+úy|Thiếu\\s+úy)\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex UrgencyIndicator = new Regex("^\\s*(KHẨN|THƯỢNG\\s+KHẨN|HỎA\\s+TỐC)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex SecurityIndicator = new Regex("^\\s*(MẬT|TỐI\\s+MẬT|TUYỆT\\s+MẬT)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex CirculationScope = new Regex("^\\s*(LƯU\\s+HÀNH\\s+NỘI\\s+BỘ|XEM\\s+XONG\\s+TRẢ\\s+LẠI)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex DraftIndicator = new Regex("^\\s*(DỰ\\s+THẢO|DỰ\\s+THẢO\\s+LẦN\\s+(?:THỨ\\s+)?\\d+)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex AppendixTitle = new Regex("^\\s*Phụ\\s+lục\\s+([IVXLCDM\\d]+)?", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static readonly Regex CopyPaper = new Regex("^\\s*(SAO\\s+Y|SAO\\s+LỤC|TRÍCH\\s+SAO)\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	public static bool IsStandaloneSignPosition(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		string text2 = text.Trim();
		Match match = SignPosition.Match(text2);
		if (!match.Success)
		{
			return false;
		}
		if (match.Value.Trim().Length == text2.Length)
		{
			return true;
		}
		if (text2.Length <= 100 && text2 == text2.ToUpperInvariant())
		{
			return !Regex.IsMatch(text2, "[.;:!?\\r\\n\\v\\a]");
		}
		return false;
	}
}
