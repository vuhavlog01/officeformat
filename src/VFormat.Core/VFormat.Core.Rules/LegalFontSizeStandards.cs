using System.Collections.Generic;
using VFormat.Core.Enums;

namespace VFormat.Core.Rules;

public static class LegalFontSizeStandards
{
	private static readonly Dictionary<SemanticRole, LegalFontSizeRange> ND30Ranges = new Dictionary<SemanticRole, LegalFontSizeRange>
	{
		{
			SemanticRole.NationalHeader,
			new LegalFontSizeRange(12, 13, 12, "Quốc hiệu NĐ 30: cỡ 12-13, đứng đậm")
		},
		{
			SemanticRole.Motto,
			new LegalFontSizeRange(13, 14, 13, "Tiêu ngữ NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.SuperiorAgency,
			new LegalFontSizeRange(12, 13, 13, "Cơ quan cấp trên NĐ 30: cỡ 12-13, đứng thường")
		},
		{
			SemanticRole.IssuingAgency,
			new LegalFontSizeRange(12, 13, 13, "Cơ quan ban hành NĐ 30: cỡ 12-13, đứng đậm")
		},
		{
			SemanticRole.DocumentNumber,
			new LegalFontSizeRange(13, 13, 13, "Số ký hiệu NĐ 30: cỡ 13, đứng thường")
		},
		{
			SemanticRole.DocumentDate,
			new LegalFontSizeRange(13, 14, 14, "Địa danh ngày tháng NĐ 30: cỡ 13-14, nghiêng")
		},
		{
			SemanticRole.DocumentType,
			new LegalFontSizeRange(13, 14, 14, "Tên loại văn bản NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.SubjectExcerpt,
			new LegalFontSizeRange(13, 14, 14, "Trích yếu văn bản NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.DispatchSubject,
			new LegalFontSizeRange(12, 13, 12, "Trích yếu công văn NĐ 30: cỡ 12-13, đứng thường")
		},
		{
			SemanticRole.LegalGround,
			new LegalFontSizeRange(13, 14, 14, "Căn cứ pháp lý NĐ 30: cỡ 13-14, nghiêng")
		},
		{
			SemanticRole.BodyText,
			new LegalFontSizeRange(13, 14, 14, "Nội dung văn bản NĐ 30: cỡ 13-14, đứng thường")
		},
		{
			SemanticRole.PartHeader,
			new LegalFontSizeRange(13, 14, 14, "Phần NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.ChapterHeader,
			new LegalFontSizeRange(13, 14, 14, "Chương NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.SectionHeader,
			new LegalFontSizeRange(13, 14, 14, "Mục NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.ArticleTitle,
			new LegalFontSizeRange(13, 14, 14, "Điều NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.ClauseText,
			new LegalFontSizeRange(13, 14, 14, "Khoản NĐ 30: cỡ 13-14, đứng")
		},
		{
			SemanticRole.PointText,
			new LegalFontSizeRange(13, 14, 14, "Điểm NĐ 30: cỡ 13-14, đứng")
		},
		{
			SemanticRole.SignAuthority,
			new LegalFontSizeRange(13, 14, 14, "Quyền hạn ký NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.SignTitle,
			new LegalFontSizeRange(13, 14, 14, "Chức vụ người ký NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.SignerName,
			new LegalFontSizeRange(13, 14, 14, "Họ tên người ký NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.DirectRecipient,
			new LegalFontSizeRange(13, 14, 14, "Kính gửi công văn NĐ 30: cỡ 13-14, đứng")
		},
		{
			SemanticRole.RecipientHeader,
			new LegalFontSizeRange(12, 12, 12, "Từ 'Nơi nhận:' NĐ 30: cỡ 12, nghiêng đậm")
		},
		{
			SemanticRole.RecipientItem,
			new LegalFontSizeRange(11, 11, 11, "Danh mục nơi nhận NĐ 30: cỡ 11, đứng thường")
		},
		{
			SemanticRole.UrgencyIndicator,
			new LegalFontSizeRange(13, 14, 13, "Độ khẩn NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.SecurityIndicator,
			new LegalFontSizeRange(13, 14, 14, "Chữ chỉ mức độ mật: mặc định ứng dụng 14; mẫu dấu theo pháp luật bảo mật riêng")
		},
		{
			SemanticRole.CirculationScope,
			new LegalFontSizeRange(13, 14, 13, "Phạm vi lưu hành NĐ 30: cỡ 13-14, đứng đậm")
		},
		{
			SemanticRole.CreatorSymbol,
			new LegalFontSizeRange(11, 11, 11, "Ký hiệu người soạn thảo NĐ 30: cỡ 11, đứng")
		},
		{
			SemanticRole.ContactInfo,
			new LegalFontSizeRange(11, 12, 11, "Thông tin liên hệ NĐ 30: cỡ 11-12, đứng")
		}
	};

	private static readonly Dictionary<SemanticRole, LegalFontSizeRange> HD05Ranges = new Dictionary<SemanticRole, LegalFontSizeRange>
	{
		{
			SemanticRole.PartyHeader,
			new LegalFontSizeRange(15, 15, 15, "Đảng hiệu HĐ 05: cỡ 15, đứng đậm")
		},
		{
			SemanticRole.SuperiorAgency,
			new LegalFontSizeRange(14, 14, 14, "Cơ quan cấp trên HĐ 05: cỡ 14, đứng thường")
		},
		{
			SemanticRole.IssuingAgency,
			new LegalFontSizeRange(14, 14, 14, "Cơ quan ban hành HĐ 05: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.DocumentNumber,
			new LegalFontSizeRange(14, 14, 14, "Số ký hiệu HĐ 05: cỡ 14, đứng thường")
		},
		{
			SemanticRole.DocumentDate,
			new LegalFontSizeRange(14, 14, 14, "Địa danh ngày tháng HĐ 05: cỡ 14, nghiêng")
		},
		{
			SemanticRole.DocumentType,
			new LegalFontSizeRange(15, 16, 15, "Tên loại văn bản HĐ 05: cỡ 15-16, đứng đậm")
		},
		{
			SemanticRole.SubjectExcerpt,
			new LegalFontSizeRange(14, 15, 14, "Trích yếu văn bản HĐ 05: cỡ 14-15, đứng đậm")
		},
		{
			SemanticRole.DispatchSubject,
			new LegalFontSizeRange(12, 12, 12, "Trích yếu công văn HĐ 05: cỡ 12, nghiêng")
		},
		{
			SemanticRole.LegalGround,
			new LegalFontSizeRange(14, 15, 14, "Căn cứ HĐ 05: theo nội dung 14-15; không bắt buộc nghiêng")
		},
		{
			SemanticRole.BodyText,
			new LegalFontSizeRange(14, 15, 14, "Nội dung văn bản HĐ 05: cỡ 14-15, đứng thường")
		},
		{
			SemanticRole.PartHeader,
			new LegalFontSizeRange(14, 15, 14, "Phần HĐ 05: cỡ 14-15, đứng đậm")
		},
		{
			SemanticRole.ChapterHeader,
			new LegalFontSizeRange(14, 15, 14, "Chương HĐ 05: cỡ 14-15, đứng đậm")
		},
		{
			SemanticRole.SectionHeader,
			new LegalFontSizeRange(14, 15, 14, "Mục HĐ 05: cỡ 14-15, đứng đậm")
		},
		{
			SemanticRole.ArticleTitle,
			new LegalFontSizeRange(14, 15, 14, "Điều HĐ 05: cỡ 14-15, đứng đậm")
		},
		{
			SemanticRole.ClauseText,
			new LegalFontSizeRange(14, 15, 14, "Khoản HĐ 05: cỡ 14-15, đứng")
		},
		{
			SemanticRole.PointText,
			new LegalFontSizeRange(14, 15, 14, "Điểm HĐ 05: cỡ 14-15, đứng")
		},
		{
			SemanticRole.SignAuthority,
			new LegalFontSizeRange(14, 14, 14, "Quyền hạn ký HĐ 05: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SignTitle,
			new LegalFontSizeRange(14, 14, 14, "Chức vụ người ký HĐ 05: cỡ 14, đứng thường")
		},
		{
			SemanticRole.SignerName,
			new LegalFontSizeRange(14, 14, 14, "Họ tên người ký HĐ 05: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.DirectRecipient,
			new LegalFontSizeRange(14, 15, 14, "Kính gửi HĐ 05: cỡ 14-15")
		},
		{
			SemanticRole.RecipientHeader,
			new LegalFontSizeRange(14, 14, 14, "Từ 'Nơi nhận:' HĐ 05: cỡ 14, đứng gạch chân")
		},
		{
			SemanticRole.RecipientItem,
			new LegalFontSizeRange(12, 12, 12, "Danh mục nơi nhận HĐ 05: cỡ 12, đứng thường")
		},
		{
			SemanticRole.UrgencyIndicator,
			new LegalFontSizeRange(14, 14, 14, "Độ khẩn HĐ 05: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SecurityIndicator,
			new LegalFontSizeRange(14, 14, 14, "Chữ chỉ mức độ mật: mặc định ứng dụng 14; HD05 không quy định riêng cỡ dấu mật")
		},
		{
			SemanticRole.CirculationScope,
			new LegalFontSizeRange(12, 12, 12, "Lưu hành HĐ 05: cỡ 12, đứng đậm")
		},
		{
			SemanticRole.DraftIndicator,
			new LegalFontSizeRange(14, 14, 14, "Dự thảo HĐ 05: cỡ 14, nghiêng")
		},
		{
			SemanticRole.CreatorSymbol,
			new LegalFontSizeRange(11, 11, 11, "Ký hiệu người soạn thảo HĐ 05: cỡ 11, đứng")
		},
		{
			SemanticRole.ContactInfo,
			new LegalFontSizeRange(10, 10, 10, "Thông tin liên hệ HĐ 05: cỡ 10, đứng")
		}
	};

	private static readonly Dictionary<SemanticRole, LegalFontSizeRange> TT68Ranges = new Dictionary<SemanticRole, LegalFontSizeRange>
	{
		{
			SemanticRole.NationalHeader,
			new LegalFontSizeRange(13, 13, 13, "Quốc hiệu TT 68: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.Motto,
			new LegalFontSizeRange(13, 13, 13, "Tiêu ngữ TT 68: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.SuperiorAgency,
			new LegalFontSizeRange(13, 13, 13, "Cơ quan cấp trên TT 68: cỡ 13, đứng thường")
		},
		{
			SemanticRole.IssuingAgency,
			new LegalFontSizeRange(13, 13, 13, "Cơ quan ban hành TT 68: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.DocumentNumber,
			new LegalFontSizeRange(13, 13, 13, "Số ký hiệu TT 68: cỡ 13, đứng thường")
		},
		{
			SemanticRole.DocumentDate,
			new LegalFontSizeRange(13, 13, 13, "Địa danh ngày tháng TT 68: cỡ 13, nghiêng")
		},
		{
			SemanticRole.DocumentType,
			new LegalFontSizeRange(14, 14, 14, "Tên loại văn bản TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SubjectExcerpt,
			new LegalFontSizeRange(14, 14, 14, "Trích yếu văn bản TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.DispatchSubject,
			new LegalFontSizeRange(12, 12, 12, "Trích yếu công văn TT 68: cỡ 12, đứng thường")
		},
		{
			SemanticRole.LegalGround,
			new LegalFontSizeRange(14, 14, 14, "Căn cứ ban hành TT 68: cỡ 14")
		},
		{
			SemanticRole.BodyText,
			new LegalFontSizeRange(14, 14, 14, "Nội dung văn bản TT 68: cỡ 14, đứng thường")
		},
		{
			SemanticRole.PartHeader,
			new LegalFontSizeRange(14, 14, 14, "Phần TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.ChapterHeader,
			new LegalFontSizeRange(14, 14, 14, "Chương TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SectionHeader,
			new LegalFontSizeRange(14, 14, 14, "Mục TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.ArticleTitle,
			new LegalFontSizeRange(14, 14, 14, "Điều TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.ClauseText,
			new LegalFontSizeRange(14, 14, 14, "Khoản TT 68: cỡ 14, đứng")
		},
		{
			SemanticRole.PointText,
			new LegalFontSizeRange(14, 14, 14, "Điểm TT 68: cỡ 14, đứng")
		},
		{
			SemanticRole.SignAuthority,
			new LegalFontSizeRange(14, 14, 14, "Quyền hạn ký TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SignTitle,
			new LegalFontSizeRange(14, 14, 14, "Chức vụ người ký TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SignRankName,
			new LegalFontSizeRange(14, 14, 14, "Quân hàm họ tên TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SignerName,
			new LegalFontSizeRange(14, 14, 14, "Họ tên người ký TT 68: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.DirectRecipient,
			new LegalFontSizeRange(14, 14, 14, "Kính gửi TT 68: cỡ 14, đứng")
		},
		{
			SemanticRole.RecipientHeader,
			new LegalFontSizeRange(12, 12, 12, "Từ 'Nơi nhận:' TT 68: cỡ 12, nghiêng đậm")
		},
		{
			SemanticRole.RecipientItem,
			new LegalFontSizeRange(11, 11, 11, "Danh mục nơi nhận TT 68: cỡ 11, đứng thường")
		},
		{
			SemanticRole.UrgencyIndicator,
			new LegalFontSizeRange(13, 13, 13, "Độ khẩn TT 68: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.SecurityIndicator,
			new LegalFontSizeRange(13, 13, 13, "Chữ chỉ mức độ mật: mặc định ứng dụng 13; mẫu dấu theo pháp luật bảo mật riêng")
		},
		{
			SemanticRole.CirculationScope,
			new LegalFontSizeRange(13, 13, 13, "Lưu hành TT 68: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.DraftIndicator,
			new LegalFontSizeRange(13, 13, 13, "Dự thảo TT 68: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.CreatorSymbol,
			new LegalFontSizeRange(11, 11, 11, "Ký hiệu người soạn thảo TT 68: cỡ 11, đứng")
		},
		{
			SemanticRole.ContactInfo,
			new LegalFontSizeRange(12, 12, 12, "Thông tin liên hệ TT 68: cỡ 12, đứng")
		},
		{
			SemanticRole.CopyElectronic,
			new LegalFontSizeRange(10, 10, 10, "Ký số văn bản số hóa TT 68: cỡ 10, đứng")
		}
	};

	private static readonly Dictionary<SemanticRole, LegalFontSizeRange> HD29Ranges = new Dictionary<SemanticRole, LegalFontSizeRange>
	{
		{
			SemanticRole.PartyHeader,
			new LegalFontSizeRange(15, 15, 15, "Tiêu đề Đoàn HD 29: cỡ 15, đứng đậm, có kẻ ngang")
		},
		{
			SemanticRole.SuperiorAgency,
			new LegalFontSizeRange(14, 14, 14, "Tổ chức Đoàn cấp trên HD 29: cỡ 14, đứng")
		},
		{
			SemanticRole.IssuingAgency,
			new LegalFontSizeRange(14, 14, 14, "Cơ quan ban hành HD 29: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.DocumentNumber,
			new LegalFontSizeRange(14, 14, 14, "Số ký hiệu HD 29: cỡ 14, đứng")
		},
		{
			SemanticRole.DocumentDate,
			new LegalFontSizeRange(13, 13, 13, "Địa điểm, ngày tháng HD 29: cỡ 13, nghiêng")
		},
		{
			SemanticRole.DispatchSubject,
			new LegalFontSizeRange(12, 12, 12, "Trích yếu công văn HD 29: cỡ 12, nghiêng")
		},
		{
			SemanticRole.DocumentType,
			new LegalFontSizeRange(16, 16, 16, "Tên loại văn bản HD 29: cỡ 16, đứng đậm")
		},
		{
			SemanticRole.SubjectExcerpt,
			new LegalFontSizeRange(14, 14, 14, "Trích yếu HD 29: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.LegalGround,
			new LegalFontSizeRange(14, 14, 14, "Căn cứ HD 29: cỡ nội dung 14, đứng")
		},
		{
			SemanticRole.BodyText,
			new LegalFontSizeRange(14, 14, 14, "Nội dung văn bản HD 29: cỡ 14, đứng")
		},
		{
			SemanticRole.PartHeader,
			new LegalFontSizeRange(15, 16, 15, "Phần HD 29: cỡ 15-16, đứng đậm")
		},
		{
			SemanticRole.ChapterHeader,
			new LegalFontSizeRange(14, 14, 14, "Tiêu đề phần HD 29: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.SectionHeader,
			new LegalFontSizeRange(14, 14, 14, "Mục HD 29: cỡ 14, in hoa đứng đậm")
		},
		{
			SemanticRole.ArticleTitle,
			new LegalFontSizeRange(14, 14, 14, "Khoản có tiêu đề HD 29: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.ClauseText,
			new LegalFontSizeRange(14, 14, 14, "Khoản HD 29: cỡ 14")
		},
		{
			SemanticRole.PointText,
			new LegalFontSizeRange(14, 14, 14, "Điểm HD 29: cỡ 14")
		},
		{
			SemanticRole.SignAuthority,
			new LegalFontSizeRange(14, 14, 14, "Thể thức đề ký HD 29: cỡ 14, in hoa đứng đậm")
		},
		{
			SemanticRole.SignTitle,
			new LegalFontSizeRange(14, 14, 14, "Chức vụ người ký HD 29: cỡ 14, in hoa đứng")
		},
		{
			SemanticRole.SignerName,
			new LegalFontSizeRange(14, 14, 14, "Họ tên người ký HD 29: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.DirectRecipient,
			new LegalFontSizeRange(14, 14, 14, "Kính gửi HD 29: cỡ 14, đứng đậm")
		},
		{
			SemanticRole.RecipientHeader,
			new LegalFontSizeRange(13, 13, 13, "Từ 'Nơi nhận:' HD 29: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.RecipientItem,
			new LegalFontSizeRange(11, 11, 11, "Danh sách nơi nhận HD 29: cỡ 11, đứng")
		},
		{
			SemanticRole.SecurityIndicator,
			new LegalFontSizeRange(13, 13, 13, "Chỉ mức độ mật HD 29: cỡ 13, đứng đậm, khung")
		},
		{
			SemanticRole.UrgencyIndicator,
			new LegalFontSizeRange(13, 13, 13, "Chỉ mức độ khẩn HD 29: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.DraftIndicator,
			new LegalFontSizeRange(13, 13, 13, "Chỉ dẫn dự thảo HD 29: cỡ 13, đứng đậm")
		},
		{
			SemanticRole.CirculationScope,
			new LegalFontSizeRange(12, 12, 12, "Chỉ dẫn tài liệu hội nghị HD 29: cỡ 12, nghiêng")
		},
		{
			SemanticRole.CreatorSymbol,
			new LegalFontSizeRange(8, 8, 8, "Ký hiệu lưu văn bản, mã số HD 29: cỡ 8, nghiêng")
		},
		{
			SemanticRole.ContactInfo,
			new LegalFontSizeRange(10, 12, 11, "Thông tin liên hệ HD 29: nguồn không quy định, mặc định ứng dụng 11")
		}
	};

	public static LegalFontSizeRange GetLegalRange(RulePackType rulePack, SemanticRole role)
	{
		if ((rulePack switch
		{
			RulePackType.HD05_VPTW => HD05Ranges, 
			RulePackType.TT68_2023_BQP => TT68Ranges, 
			RulePackType.HD29_TWDTN => HD29Ranges, 
			_ => ND30Ranges, 
		}).TryGetValue(role, out var value))
		{
			return value;
		}
		return new LegalFontSizeRange(13, 14, 14, "Mặc định: cỡ 13-14");
	}

	public static float EnsureLegalInteger(RulePackType rulePack, SemanticRole role, float size)
	{
		LegalFontSizeRange legalRange = GetLegalRange(rulePack, role);
		return legalRange.Clamp(size);
	}

	public static LegalFontSizeRange GetTableFontSizeRange(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => new LegalFontSizeRange(14, 15, 14, "Bảng biểu HD05: dùng dải nội dung 14-15 khi nguồn không quy định riêng cỡ bảng"), 
			RulePackType.TT68_2023_BQP => new LegalFontSizeRange(14, 14, 14, "Bảng biểu TT68: dùng cỡ nội dung 14 khi nguồn không quy định riêng cỡ bảng"), 
			RulePackType.HD29_TWDTN => new LegalFontSizeRange(12, 14, 14, "Bảng biểu HD 29: nguồn không quy định riêng, dùng dải 12-14 (mặc định cỡ nội dung 14)"), 
			_ => new LegalFontSizeRange(13, 14, 14, "Bảng biểu NĐ30: dùng dải nội dung 13-14 khi nguồn không quy định riêng cỡ bảng"), 
		};
	}

	public static LegalFontSizeRange GetPageNumberFontSizeRange(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => new LegalFontSizeRange(14, 14, 14, "Số trang HĐ 05: cỡ 14"), 
			RulePackType.TT68_2023_BQP => new LegalFontSizeRange(14, 14, 14, "Số trang TT 68: cỡ 14"), 
			RulePackType.HD29_TWDTN => new LegalFontSizeRange(13, 14, 14, "Số trang HD 29: cỡ 13-14, đứng, mặc định 14"), 
			_ => new LegalFontSizeRange(13, 14, 14, "Số trang NĐ 30: cỡ 13-14, mặc định 14"), 
		};
	}
}
