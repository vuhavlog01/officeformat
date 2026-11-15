using System;
using System.Collections.Generic;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.Core.Rules;

public static class HD29Rules
{
	public const string ProfileId = "PROFILE_DOAN";

	public const string Name = "Văn bản của Đoàn TNCS Hồ Chí Minh";

	public const string Source = "Hướng dẫn số 29-HD/TWĐTN-VP ngày 29/10/2013 của Ban Bí thư Trung ương Đoàn";

	public const string Description = "Thể thức văn bản của Đoàn TNCS Hồ Chí Minh từ Trung ương đến chi đoàn, phân đoàn";

	public const string HeaderTitle = "ĐOÀN TNCS HỒ CHÍ MINH";

	public const string AgencySeparator = "***";

	public const string ExcerptSeparator = "---------------";

	public const string NumberPrefix = "Số: ";

	public const string CongressSymbol = "ĐH";

	public const string CopySymbol = "BS";

	public const float FirstLineIndentMm = 12.7f;

	public const float MaxParagraphSpacingPt = 6f;

	public const float MinLineSpacingMultiple = 1f;

	public const float MaxLineSpacingMultiple = 1.2f;

	public const float MarginTopMm = 20f;

	public const float MarginBottomMm = 20f;

	public const float MarginLeftMm = 30f;

	public const float MarginRightMm = 20f;

	public const float DuplexMarginLeftMm = 20f;

	public const float DuplexMarginRightMm = 30f;

	public static readonly string[] SignAuthorityPrefixes = new string[4] { "TM.", "KT.", "TL.", "Q." };

	public static readonly string[] SecurityLevels = new string[1] { "MẬT" };

	public static readonly string[] UrgencyLevels = new string[3] { "HỎA TỐC", "THƯỢNG KHẨN", "KHẨN" };

	public static readonly string[] CopyTypes = new string[3] { "Sao nguyên văn bản chính", "Sao lục", "Trích sao từ bản chính số..... ngày.....của ....." };

	public static readonly string[] A5DocumentTypes = new string[3] { "GIẤY GIỚI THIỆU", "PHIẾU GỬI", "PHIẾU CHUYỂN" };

	public static readonly Dictionary<string, string> DocTypeAbbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
	{
		{ "Quyết định", "QĐ" },
		{ "Quy định", "QĐ" },
		{ "Chỉ thị", "CT" },
		{ "Chương trình", "CTr" },
		{ "Thông tri", "TT" },
		{ "Tờ trình", "TTr" },
		{ "Công văn", "CV" },
		{ "Báo cáo", "BC" },
		{ "Kế hoạch", "KH" },
		{ "Thông báo", "TB" },
		{ "Nghị quyết", "NQ" },
		{ "Hướng dẫn", "HD" },
		{ "Biên bản", "BB" },
		{ "Giấy giới thiệu", "GGT" }
	};

	public static readonly Dictionary<string, string> AgencyAbbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
	{
		{ "Ban Bí thư, Ban Thường vụ, Ban Chấp hành Trung ương Đoàn", "TWĐTN" },
		{ "Ủy ban Kiểm tra Trung ương Đoàn", "UBKTTWĐ" },
		{ "Tỉnh đoàn, thành đoàn thuộc Trung ương", "TĐTN" },
		{ "Đoàn khối các cơ quan Trung ương, Đoàn khối Doanh nghiệp Trung ương", "ĐTNK" },
		{ "Đoàn Thanh niên Bộ Công an", "ĐTNCA" },
		{ "Ban Cán sự Đoàn", "BCSĐ" },
		{ "Đoàn cấp quận và tương đương, Đoàn cấp cơ sở, Đoàn bộ phận", "ĐTN" },
		{ "Chi đoàn cơ sở", "CĐCS" },
		{ "Chi đoàn", "CĐ" },
		{ "Phân đoàn", "PĐ" },
		{ "Đại hội các cấp", "ĐH" }
	};

	public static readonly Dictionary<string, string> UnitAbbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
	{
		{ "Văn phòng", "VP" },
		{ "Ban Tổ chức", "BTC" },
		{ "Ban Kiểm tra", "BKT" },
		{ "Ban Tổ chức - Kiểm tra", "TCKT" },
		{ "Ban Thanh niên trường học", "TNTH" },
		{ "Ban Đoàn kết tập hợp thanh niên", "ĐKTHTN" },
		{ "Trung tâm Thanh thiếu niên Trung ương", "TTNTW" },
		{ "Tiểu ban", "TB" },
		{ "Tổ công tác", "TCT" },
		{ "Hội đồng", "HĐ" },
		{ "Đoàn kiểm tra", "ĐKT" }
	};

	public static readonly string[] AgencySymbols = new string[10] { "TWĐTN", "UBKTTWĐ", "TĐTN", "ĐTNK", "ĐTNCA", "BCSĐ", "ĐTN", "CĐCS", "CĐ", "PĐ" };

	public static ProfileModel CreateProfile()
	{
		ProfileModel profileModel = new ProfileModel("PROFILE_DOAN", RulePackType.HD29_TWDTN, "Văn bản của Đoàn TNCS Hồ Chí Minh", "Hướng dẫn số 29-HD/TWĐTN-VP ngày 29/10/2013 của Ban Bí thư Trung ương Đoàn", "Thể thức văn bản của Đoàn TNCS Hồ Chí Minh từ Trung ương đến chi đoàn, phân đoàn");
		profileModel.Typography.FontName = "Times New Roman";
		profileModel.Typography.BodySizePt = 14f;
		profileModel.Typography.LineSpacingRule = LineSpacingType.Multiple;
		profileModel.Typography.LineSpacingValue = 1.15f;
		profileModel.Typography.SpacingBeforePt = 0f;
		profileModel.Typography.SpacingAfterPt = 6f;
		profileModel.Typography.FirstLineIndentMm = 12.7f;
		profileModel.Typography.Alignment = AlignmentType.Justify;
		profileModel.PageSetup.PaperSizeName = "A4";
		profileModel.PageSetup.Orientation = PaperOrientation.Portrait;
		profileModel.PageSetup.MarginTopMm = 20f;
		profileModel.PageSetup.MarginBottomMm = 20f;
		profileModel.PageSetup.MarginLeftMm = 30f;
		profileModel.PageSetup.MarginRightMm = 20f;
		profileModel.PageSetup.DuplexMarginLeftMm = 20f;
		profileModel.PageSetup.DuplexMarginRightMm = 30f;
		profileModel.PageSetup.MirrorMargins = false;
		profileModel.PageNumber.SizePt = 14f;
		profileModel.PageNumber.Position = PageNumberPosition.TopCenter;
		profileModel.PageNumber.DistanceFromTopMm = 10f;
		profileModel.PageNumber.HideFirstPage = true;
		profileModel.TablePreset.HeaderFontSizePt = 14f;
		profileModel.TablePreset.HeaderBold = true;
		profileModel.TablePreset.BodyFontSizePt = 14f;
		profileModel.TablePreset.CellPaddingTopMm = 1.5f;
		profileModel.TablePreset.CellPaddingBottomMm = 1.5f;
		profileModel.TablePreset.CellPaddingLeftMm = 2f;
		profileModel.TablePreset.CellPaddingRightMm = 2f;
		Dictionary<SemanticRole, ComponentStyle> componentStyles = profileModel.ComponentStyles;
		componentStyles[SemanticRole.PartyHeader] = new ComponentStyle(SemanticRole.PartyHeader, 15f, bold: true, italic: false, AlignmentType.Center, UnderlineRuleType.SolidFullWidth)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SuperiorAgency] = new ComponentStyle(SemanticRole.SuperiorAgency, 14f, bold: false, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.IssuingAgency] = new ComponentStyle(SemanticRole.IssuingAgency, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.DocumentNumber] = new ComponentStyle(SemanticRole.DocumentNumber, 14f, bold: false, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.DocumentDate] = new ComponentStyle(SemanticRole.DocumentDate, 13f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.DocumentType] = new ComponentStyle(SemanticRole.DocumentType, 16f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true,
			FirstLineIndentMm = 0f
		};
		componentStyles[SemanticRole.SubjectExcerpt] = new ComponentStyle(SemanticRole.SubjectExcerpt, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			FirstLineIndentMm = 0f
		};
		componentStyles[SemanticRole.DispatchSubject] = new ComponentStyle(SemanticRole.DispatchSubject, 12f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.LegalGround] = new ComponentStyle(SemanticRole.LegalGround, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 12.7f,
			SpacingAfterPt = 6f
		};
		componentStyles[SemanticRole.BodyText] = new ComponentStyle(SemanticRole.BodyText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 12.7f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Multiple,
			LineSpacingValue = 1.15f
		};
		componentStyles[SemanticRole.PartHeader] = new ComponentStyle(SemanticRole.PartHeader, 15f, bold: true, italic: false, AlignmentType.Center)
		{
			FirstLineIndentMm = 0f
		};
		componentStyles[SemanticRole.ChapterHeader] = new ComponentStyle(SemanticRole.ChapterHeader, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			FirstLineIndentMm = 0f
		};
		componentStyles[SemanticRole.SectionHeader] = new ComponentStyle(SemanticRole.SectionHeader, 14f, bold: true, italic: false, AlignmentType.Justify)
		{
			IsUppercase = true,
			FirstLineIndentMm = 12.7f,
			SpacingBeforePt = 6f,
			SpacingAfterPt = 6f
		};
		componentStyles[SemanticRole.ArticleTitle] = new ComponentStyle(SemanticRole.ArticleTitle, 14f, bold: true, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 12.7f
		};
		componentStyles[SemanticRole.ClauseText] = new ComponentStyle(SemanticRole.ClauseText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 12.7f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Multiple,
			LineSpacingValue = 1.15f
		};
		componentStyles[SemanticRole.PointText] = new ComponentStyle(SemanticRole.PointText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 12.7f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Multiple,
			LineSpacingValue = 1.15f
		};
		componentStyles[SemanticRole.SignAuthority] = new ComponentStyle(SemanticRole.SignAuthority, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SignTitle] = new ComponentStyle(SemanticRole.SignTitle, 14f, bold: false, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SignerName] = new ComponentStyle(SemanticRole.SignerName, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsTitleCase = true
		};
		componentStyles[SemanticRole.DirectRecipient] = new ComponentStyle(SemanticRole.DirectRecipient, 14f, bold: true, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.RecipientHeader] = new ComponentStyle(SemanticRole.RecipientHeader, 13f, bold: true);
		componentStyles[SemanticRole.RecipientItem] = new ComponentStyle(SemanticRole.RecipientItem, 11f);
		componentStyles[SemanticRole.SecurityIndicator] = new ComponentStyle(SemanticRole.SecurityIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true,
			HasBorderBox = true
		};
		componentStyles[SemanticRole.UrgencyIndicator] = new ComponentStyle(SemanticRole.UrgencyIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.DraftIndicator] = new ComponentStyle(SemanticRole.DraftIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true,
			HasBorderBox = true
		};
		componentStyles[SemanticRole.CirculationScope] = new ComponentStyle(SemanticRole.CirculationScope, 12f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.CreatorSymbol] = new ComponentStyle(SemanticRole.CreatorSymbol, 8f, bold: false, italic: true);
		componentStyles[SemanticRole.ContactInfo] = new ComponentStyle(SemanticRole.ContactInfo, 11f);
		componentStyles[SemanticRole.AppendixTitle] = new ComponentStyle(SemanticRole.AppendixTitle, 14f, bold: true, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.CopyPaper] = new ComponentStyle(SemanticRole.CopyPaper, 14f, bold: false, italic: true, AlignmentType.Right);
		ComponentLayoutDefaults.Apply(profileModel);
		return profileModel;
	}

	public static string GetDocTypeAbbreviation(string documentType)
	{
		if (string.IsNullOrWhiteSpace(documentType))
		{
			return string.Empty;
		}
		string text = documentType.Trim();
		if (DocTypeAbbreviations.TryGetValue(text, out var value))
		{
			return value;
		}
		foreach (KeyValuePair<string, string> docTypeAbbreviation in DocTypeAbbreviations)
		{
			if (string.Equals(docTypeAbbreviation.Key, text, StringComparison.CurrentCultureIgnoreCase))
			{
				return docTypeAbbreviation.Value;
			}
		}
		return string.Empty;
	}
}
