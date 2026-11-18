using System;
using System.Collections.Generic;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.Core.Rules;

public static class TT68Rules
{
	public const string ProfileId = "PROFILE_TT68_BQP";

	public const string Name = "Văn bản Quân đội";

	public const string Source = "Thông tư số 68/2023/TT-BQP ngày 25/9/2023 của Bộ trưởng Bộ Quốc phòng";

	public const string Description = "Quy định công tác văn thư trong Bộ Quốc phòng";

	public const string EndOfDocumentMark = "./.";

	public const string UrgencyBoxSizeThuongKhan = "40 mm x 8 mm";

	public const string UrgencyBoxSizeHoaToc = "30 mm x 8 mm";

	public const string UrgencyBoxSizeKhan = "20 mm x 8 mm";

	public const string MandatorySigningInkColor = "Xanh";

	public const float DigitalStampDigitizedSizeMm = 23f;

	public const string DigitalStampFormat = ".png";

	public const string IncomingStampRelativePosition = "UnderDocumentDate";

	public const string DigitalCopyPdfVersionMin = "1.4";

	public const int DigitalCopyMinResolutionDpi = 200;

	public const string DigitalCopyColorMode = "Color";

	public const int DigitalCopyScalePercentage = 100;

	public const float DigitalCopyMetadataFontSizePt = 10f;

	public const string MilitaryOfficeNoSlashPrefix = "Văn phòng ";

	public const int ArchivingAdvanceNoticeDays = 30;

	public const int CentralConferenceWordThreshold = 11;

	public const string RecipientSuperiorFirstLine = "- Như trên;";

	public const string SubClauseTietPrefix = "- ";

	public const string SubClauseTieuTietPrefix = "+ ";

	public const int MaxOverlappingSheets = 5;

	public const int MaxRetainDossierYears = 2;

	public const int ArchiveTransferDeadlineConstructionMonths = 3;

	public const int ArchiveTransferDeadlineGeneralYears = 1;

	public const int MinRetentionPeriodForArchivingYears = 5;

	public const int AnnualReportCutoffStartMonth = 12;

	public const int AnnualReportCutoffStartDay = 15;

	public const int AnnualReportCutoffEndMonth = 12;

	public const int AnnualReportCutoffEndDay = 14;

	public const string CircularSignAuthority = "KT. BỘ TRƯỞNG";

	public const string CircularSignTitle = "THỨ TRƯỞNG";

	public const string CircularSignerRankAndName = "Thượng tướng Lê Huy Vịnh";

	public const string SuperiorAgencyGeneralStaff = "QUÂN ĐỘI NHÂN DÂN VIỆT NAM";

	public const float SubjectExcerptUnderlineRatio = 1f / 3f;

	public const float IssuingAgencyUnderlineRatioMin = 1f / 3f;

	public const float IssuingAgencyUnderlineRatioMax = 0.5f;

	public const string CityAbbreviationPrefix = "TP.";

	public const float DigitizedStampMetadataFontSizePt = 10f;

	public const bool AllowAcademicTitleInAdministrativeDocs = false;

	public static readonly string[] UrgencyLevels = new string[4] { "HỎA TỐC HẸN GIỜ", "HỎA TỐC", "THƯỢNG KHẨN", "KHẨN" };

	public static readonly string[] IntendedSecurityLevels = new string[4] { "TuM", "TUM", "TM", "M" };

	public static readonly string[] MilitaryRanks = new string[12]
	{
		"Đại tướng", "Thượng tướng", "Trung tướng", "Thiếu tướng", "Đại tá", "Thượng tá", "Trung tá", "Thiếu tá", "Đại úy", "Thượng úy",
		"Trung úy", "Thiếu úy"
	};

	public static readonly string[] AdministrativeDocumentTypes = new string[29]
	{
		"Nghị quyết (cá biệt)", "Quyết định (cá biệt)", "Chỉ thị", "Quy chế", "Quy định", "Thông cáo", "Thông báo", "Hướng dẫn", "Chương trình", "Kế hoạch",
		"Phương án", "Đề án", "Dự án", "Báo cáo", "Biên bản", "Tờ trình", "Hợp đồng", "Công văn", "Công điện", "Bản ghi nhớ",
		"Bản thỏa thuận", "Giấy ủy quyền", "Giấy mời", "Giấy giới thiệu", "Giấy nghỉ phép", "Phiếu gửi", "Phiếu chuyển", "Phiếu báo", "Thư công"
	};

	public static readonly string[] CopyTypes = new string[3] { "Sao y", "Sao lục", "Trích sao" };

	public static readonly string[] CopyPaperTypes = new string[3] { "SAO Y", "SAO LỤC", "TRÍCH SAO" };

	public static readonly string[] SingleSyllableGeographicNames = new string[11]
	{
		"Sông Hồng", "Sông Đáy", "Sông Lô", "Hồ Tây", "Hồ Gươm", "Núi Cốc", "Cửa Lò", "Bến Nghé", "Buôn Hồ", "Buôn Đôn",
		"Bản Keo"
	};

	public static readonly string[] SpecialAdminUnits = new string[1] { "Thành phố Hồ Chí Minh" };

	public static readonly string[] HistoricalGeographicNames = new string[5] { "Thủ đô Hà Nội", "Kinh thành Thăng Long", "Kinh đô Huế", "Cố đô Hoa Lư", "Tổ quốc Việt Nam" };

	public static readonly string[] SpecialMilitaryOrganizations = new string[6] { "Quân ủy Trung ương", "Văn phòng Trung ương Đảng", "Ban Chấp hành Trung ương Đảng Cộng sản Việt Nam", "Bộ Quốc phòng", "Bộ Tổng Tham mưu", "Tổng cục Chính trị" };

	public static readonly string[] MilitaryLeadershipPrefixes = new string[5] { "Thủ trưởng Bộ Tư lệnh", "Thủ trưởng Ban Giám đốc", "Thủ trưởng Ban Giám hiệu", "Thủ trưởng Bộ Chỉ huy", "Lãnh đạo" };

	public static readonly string[] StateLeadershipTitles = new string[13]
	{
		"Tổng Bí thư", "Ủy viên Bộ Chính trị", "Thường trực Ban Bí thư", "Bí thư Trung ương Đảng", "Ủy viên Ban Chấp hành Trung ương Đảng", "Chủ tịch Quốc hội", "Phó Chủ tịch Quốc hội", "Chủ tịch nước", "Phó Chủ tịch nước", "Thủ tướng Chính phủ",
		"Phó Thủ tướng Chính phủ", "Chánh án Toà án nhân dân tối cao", "Viện trưởng Viện Kiểm sát nhân dân tối cao"
	};

	public static readonly string[] StateHonors = new string[9] { "Anh hùng Lao động", "Anh hùng Lực lượng vũ trang nhân dân", "Bà mẹ Việt Nam Anh hùng", "Nhà giáo Ưu tú", "Nhà giáo Nhân dân", "Nghệ sĩ Ưu tú", "Nghệ sĩ Nhân dân", "Thầy thuốc Ưu tú", "Thầy thuốc Nhân dân" };

	public static readonly string[] MedalTypes = new string[11]
	{
		"Huân chương Sao vàng", "Huân chương Lao động", "Huân chương Kháng chiến", "Huân chương Hữu nghị", "Huân chương Hồ Chí Minh", "Huân chương Quân công", "Huân chương Chiến công", "Huân chương Bảo vệ Tổ quốc", "Huy chương Quân kỳ quyết thắng", "Huy chương Chiến sĩ vẻ vang",
		"Huy chương Chiến sĩ nở hoa thắng trận"
	};

	public static readonly string[] CitationUppercaseElements = new string[5] { "Điều", "Tiểu mục", "Mục", "Chương", "Phần" };

	public static readonly string[] CitationLowercaseElements = new string[2] { "điểm", "khoản" };

	public static readonly string[] SpecialPromulgationAuthorities = new string[3] { "Thủ trưởng Bộ Quốc phòng", "Thủ trưởng Bộ Tổng Tham mưu", "Thủ trưởng Tổng cục Chính trị" };

	public static readonly string[] EdocOutgoingMetadataFields = new string[16]
	{
		"FileCode", "DocOrdinal", "TypeName", "CodeNumber", "CodeNotation", "IssuedDate", "OrganName", "Subject", "Language", "PageAmount",
		"Description", "SignerInfo", "To", "Priority", "IssuedAmount", "DueDate"
	};

	public static readonly string[] PressTypeNames = new string[4] { "báo", "tạp chí", "tập san", "bản tin" };

	public static readonly Dictionary<string, string> DocTypeAbbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
	{
		{ "Nghị quyết", "NQ" },
		{ "Quyết định", "QĐ" },
		{ "Chỉ thị", "CT" },
		{ "Quy chế", "QC" },
		{ "Quy định", "QyĐ" },
		{ "Thông cáo", "TC" },
		{ "Thông báo", "TB" },
		{ "Hướng dẫn", "HD" },
		{ "Chương trình", "CTr" },
		{ "Kế hoạch", "KH" },
		{ "Phương án", "PA" },
		{ "Đề án", "ĐA" },
		{ "Dự án", "DA" },
		{ "Báo cáo", "BC" },
		{ "Biên bản", "BB" },
		{ "Tờ trình", "TTr" },
		{ "Hợp đồng", "HĐ" },
		{ "Công điện", "CĐ" },
		{ "Bản ghi nhớ", "BGN" },
		{ "Bản thỏa thuận", "BTT" },
		{ "Giấy ủy quyền", "GUQ" },
		{ "Giấy công tác", "GCT" },
		{ "Giấy mời", "GM" },
		{ "Giấy giới thiệu", "GGT" },
		{ "Giấy nghỉ phép", "GNP" },
		{ "Phiếu gửi", "PG" },
		{ "Phiếu chuyển", "PC" },
		{ "Phiếu báo", "PB" }
	};

	public static readonly Dictionary<string, string> CopyTypeAbbreviations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
	{
		{ "Bản sao y", "SY" },
		{ "Bản sao lục", "SL" },
		{ "Bản trích sao", "TrS" }
	};

	public static readonly string[] DraftIndicators = new string[2] { "DỰ THẢO", "DỰ THẢO LẦN" };

	public static readonly string[] CirculationScopes = new string[2] { "XEM XONG TRẢ LẠI", "LƯU HÀNH NỘI BỘ" };

	public static ProfileModel CreateProfile()
	{
		ProfileModel profileModel = new ProfileModel("PROFILE_TT68_BQP", RulePackType.TT68_2023_BQP, "Văn bản Quân đội", "Thông tư số 68/2023/TT-BQP ngày 25/9/2023 của Bộ trưởng Bộ Quốc phòng", "Quy định công tác văn thư trong Bộ Quốc phòng");
		profileModel.Typography.FontName = "Times New Roman";
		profileModel.Typography.BodySizePt = 14f;
		profileModel.Typography.LineSpacingRule = LineSpacingType.Multiple;
		profileModel.Typography.LineSpacingValue = 1.15f;
		profileModel.Typography.SpacingBeforePt = 0f;
		profileModel.Typography.SpacingAfterPt = 6f;
		profileModel.Typography.FirstLineIndentMm = 10f;
		profileModel.Typography.Alignment = AlignmentType.Justify;
		profileModel.PageSetup.PaperSizeName = "A4";
		profileModel.PageSetup.Orientation = PaperOrientation.Portrait;
		profileModel.PageSetup.MarginTopMm = 20f;
		profileModel.PageSetup.MarginBottomMm = 20f;
		profileModel.PageSetup.MarginLeftMm = 30f;
		profileModel.PageSetup.MarginRightMm = 15f;
		profileModel.PageSetup.DuplexMarginLeftMm = 15f;
		profileModel.PageSetup.DuplexMarginRightMm = 30f;
		profileModel.PageSetup.MirrorMargins = false;
		profileModel.PageNumber.SizePt = 14f;
		profileModel.PageNumber.Position = PageNumberPosition.TopCenter;
		profileModel.PageNumber.DistanceFromTopMm = 10f;
		profileModel.PageNumber.CenterVerticallyInTopMargin = true;
		profileModel.PageNumber.HideFirstPage = true;
		profileModel.TablePreset.HeaderFontSizePt = 14f;
		profileModel.TablePreset.HeaderBold = true;
		profileModel.TablePreset.BodyFontSizePt = 14f;
		profileModel.TablePreset.CellPaddingTopMm = 1.5f;
		profileModel.TablePreset.CellPaddingBottomMm = 1.5f;
		profileModel.TablePreset.CellPaddingLeftMm = 2f;
		profileModel.TablePreset.CellPaddingRightMm = 2f;
		Dictionary<SemanticRole, ComponentStyle> componentStyles = profileModel.ComponentStyles;
		componentStyles[SemanticRole.NationalHeader] = new ComponentStyle(SemanticRole.NationalHeader, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.Motto] = new ComponentStyle(SemanticRole.Motto, 13f, bold: true, italic: false, AlignmentType.Center, UnderlineRuleType.SolidFullWidth);
		componentStyles[SemanticRole.SuperiorAgency] = new ComponentStyle(SemanticRole.SuperiorAgency, 13f, bold: false, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.IssuingAgency] = new ComponentStyle(SemanticRole.IssuingAgency, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.DocumentNumber] = new ComponentStyle(SemanticRole.DocumentNumber, 13f, bold: false, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.DocumentDate] = new ComponentStyle(SemanticRole.DocumentDate, 13f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.DocumentType] = new ComponentStyle(SemanticRole.DocumentType, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true,
			FirstLineIndentMm = 0f
		};
		componentStyles[SemanticRole.SubjectExcerpt] = new ComponentStyle(SemanticRole.SubjectExcerpt, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			FirstLineIndentMm = 0f
		};
		componentStyles[SemanticRole.DispatchSubject] = new ComponentStyle(SemanticRole.DispatchSubject, 12f, bold: false, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.LegalGround] = new ComponentStyle(SemanticRole.LegalGround, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			SpacingAfterPt = 6f,
			FirstLineIndentMm = 10f
		};
		componentStyles[SemanticRole.BodyText] = new ComponentStyle(SemanticRole.BodyText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 10f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Multiple,
			LineSpacingValue = 1.15f
		};
		componentStyles[SemanticRole.PartHeader] = new ComponentStyle(SemanticRole.PartHeader, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.ChapterHeader] = new ComponentStyle(SemanticRole.ChapterHeader, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SectionHeader] = new ComponentStyle(SemanticRole.SectionHeader, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			FirstLineIndentMm = 0f,
			SpacingBeforePt = 6f,
			SpacingAfterPt = 6f
		};
		componentStyles[SemanticRole.ArticleTitle] = new ComponentStyle(SemanticRole.ArticleTitle, 14f, bold: true, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 10f
		};
		componentStyles[SemanticRole.ClauseText] = new ComponentStyle(SemanticRole.ClauseText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 10f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Multiple,
			LineSpacingValue = 1.15f
		};
		componentStyles[SemanticRole.PointText] = new ComponentStyle(SemanticRole.PointText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 10f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Multiple,
			LineSpacingValue = 1.15f
		};
		componentStyles[SemanticRole.SignAuthority] = new ComponentStyle(SemanticRole.SignAuthority, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SignTitle] = new ComponentStyle(SemanticRole.SignTitle, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SignerName] = new ComponentStyle(SemanticRole.SignerName, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsTitleCase = true
		};
		componentStyles[SemanticRole.SignRankName] = new ComponentStyle(SemanticRole.SignRankName, 14f, bold: true, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.DirectRecipient] = new ComponentStyle(SemanticRole.DirectRecipient, 14f, bold: false, italic: false, AlignmentType.Left);
		componentStyles[SemanticRole.RecipientHeader] = new ComponentStyle(SemanticRole.RecipientHeader, 12f, bold: true, italic: true);
		componentStyles[SemanticRole.RecipientItem] = new ComponentStyle(SemanticRole.RecipientItem, 11f);
		componentStyles[SemanticRole.UrgencyIndicator] = new ComponentStyle(SemanticRole.UrgencyIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			HasBorderBox = true,
			IsUppercase = true
		};
		componentStyles[SemanticRole.SecurityIndicator] = new ComponentStyle(SemanticRole.SecurityIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.CirculationScope] = new ComponentStyle(SemanticRole.CirculationScope, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			HasBorderBox = true,
			IsUppercase = true
		};
		componentStyles[SemanticRole.DraftIndicator] = new ComponentStyle(SemanticRole.DraftIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			HasBorderBox = true,
			IsUppercase = true
		};
		componentStyles[SemanticRole.CreatorSymbol] = new ComponentStyle(SemanticRole.CreatorSymbol, 11f);
		componentStyles[SemanticRole.ContactInfo] = new ComponentStyle(SemanticRole.ContactInfo, 12f)
		{
			UnderlineRule = UnderlineRuleType.SolidFullCell
		};
		componentStyles[SemanticRole.AppendixTitle] = new ComponentStyle(SemanticRole.AppendixTitle, 14f, bold: true, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.CopyPaper] = new ComponentStyle(SemanticRole.CopyPaper, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.CopyElectronic] = new ComponentStyle(SemanticRole.CopyElectronic, 10f, bold: false, italic: false, AlignmentType.Right);
		ComponentLayoutDefaults.Apply(profileModel);
		return profileModel;
	}
}
