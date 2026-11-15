using System.Collections.Generic;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.Core.Rules;

public static class ND30Rules
{
	public const string ProfileId = "PROFILE_HANH_CHINH";

	public const string Name = "Văn bản Hành chính";

	public const string Source = "Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ";

	public const string Description = "Chuẩn thể thức và kỹ thuật trình bày văn bản hành chính Nhà nước và Doanh nghiệp";

	public static ProfileModel CreateProfile()
	{
		ProfileModel profileModel = new ProfileModel("PROFILE_HANH_CHINH", RulePackType.ND30_2020, "Văn bản Hành chính", "Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ", "Chuẩn thể thức và kỹ thuật trình bày văn bản hành chính Nhà nước và Doanh nghiệp");
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
		profileModel.PageSetup.MirrorMargins = false;
		profileModel.PageNumber.SizePt = 14f;
		profileModel.PageNumber.Position = PageNumberPosition.TopCenter;
		profileModel.PageNumber.DistanceFromTopMm = 10f;
		profileModel.PageNumber.HideFirstPage = true;
		profileModel.TablePreset.HeaderFontSizePt = 13f;
		profileModel.TablePreset.HeaderBold = true;
		profileModel.TablePreset.BodyFontSizePt = 13f;
		profileModel.TablePreset.CellPaddingTopMm = 1.5f;
		profileModel.TablePreset.CellPaddingBottomMm = 1.5f;
		profileModel.TablePreset.CellPaddingLeftMm = 2f;
		profileModel.TablePreset.CellPaddingRightMm = 2f;
		Dictionary<SemanticRole, ComponentStyle> componentStyles = profileModel.ComponentStyles;
		componentStyles[SemanticRole.NationalHeader] = new ComponentStyle(SemanticRole.NationalHeader, 12f, bold: true, italic: false, AlignmentType.Center)
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
		componentStyles[SemanticRole.DocumentDate] = new ComponentStyle(SemanticRole.DocumentDate, 14f, bold: false, italic: true, AlignmentType.Center);
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
		componentStyles[SemanticRole.LegalGround] = new ComponentStyle(SemanticRole.LegalGround, 14f, bold: false, italic: true, AlignmentType.Justify)
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
		componentStyles[SemanticRole.DirectRecipient] = new ComponentStyle(SemanticRole.DirectRecipient, 14f);
		componentStyles[SemanticRole.RecipientHeader] = new ComponentStyle(SemanticRole.RecipientHeader, 12f, bold: true, italic: true);
		componentStyles[SemanticRole.RecipientItem] = new ComponentStyle(SemanticRole.RecipientItem, 11f);
		componentStyles[SemanticRole.UrgencyIndicator] = new ComponentStyle(SemanticRole.UrgencyIndicator, 13f, bold: true, italic: false, AlignmentType.Center)
		{
			HasBorderBox = true,
			IsUppercase = true
		};
		componentStyles[SemanticRole.SecurityIndicator] = new ComponentStyle(SemanticRole.SecurityIndicator, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.CirculationScope] = new ComponentStyle(SemanticRole.CirculationScope, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			HasBorderBox = true,
			IsUppercase = true
		};
		componentStyles[SemanticRole.ContactInfo] = new ComponentStyle(SemanticRole.ContactInfo, 11f)
		{
			UnderlineRule = UnderlineRuleType.SolidFullCell
		};
		componentStyles[SemanticRole.CreatorSymbol] = new ComponentStyle(SemanticRole.CreatorSymbol, 11f, bold: false, italic: false, AlignmentType.Right);
		ComponentLayoutDefaults.Apply(profileModel);
		return profileModel;
	}
}
