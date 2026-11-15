using System.Collections.Generic;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.Core.Rules;

public static class HD05Rules
{
	public const string ProfileId = "PROFILE_DANG";

	public const string Name = "Văn bản của Đảng";

	public const string Source = "Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng";

	public const string Description = "Chuẩn thể thức và kỹ thuật trình bày văn bản của Đảng Cộng sản Việt Nam";

	public static ProfileModel CreateProfile()
	{
		ProfileModel profileModel = new ProfileModel("PROFILE_DANG", RulePackType.HD05_VPTW, "Văn bản của Đảng", "Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng", "Chuẩn thể thức và kỹ thuật trình bày văn bản của Đảng Cộng sản Việt Nam");
		profileModel.Typography.FontName = "Times New Roman";
		profileModel.Typography.BodySizePt = 14f;
		profileModel.Typography.LineSpacingRule = LineSpacingType.Exactly;
		profileModel.Typography.LineSpacingValue = 18f;
		profileModel.Typography.SpacingBeforePt = 6f;
		profileModel.Typography.SpacingAfterPt = 6f;
		profileModel.Typography.NoSpaceBetweenParagraphsOfSameStyle = false;
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
		componentStyles[SemanticRole.DocumentDate] = new ComponentStyle(SemanticRole.DocumentDate, 14f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.DocumentType] = new ComponentStyle(SemanticRole.DocumentType, 15f, bold: true, italic: false, AlignmentType.Center)
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
			Prefix = "- ",
			FirstLineIndentMm = 10f,
			SpacingAfterPt = 6f
		};
		componentStyles[SemanticRole.BodyText] = new ComponentStyle(SemanticRole.BodyText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 10f,
			SpacingBeforePt = 6f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Exactly,
			LineSpacingValue = 18f
		};
		componentStyles[SemanticRole.PartHeader] = new ComponentStyle(SemanticRole.PartHeader, 14f, bold: true, italic: false, AlignmentType.Center);
		componentStyles[SemanticRole.ChapterHeader] = new ComponentStyle(SemanticRole.ChapterHeader, 14f, bold: true, italic: false, AlignmentType.Center);
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
			SpacingBeforePt = 6f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Exactly,
			LineSpacingValue = 18f
		};
		componentStyles[SemanticRole.PointText] = new ComponentStyle(SemanticRole.PointText, 14f, bold: false, italic: false, AlignmentType.Justify)
		{
			FirstLineIndentMm = 10f,
			SpacingBeforePt = 6f,
			SpacingAfterPt = 6f,
			LineSpacingRule = LineSpacingType.Exactly,
			LineSpacingValue = 18f
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
		componentStyles[SemanticRole.DirectRecipient] = new ComponentStyle(SemanticRole.DirectRecipient, 14f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.RecipientHeader] = new ComponentStyle(SemanticRole.RecipientHeader, 14f, bold: false, italic: false, AlignmentType.Left, UnderlineRuleType.SolidFullWidth);
		componentStyles[SemanticRole.RecipientItem] = new ComponentStyle(SemanticRole.RecipientItem, 12f);
		componentStyles[SemanticRole.UrgencyIndicator] = new ComponentStyle(SemanticRole.UrgencyIndicator, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.SecurityIndicator] = new ComponentStyle(SemanticRole.SecurityIndicator, 14f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.CirculationScope] = new ComponentStyle(SemanticRole.CirculationScope, 12f, bold: true, italic: false, AlignmentType.Center)
		{
			IsUppercase = true
		};
		componentStyles[SemanticRole.ContactInfo] = new ComponentStyle(SemanticRole.ContactInfo, 10f);
		componentStyles[SemanticRole.DraftIndicator] = new ComponentStyle(SemanticRole.DraftIndicator, 14f, bold: false, italic: true, AlignmentType.Center);
		componentStyles[SemanticRole.CreatorSymbol] = new ComponentStyle(SemanticRole.CreatorSymbol, 11f);
		ComponentLayoutDefaults.Apply(profileModel);
		return profileModel;
	}
}
