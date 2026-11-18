using System;
using System.Collections.Generic;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.Core.Rules;

public static class ComponentLayoutDefaults
{
	public static void Apply(ProfileModel profile)
	{
		foreach (KeyValuePair<SemanticRole, ComponentStyle> componentStyle in profile.ComponentStyles)
		{
			SemanticRole key = componentStyle.Key;
			if (key == SemanticRole.BodyText || key == SemanticRole.ClauseText || key == SemanticRole.PointText || key == SemanticRole.LegalGround)
			{
				ComponentStyle value = componentStyle.Value;
				value.SpacingBeforePt = profile.Typography.SpacingBeforePt;
				value.SpacingAfterPt = ((profile.RulePack == RulePackType.HD29_TWDTN) ? Math.Min(6f, profile.Typography.SpacingAfterPt) : Math.Max(6f, profile.Typography.SpacingAfterPt));
				value.LineSpacingRule = profile.Typography.LineSpacingRule;
				value.LineSpacingValue = profile.Typography.LineSpacingValue;
				value.NoSpaceBetweenParagraphsOfSameStyle = profile.Typography.NoSpaceBetweenParagraphsOfSameStyle;
				if (!value.FirstLineIndentMm.HasValue)
				{
					value.FirstLineIndentMm = profile.Typography.FirstLineIndentMm;
				}
			}
			bool flag = key == SemanticRole.DocumentType || key == SemanticRole.SubjectExcerpt;
			if (flag || key == SemanticRole.NationalHeader || key == SemanticRole.Motto || key == SemanticRole.PartyHeader || key == SemanticRole.SuperiorAgency || key == SemanticRole.IssuingAgency || key == SemanticRole.DocumentNumber || key == SemanticRole.DocumentDate || key == SemanticRole.DispatchSubject || key == SemanticRole.RecipientHeader || key == SemanticRole.RecipientItem || key == SemanticRole.SignAuthority || key == SemanticRole.SignTitle || key == SemanticRole.SignerName || key == SemanticRole.SignRankName)
			{
				ComponentStyle value2 = componentStyle.Value;
				if (!value2.LineSpacingRule.HasValue)
				{
					value2.LineSpacingRule = LineSpacingType.Single;
				}
				if (!value2.LineSpacingValue.HasValue)
				{
					value2.LineSpacingValue = 1f;
				}
				if (!value2.SpacingBeforePt.HasValue)
				{
					value2.SpacingBeforePt = ((key == SemanticRole.DocumentType) ? 6 : 0);
				}
				if (!value2.SpacingAfterPt.HasValue)
				{
					value2.SpacingAfterPt = (flag ? 2 : 0);
				}
				if (!value2.FirstLineIndentMm.HasValue)
				{
					value2.FirstLineIndentMm = 0f;
				}
				if (!value2.KeepLinesTogether.HasValue)
				{
					value2.KeepLinesTogether = true;
				}
				if (!value2.KeepWithNext.HasValue)
				{
					value2.KeepWithNext = flag || key == SemanticRole.SignAuthority || key == SemanticRole.SignTitle;
				}
			}
		}
	}
}
