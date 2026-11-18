using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Formatters;

public static class ParagraphFormatter
{
	public static void FormatParagraph(Paragraph p, ComponentStyle style, TypographyPreset defaultTypo)
	{
		if (p == null)
		{
			return;
		}
		Range comObj = null;
		Font comObj2 = null;
		ParagraphFormat comObj3 = null;
		try
		{
			comObj = p.Range;
			if (comObj == null)
			{
				return;
			}
			comObj2 = comObj.Font;
			comObj3 = p.Format;
			string name = ((defaultTypo != null) ? defaultTypo.FontName : "Times New Roman");
			float size = style?.SizePt ?? defaultTypo?.BodySizePt ?? 14f;
			comObj2.Name = name;
			comObj2.Size = size;
			comObj2.ColorIndex = WdColorIndex.wdBlack;
			if (style != null && style.Bold)
			{
				comObj2.Bold = -1;
			}
			else if (style != null && (style.Role == SemanticRole.NationalHeader || style.Role == SemanticRole.Motto || style.Role == SemanticRole.DocumentDate || style.Role == SemanticRole.DocumentNumber || style.Role == SemanticRole.DispatchSubject || style.Role == SemanticRole.DirectRecipient || style.Role == SemanticRole.LegalGround || style.Role == SemanticRole.RecipientHeader || style.Role == SemanticRole.RecipientItem || style.Role == SemanticRole.SignTitle))
			{
				comObj2.Bold = 0;
			}
			if (style != null && style.Italic)
			{
				comObj2.Italic = -1;
			}
			else if (style != null && (style.Role == SemanticRole.NationalHeader || style.Role == SemanticRole.DocumentNumber || style.Role == SemanticRole.DocumentType || style.Role == SemanticRole.SubjectExcerpt || style.Role == SemanticRole.DispatchSubject || style.Role == SemanticRole.DirectRecipient || style.Role == SemanticRole.LegalGround || style.Role == SemanticRole.ArticleTitle || style.Role == SemanticRole.SectionHeader || style.Role == SemanticRole.PartHeader || style.Role == SemanticRole.ChapterHeader || style.Role == SemanticRole.BodyText || style.Role == SemanticRole.ClauseText || style.Role == SemanticRole.PointText || style.Role == SemanticRole.RecipientHeader || style.Role == SemanticRole.RecipientItem || style.Role == SemanticRole.SignTitle || style.Role == SemanticRole.SignAuthority || style.Role == SemanticRole.SignerName))
			{
				comObj2.Italic = 0;
			}
			if (style == null || style.UnderlineRule == UnderlineRuleType.None)
			{
				comObj2.Underline = WdUnderline.wdUnderlineNone;
				comObj.Underline = WdUnderline.wdUnderlineNone;
			}
			else if (style.UnderlineRule == UnderlineRuleType.SolidFullWidth && style.Role == SemanticRole.RecipientHeader)
			{
				comObj2.Underline = WdUnderline.wdUnderlineSingle;
			}
			try
			{
				comObj3.Borders.Enable = 0;
				foreach (Border border in comObj3.Borders)
				{
					try
					{
						border.LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
			}
			catch
			{
			}
			AlignmentType alignmentType = style?.Alignment ?? defaultTypo?.Alignment ?? AlignmentType.Justify;
			comObj3.Alignment = ConvertAlignment(alignmentType);
			comObj3.LeftIndent = 0f;
			comObj3.RightIndent = 0f;
			LineSpacingType rule = ((style != null && style.LineSpacingRule.HasValue) ? style.LineSpacingRule.Value : (defaultTypo?.LineSpacingRule ?? LineSpacingType.Multiple));
			float value = ((style != null && style.LineSpacingValue.HasValue) ? style.LineSpacingValue.Value : (defaultTypo?.LineSpacingValue ?? 1.15f));
			ApplyLineSpacing(comObj3, rule, value);
			float spaceBefore = ((style != null && style.SpacingBeforePt.HasValue) ? style.SpacingBeforePt.Value : (defaultTypo?.SpacingBeforePt ?? 0f));
			float spaceAfter = ((style != null && style.SpacingAfterPt.HasValue) ? style.SpacingAfterPt.Value : (defaultTypo?.SpacingAfterPt ?? 6f));
			comObj3.SpaceBefore = spaceBefore;
			comObj3.SpaceAfter = spaceAfter;
			bool noSpace = (style != null && style.NoSpaceBetweenParagraphsOfSameStyle.HasValue)
				? style.NoSpaceBetweenParagraphsOfSameStyle.Value
				: (defaultTypo != null && defaultTypo.NoSpaceBetweenParagraphsOfSameStyle);
			try
			{
				Style pStyle = comObj3.get_Style() as Style;
				if (pStyle != null && pStyle.NoSpaceBetweenParagraphsOfSameStyle != noSpace)
				{
					pStyle.NoSpaceBetweenParagraphsOfSameStyle = noSpace;
				}
			}
			catch
			{
			}
			if (style != null && style.KeepWithNext.HasValue)
			{
				comObj3.KeepWithNext = (style.KeepWithNext.Value ? (-1) : 0);
			}
			if (style != null && style.KeepLinesTogether.HasValue)
			{
				comObj3.KeepTogether = (style.KeepLinesTogether.Value ? (-1) : 0);
			}
			try
			{
				comObj3.WidowControl = -1;
			}
			catch
			{
			}
			if (style != null && style.FirstLineIndentMm.HasValue)
			{
				comObj3.FirstLineIndent = style.FirstLineIndentPt.Value;
			}
			else if (style != null && (alignmentType == AlignmentType.Center || alignmentType == AlignmentType.Right))
			{
				comObj3.FirstLineIndent = 0f;
			}
			else if (style != null && (style.Role == SemanticRole.PartHeader || style.Role == SemanticRole.ChapterHeader || style.Role == SemanticRole.SubjectExcerpt || style.Role == SemanticRole.DocumentType || style.Role == SemanticRole.NationalHeader || style.Role == SemanticRole.Motto || style.Role == SemanticRole.PartyHeader || style.Role == SemanticRole.RecipientHeader || style.Role == SemanticRole.RecipientItem || style.Role == SemanticRole.SignAuthority || style.Role == SemanticRole.SignTitle || style.Role == SemanticRole.SignerName || style.Role == SemanticRole.DirectRecipient))
			{
				comObj3.FirstLineIndent = 0f;
			}
			else if (defaultTypo != null && defaultTypo.FirstLineIndentMm > 0f)
			{
				comObj3.FirstLineIndent = defaultTypo.FirstLineIndentPt;
			}
		}
		catch
		{
		}
		finally
		{
			ComReleaser.Release(ref comObj3);
			ComReleaser.Release(ref comObj2);
			ComReleaser.Release(ref comObj);
		}
		try
		{
			FixOrphanWord(p);
		}
		catch
		{
		}
	}

	public static bool FixOrphanWord(Paragraph p)
	{
		if (p == null)
		{
			return false;
		}
		Range pRange = null;
		try
		{
			pRange = p.Range;
			if (pRange == null)
			{
				return false;
			}
			string text = pRange.Text;
			System.Collections.Generic.List<Tuple<int, int>> replacements = OrphanWordHelper.GetOrphanSpaceReplacements(text);
			if (replacements == null || replacements.Count == 0)
			{
				return false;
			}
			int pStart = pRange.Start;
			bool modified = false;
			foreach (Tuple<int, int> rep in replacements)
			{
				Range spaceRange = null;
				try
				{
					spaceRange = pRange.Duplicate;
					spaceRange.SetRange(pStart + rep.Item1, pStart + rep.Item1 + rep.Item2);
					string spaceText = spaceRange.Text ?? "";
					if (spaceText.Trim(' ', '\t', '\u00a0').Length == 0 && spaceRange.Fields.Count == 0)
					{
						spaceRange.Text = "\u00a0";
						modified = true;
					}
				}
				catch
				{
				}
				finally
				{
					ComReleaser.Release(ref spaceRange);
				}
			}
			return modified;
		}
		catch
		{
			return false;
		}
		finally
		{
			ComReleaser.Release(ref pRange);
		}
	}

	private static void ApplyLineSpacing(ParagraphFormat format, LineSpacingType rule, float value)
	{
		switch (rule)
		{
		case LineSpacingType.Exactly:
			format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
			format.LineSpacing = value;
			break;
		case LineSpacingType.AtLeast:
			format.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast;
			format.LineSpacing = value;
			break;
		case LineSpacingType.Multiple:
			format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
			format.LineSpacing = value * 12f;
			break;
		case LineSpacingType.OneAndHalf:
			format.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
			break;
		case LineSpacingType.Double:
			format.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
			break;
		default:
			format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			break;
		}
	}

	private static WdParagraphAlignment ConvertAlignment(AlignmentType align)
	{
		return align switch
		{
			AlignmentType.Center => WdParagraphAlignment.wdAlignParagraphCenter, 
			AlignmentType.Right => WdParagraphAlignment.wdAlignParagraphRight, 
			AlignmentType.Justify => WdParagraphAlignment.wdAlignParagraphJustify, 
			_ => WdParagraphAlignment.wdAlignParagraphLeft, 
		};
	}
}
