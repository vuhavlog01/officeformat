using System.Collections.Generic;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Builders;

public static class SignerBlockBuilder
{
	private static bool IsBlankParagraph(Paragraph p)
	{
		if (p == null)
		{
			return false;
		}
		try
		{
			if ((bool)(dynamic)p.Range.get_Information(WdInformation.wdWithInTable))
			{
				return false;
			}
			string text = (p.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Replace("\v", "")
				.Trim();
			return text.Length == 0;
		}
		catch
		{
			return false;
		}
	}

	private static void EnsureTrailingBlankParagraphs(Document doc)
	{
		try
		{
			for (int i = 0; i < 2; i++)
			{
				int count = doc.Paragraphs.Count;
				Paragraph p = doc.Paragraphs[count];
				Paragraph paragraph = ((count >= 2) ? doc.Paragraphs[count - 1] : null);
				bool flag = IsBlankParagraph(p);
				bool flag2 = paragraph == null || IsBlankParagraph(paragraph);
				if (flag && flag2)
				{
					break;
				}
				doc.Content.InsertParagraphAfter();
				Paragraph last = doc.Paragraphs.Last;
				try
				{
					last.Format.FirstLineIndent = 0f;
					last.Format.SpaceBefore = 0f;
					last.Format.SpaceAfter = 0f;
				}
				catch
				{
				}
			}
		}
		catch
		{
		}
	}

	public static Table InsertOrUpdateSignerTable(Document doc, ProfileModel profile, SignerProfile signer, RecipientProfile recipient)
	{
		if (doc == null || profile == null)
		{
			return null;
		}
		Range comObj = null;
		Table table = null;
		try
		{
			EnsureTrailingBlankParagraphs(doc);
			object Start = doc.Content.End - 1;
			object End = doc.Content.End - 1;
			comObj = doc.Range(ref Start, ref End);
			Tables tables = doc.Tables;
			Range range = comObj;
			object DefaultTableBehavior = Missing.Value;
			object AutoFitBehavior = Missing.Value;
			table = tables.Add(range, 1, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
			table.Borders.Enable = 0;
			table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
			table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
			table.PreferredWidth = 100f;
			try
			{
				table.Columns[1].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
				table.Columns[1].PreferredWidth = 45f;
				table.Columns[2].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
				table.Columns[2].PreferredWidth = 55f;
			}
			catch { }
			Cell cell = table.Cell(1, 1);
			Range range2 = cell.Range;
			range2.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
			range2.ParagraphFormat.SpaceBefore = 0f;
			range2.ParagraphFormat.SpaceAfter = 0f;
			range2.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			string text = ((profile.RulePack == RulePackType.HD05_VPTW) ? "Nơi nhận:" : "Nơi nhận:");
			ComponentStyle componentStyle = profile.GetComponentStyle(SemanticRole.RecipientHeader);
			float size;
			if (componentStyle != null && componentStyle.SizePt > 0f)
			{
				size = componentStyle.SizePt;
			}
			else
			{
				size = ((profile.RulePack == RulePackType.HD05_VPTW) ? 14f : 12f);
			}
			bool flag = RulePackTraits.IsYouthUnion(profile.RulePack);
			bool flag2 = profile.RulePack != RulePackType.HD05_VPTW && !flag;
			List<string> list2;
			if (recipient == null)
			{
				List<string> list = new List<string>();
				list.Add("- Như trên;");
				list.Add("- Lưu: VT.");
				list2 = list;
			}
			else
			{
				list2 = recipient.GetFormattedItems();
			}
			List<string> list3 = list2;
			string text2 = text + "\r" + string.Join("\r", list3.ToArray());
			range2.Text = text2;
			bool isParty = profile.RulePack == RulePackType.HD05_VPTW;
			Range range3 = range2.Paragraphs[1].Range;
			range3.Font.Size = size;
			range3.Font.Bold = (isParty ? 0 : (-1));
			range3.Font.Italic = (flag2 ? (-1) : 0);
			range3.Font.Underline = (isParty ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
			ComponentStyle componentStyle2 = profile.GetComponentStyle(SemanticRole.RecipientItem);
			float size2;
			if (componentStyle2 != null && componentStyle2.SizePt > 0f)
			{
				size2 = componentStyle2.SizePt;
			}
			else
			{
				size2 = (isParty ? 12f : 11f);
			}
			int count = range2.Paragraphs.Count;
			for (int i = 2; i <= count; i++)
			{
				Range range4 = range2.Paragraphs[i].Range;
				range4.Font.Size = size2;
				range4.Font.Bold = 0;
				range4.Font.Italic = 0;
				range4.Font.Underline = WdUnderline.wdUnderlineNone;
			}
			Cell cell2 = table.Cell(1, 2);
			Range range5 = cell2.Range;
			range5.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range5.ParagraphFormat.SpaceBefore = 0f;
			range5.ParagraphFormat.SpaceAfter = 0f;
			range5.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			string text3;
			if (signer != null && !string.IsNullOrEmpty(signer.AuthorityTitle))
			{
				text3 = signer.AuthorityTitle;
			}
			else
			{
				text3 = ((signer != null && !string.IsNullOrEmpty(signer.Position)) ? signer.Position : "CHỦ TỊCH");
			}
			string text4 = ((signer != null) ? signer.DisplaySignerName : "Họ và tên");
			if (flag)
			{
				text3 = text3.Replace("\r\n", "\r").Replace("\n", "\r");
			}
			if (isParty)
			{
				text3 = text3.Replace("\r\n", "\r").Replace("\n", "\r");
			}
			string text5 = text3 + "\r\r\r\r" + text4;
			range5.Text = text5;
			ComponentStyle componentStyle3 = profile.GetComponentStyle(SemanticRole.SignerName);
			float size3 = ((componentStyle3 != null && componentStyle3.SizePt > 0f) ? componentStyle3.SizePt : 14f);
			range5.Font.Size = size3;
			range5.Font.Bold = -1;
			if (flag || isParty)
			{
				ComponentStyle componentStyle4 = profile.GetComponentStyle(SemanticRole.SignTitle);
				int count2 = range5.Paragraphs.Count;
				for (int j = 1; j <= count2; j++)
				{
					Range range6 = range5.Paragraphs[j].Range;
					string text6 = (range6.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (text6.Length != 0 && j != count2 && !RegexPatterns.SignAuthority.IsMatch(text6))
					{
						range6.Font.Bold = ((componentStyle4 != null && componentStyle4.Bold) ? (-1) : 0);
					}
				}
			}
			return table;
		}
		catch
		{
			return table;
		}
		finally
		{
			ComReleaser.Release(ref comObj);
		}
	}
}
