using System;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;
using VFormat.WordEngine.Formatters;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Builders;

public static class NationalHeaderBuilder
{
	public static Table InsertOrUpdateHeaderTable(Document doc, ProfileModel profile, AgencyProfile agency, string documentNumber = null, string documentDate = null)
	{
		if (doc == null || profile == null)
		{
			return null;
		}
		Range comObj = null;
		Table table = null;
		try
		{
			if (doc.Tables.Count > 0)
			{
				Table table2 = doc.Tables[1];
				if (table2.Columns.Count == 2)
				{
					string text = table2.Range.Text;
					if (RegexPatterns.NationalHeader.IsMatch(text) || RegexPatterns.PartyHeader.IsMatch(text) || RegexPatterns.YouthUnionHeader.IsMatch(text))
					{
						table = table2;
					}
				}
			}
			if (table == null)
			{
				object Start = 0;
				object End = 0;
				comObj = doc.Range(ref Start, ref End);
				Tables tables = doc.Tables;
				Range range = comObj;
				object DefaultTableBehavior = Missing.Value;
				object AutoFitBehavior = Missing.Value;
				table = tables.Add(range, 2, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
			}
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
			range2.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range2.ParagraphFormat.SpaceBefore = 0f;
			range2.ParagraphFormat.SpaceAfter = 0f;
			range2.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			string text2 = ((agency != null) ? (agency.ParentAgency ?? "") : "");
			string text3 = ((agency != null) ? (agency.AgencyName ?? "") : "");
			ComponentStyle componentStyle = profile.GetComponentStyle(SemanticRole.IssuingAgency);
			bool flag = RulePackTraits.UsesOrganizationHeader(profile.RulePack);
			float num;
			if (componentStyle != null && componentStyle.SizePt > 0f)
			{
				num = componentStyle.SizePt;
			}
			else
			{
				num = (flag ? 14f : 13f);
			}
			ComponentStyle componentStyle2 = profile.GetComponentStyle(SemanticRole.SuperiorAgency);
			float size = ((componentStyle2 != null && componentStyle2.SizePt > 0f) ? componentStyle2.SizePt : num);
			string text4 = (flag ? RulePackTraits.GetAgencySeparator(profile.RulePack) : string.Empty);
			string text5 = ((flag && !string.IsNullOrEmpty(text4)) ? ("\r" + text4) : string.Empty);
			if (!string.IsNullOrEmpty(text2))
			{
				range2.Text = text2.ToUpperInvariant() + "\r" + text3.ToUpperInvariant() + text5;
				range2.Paragraphs[1].Range.Font.Size = size;
				range2.Paragraphs[1].Range.Font.Bold = 0;
				range2.Paragraphs[2].Range.Font.Size = num;
				range2.Paragraphs[2].Range.Font.Bold = -1;
			}
			else
			{
				range2.Text = text3.ToUpperInvariant() + text5;
				range2.Font.Size = num;
				range2.Font.Bold = -1;
			}
			if (text5.Length > 0)
			{
				Range range3 = cell.Range.Paragraphs[cell.Range.Paragraphs.Count].Range;
				range3.Font.Size = num;
				range3.Font.Bold = 0;
			}
			Cell cell2 = table.Cell(1, 2);
			Range range4 = cell2.Range;
			range4.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range4.ParagraphFormat.SpaceBefore = 0f;
			range4.ParagraphFormat.SpaceAfter = 0f;
			range4.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			if (flag)
			{
				ComponentStyle componentStyle3 = profile.GetComponentStyle(SemanticRole.PartyHeader);
				float size2 = ((componentStyle3 != null && componentStyle3.SizePt > 0f) ? componentStyle3.SizePt : 15f);
				range4.Text = RulePackTraits.GetOrganizationHeaderText(profile.RulePack);
				range4.Font.Size = size2;
				range4.Font.Bold = -1;
				range4.Font.Underline = WdUnderline.wdUnderlineNone;
			}
			else
			{
				ComponentStyle componentStyle4 = profile.GetComponentStyle(SemanticRole.NationalHeader);
				float size3;
				if (componentStyle4 != null && componentStyle4.SizePt > 0f)
				{
					size3 = componentStyle4.SizePt;
				}
				else
				{
					size3 = ((profile.RulePack == RulePackType.TT68_2023_BQP) ? 13f : 12f);
				}
				ComponentStyle componentStyle5 = profile.GetComponentStyle(SemanticRole.Motto);
				float size4 = ((componentStyle5 != null && componentStyle5.SizePt > 0f) ? componentStyle5.SizePt : 13f);
				range4.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\rĐộc lập - Tự do - Hạnh phúc";
				range4.Paragraphs[1].Range.Font.Size = size3;
				range4.Paragraphs[1].Range.Font.Bold = -1;
				range4.Paragraphs[2].Range.Font.Size = size4;
				range4.Paragraphs[2].Range.Font.Bold = -1;
				range4.Paragraphs[2].Range.Font.Underline = WdUnderline.wdUnderlineNone;
			}
			Cell cell3 = table.Cell(2, 1);
			Range range5 = cell3.Range;
			range5.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range5.ParagraphFormat.SpaceBefore = 4f;
			range5.ParagraphFormat.SpaceAfter = 0f;
			range5.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			bool flag2 = flag;
			ComponentStyle componentStyle6 = profile.GetComponentStyle(SemanticRole.DocumentNumber);
			float size5;
			if (componentStyle6 != null && componentStyle6.SizePt > 0f)
			{
				size5 = componentStyle6.SizePt;
			}
			else
			{
				size5 = (flag2 ? 14f : 13f);
			}
			range5.Font.Size = size5;
			range5.Font.Bold = 0;
			bool flag3 = RulePackTraits.IsYouthUnion(profile.RulePack);
			string arg;
			if (agency != null && !string.IsNullOrEmpty(agency.CodePrefix))
			{
				arg = agency.CodePrefix;
			}
			else if (flag3)
			{
				arg = "ĐTN";
			}
			else
			{
				arg = (flag2 ? "ĐU" : "UBND");
			}
			string text6;
			if (string.IsNullOrEmpty(documentNumber))
			{
				if (flag3)
				{
					text6 = $"Số:     -.../{arg}";
				}
				else
				{
					text6 = (flag2 ? $"Số     -{arg}" : $"Số: .../{arg}");
				}
			}
			else
			{
				text6 = documentNumber;
			}
			range5.Text = text6;
			Cell cell4 = table.Cell(2, 2);
			Range range6 = cell4.Range;
			range6.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range6.ParagraphFormat.SpaceBefore = 4f;
			range6.ParagraphFormat.SpaceAfter = 0f;
			range6.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			ComponentStyle componentStyle7 = profile.GetComponentStyle(SemanticRole.DocumentDate);
			float size6;
			if (componentStyle7 != null && componentStyle7.SizePt > 0f)
			{
				size6 = componentStyle7.SizePt;
			}
			else
			{
				size6 = ((profile.RulePack == RulePackType.TT68_2023_BQP || flag3) ? 13f : 14f);
			}
			range6.Font.Size = size6;
			range6.Font.Italic = -1;
			range6.Font.Bold = 0;
			string text7 = ((agency != null && !string.IsNullOrEmpty(agency.Location)) ? agency.Location : "...");
			DateTime today = DateTime.Today;
			string text8 = ((!string.IsNullOrEmpty(documentDate)) ? documentDate : $"{text7}, ngày {today.Day:D2} tháng {today.Month:D2} năm {today.Year}");
			range6.Text = text8;
			HeaderRebuilder.FormatExistingHeaderTable(table, profile.RulePack, profile, skipRebuildCheck: true);
			try
			{
				float num2 = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;
				float col1Width = num2 * 0.45f;
				float col2Width = num2 * 0.55f;
				UnderlineFormatter.EnsureHeaderShapeLines(table, profile.RulePack, col1Width, col2Width);
			}
			catch
			{
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
