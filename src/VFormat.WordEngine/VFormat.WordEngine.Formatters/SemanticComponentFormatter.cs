using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.WordEngine.Builders;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Formatters;

public static class SemanticComponentFormatter
{
	private sealed class Target
	{
		public int Start;

		public int End;

		public SemanticRole Role;

		public bool MixedParagraph;

		public string Text;
	}

	private static readonly Regex Agency = new Regex("^(HỘI\\b|BAN\\b|ỦY BAN\\b|UBND\\b|HĐND\\b|BỘ\\b|SỞ\\b|CỤC\\b|TỔNG CỤC\\b|PHÒNG\\b|VĂN PHÒNG\\b|CHI NHÁNH\\b|TRUNG TÂM\\b|VIỆN\\b|TRƯỜNG\\b|ĐẢNG\\b|CHI BỘ\\b|TỈNH ỦY\\b|THÀNH ỦY\\b|HUYỆN ỦY\\b|BCH\\b|TỈNH ĐOÀN\\b|THÀNH ĐOÀN\\b|HUYỆN ĐOÀN\\b|QUẬN ĐOÀN\\b|THỊ ĐOÀN\\b|ĐOÀN\\b|CHI ĐOÀN\\b|LIÊN CHI ĐOÀN\\b|PHÂN ĐOÀN\\b|ĐẠI HỘI\\b|QUÂN\\b|SƯ ĐOÀN\\b|LỮ ĐOÀN\\b|TRUNG ĐOÀN\\b|TIỂU ĐOÀN\\b)", RegexOptions.IgnoreCase);

	private static readonly Regex Date = new Regex("^[^,\\r]+,\\s*ngày\\b.*\\btháng\\b.*\\bnăm\\b", RegexOptions.IgnoreCase);

	private static readonly Regex Name = new Regex("^[\\p{L}][\\p{L}\\p{M} .-]*$");

	public static int RawHeaderEnd(Document doc)
	{
		bool flag = false;
		int num = 0;
		foreach (Paragraph paragraph in doc.Paragraphs)
		{
			OperationProgress.Step();
			if (++num > 25)
			{
				break;
			}
			if (paragraph.Range.Tables.Count <= 0)
			{
				string text = paragraph.Range.Text.Trim();
				if (flag && RegexPatterns.DocumentType.IsMatch(text) && text == text.ToUpperInvariant() && text.Length <= 60)
				{
					return paragraph.Range.Start;
				}
				flag |= RegexPatterns.NationalHeader.IsMatch(text) || RegexPatterns.PartyHeader.IsMatch(text) || RegexPatterns.YouthUnionHeader.IsMatch(text);
			}
		}
		return 0;
	}

	public static int Apply(Document doc, ProfileModel profile, Range scope = null, bool reflowTables = true, bool signatureOnly = false)
	{
		string text = doc.Content.Text;
		List<Target> list = new List<Target>();
		List<Table> list2 = new List<Table>();
		int num = RawHeaderEnd(doc);
		if (num > 0 && !signatureOnly)
		{
			bool recipients = false;
			bool signer = false;
			object Start = 0;
			object End = num;
			Parse(doc.Range(ref Start, ref End), list, header: true, signature: false, 0, ref recipients, ref signer);
		}
		foreach (Table table in doc.Tables)
		{
			OperationProgress.Step();
			string text2 = table.Range.Text;
			bool flag = table.Range.Start < 2000 && (RegexPatterns.NationalHeader.IsMatch(text2) || RegexPatterns.PartyHeader.IsMatch(text2) || RegexPatterns.YouthUnionHeader.IsMatch(text2));
			if (flag && (reflowTables || signatureOnly))
			{
				continue;
			}
			bool flag2 = Regex.IsMatch(text2, "(^|[\\r\\n\\v\\a])\\s*Nơi\\s+nhận\\s*:?", RegexOptions.IgnoreCase);
			if (!flag2 && !flag)
			{
				continue;
			}
			if (!SignatureRebuilder.HasAtLeastTwoColumns(table))
			{
				continue;
			}
			list2.Add(table);
			bool recipients2 = false;
			bool signer2 = false;
			foreach (Cell cell in table.Range.Cells)
			{
				Parse(cell.Range, list, flag, flag2, cell.ColumnIndex, ref recipients2, ref signer2);
			}
		}
		bool flag3 = false;
		bool recipients3 = false;
		bool signer3 = false;
		int num2 = 0;
		int count = doc.Paragraphs.Count;
		foreach (Paragraph paragraph in doc.Paragraphs)
		{
			OperationProgress.Step();
			num2++;
			if (paragraph.Range.Tables.Count <= 0)
			{
				string text3 = paragraph.Range.Text.Trim();
				if (RegexPatterns.AppendixTitle.IsMatch(text3))
				{
					flag3 = (recipients3 = (signer3 = false));
				}
				if (!flag3 && (signatureOnly ? (num2 >= Math.Max(1, count - 12)) : (num2 > count / 2)) && (RegexPatterns.RecipientHeader.IsMatch(text3) || RegexPatterns.SignAuthority.IsMatch(text3) || RegexPatterns.IsStandaloneSignPosition(text3)))
				{
					flag3 = true;
				}
				if (flag3)
				{
					Parse(paragraph.Range, list, header: false, signature: true, 0, ref recipients3, ref signer3);
				}
			}
		}
		foreach (Target item in list)
		{
			if ((scope != null && (item.Start < scope.Start || item.End > scope.End)) || !profile.ComponentStyles.TryGetValue(item.Role, out var value))
			{
				continue;
			}
			object Start2 = item.Start;
			object End2 = item.End;
			Range range = doc.Range(ref Start2, ref End2);
			Range comObj = range.Paragraphs[1].Range;
			string text4 = (comObj.Text ?? "").Trim('\r', '\n', '\a', ' ', '\t', '\u00a0');
			bool isWholeParagraph = text4 == (range.Text ?? "").Trim();
			if (isWholeParagraph && (scope == null || comObj.End <= scope.End))
			{
				range.SetRange(range.Start, comObj.End);
			}
			ComReleaser.Release(ref comObj);
			range.Font.Name = profile.Typography.FontName;
			range.Font.Size = value.SizePt;
			range.Font.Bold = (value.Bold ? (-1) : 0);
			range.Font.Italic = (value.Italic ? (-1) : 0);
			range.Font.ColorIndex = WdColorIndex.wdBlack;
			range.Font.Underline = ((value.UnderlineRule == UnderlineRuleType.SolidFullWidth && item.Role != SemanticRole.Motto && item.Role != SemanticRole.PartyHeader) ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
			if (item.MixedParagraph && !isWholeParagraph)
			{
				continue;
			}
			ParagraphFormat paragraphFormat = range.ParagraphFormat;
			paragraphFormat.Alignment = ((value.Alignment == AlignmentType.Center) ? WdParagraphAlignment.wdAlignParagraphCenter : ((value.Alignment == AlignmentType.Right) ? WdParagraphAlignment.wdAlignParagraphRight : WdParagraphAlignment.wdAlignParagraphLeft));
			float leftIndent = (paragraphFormat.RightIndent = 0f);
			paragraphFormat.LeftIndent = leftIndent;
			if (value.FirstLineIndentPt.HasValue)
			{
				paragraphFormat.FirstLineIndent = value.FirstLineIndentPt.Value;
			}
			if (value.SpacingBeforePt.HasValue)
			{
				paragraphFormat.SpaceBefore = value.SpacingBeforePt.Value;
			}
			if (value.SpacingAfterPt.HasValue)
			{
				paragraphFormat.SpaceAfter = value.SpacingAfterPt.Value;
			}
			LineSpacingType? lineSpacingRule = value.LineSpacingRule;
			if (lineSpacingRule.GetValueOrDefault() == LineSpacingType.Single && lineSpacingRule.HasValue)
			{
				paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			}
			else
			{
				LineSpacingType? lineSpacingRule2 = value.LineSpacingRule;
				if (lineSpacingRule2 == LineSpacingType.Exactly && lineSpacingRule2.HasValue && value.LineSpacingValue.HasValue)
				{
					paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
					paragraphFormat.LineSpacing = value.LineSpacingValue.Value;
				}
				else
				{
					LineSpacingType? lineSpacingRule3 = value.LineSpacingRule;
					if (lineSpacingRule3 == LineSpacingType.Multiple && lineSpacingRule3.HasValue && value.LineSpacingValue.HasValue)
					{
						paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;
						paragraphFormat.LineSpacing = doc.Application.LinesToPoints(value.LineSpacingValue.Value);
					}
				}
			}
			if (value.KeepWithNext.HasValue)
			{
				paragraphFormat.KeepWithNext = (value.KeepWithNext.Value ? (-1) : 0);
			}
			if (value.KeepLinesTogether.HasValue)
			{
				paragraphFormat.KeepTogether = (value.KeepLinesTogether.Value ? (-1) : 0);
			}
		}
		foreach (Table item2 in list2)
		{
			if (reflowTables && (scope == null || (item2.Range.Start >= scope.Start && item2.Range.End <= scope.End)))
			{
				OperationProgress.Step();
				item2.Rows.SetLeftIndent(0f, WdRulerStyle.wdAdjustNone);
				item2.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
				item2.AllowAutoFit = false;
				item2.Rows.AllowBreakAcrossPages = 0;
			}
		}
		if (doc.Content.Text != text)
		{
			throw new InvalidOperationException("Kiểm tra bảo toàn nội dung thất bại khi định dạng thành phần.");
		}
		return list.Count;
	}

	private static void Parse(Range source, List<Target> plan, bool header, bool signature, int column, ref bool recipients, ref bool signer)
	{
		string input = source.Text ?? "";
		List<Target> list = new List<Target>();
		foreach (Match item2 in Regex.Matches(input, "[^\\r\\n\\v\\a\\t]+"))
		{
			string text = item2.Value.Trim();
			if (text.Length == 0)
			{
				continue;
			}
			SemanticRole semanticRole = SemanticRole.Unknown;
			if (header)
			{
				if (RegexPatterns.NationalHeader.IsMatch(text))
				{
					semanticRole = SemanticRole.NationalHeader;
				}
				else if (RegexPatterns.PartyHeader.IsMatch(text) || RegexPatterns.YouthUnionHeader.IsMatch(text))
				{
					semanticRole = SemanticRole.PartyHeader;
				}
				else if (RegexPatterns.Motto.IsMatch(text))
				{
					semanticRole = SemanticRole.Motto;
				}
				else if (RegexPatterns.DocumentNumberGeneral.IsMatch(text))
				{
					semanticRole = SemanticRole.DocumentNumber;
				}
				else if (Date.IsMatch(text))
				{
					semanticRole = SemanticRole.DocumentDate;
				}
				else if (RegexPatterns.DispatchSubject.IsMatch(text))
				{
					semanticRole = SemanticRole.DispatchSubject;
				}
				else if (column != 2 && text.Length <= 160 && (Agency.IsMatch(text) || RegexPatterns.YouthUnionAgency.IsMatch(text) || (list.Count > 0 && HeaderRebuilder.IsSingleAgencySplitInTwo(list[list.Count - 1].Text, text))))
				{
					semanticRole = SemanticRole.IssuingAgency;
				}
			}
			else if (signature)
			{
				if (RegexPatterns.RecipientHeader.IsMatch(text))
				{
					semanticRole = SemanticRole.RecipientHeader;
					recipients = true;
				}
				else if (RegexPatterns.SignAuthority.IsMatch(text))
				{
					semanticRole = SemanticRole.SignAuthority;
					signer = true;
				}
				else if (RegexPatterns.SignPosition.IsMatch(text))
				{
					semanticRole = SemanticRole.SignTitle;
					signer = true;
				}
				else if (recipients && text.StartsWith("-") && column != 2)
				{
					semanticRole = SemanticRole.RecipientItem;
				}
				else if (signer && column != 1 && text.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length <= 7 && Name.IsMatch(text))
				{
					semanticRole = (RegexPatterns.MilitaryRank.IsMatch(text) ? SemanticRole.SignRankName : SemanticRole.SignerName);
				}
			}
			if (semanticRole != SemanticRole.Unknown)
			{
				int num = item2.Value.Length - item2.Value.TrimStart().Length;
				Target target = new Target();
				target.Start = source.Start + item2.Index + num;
				target.End = source.Start + item2.Index + num + text.Length;
				target.Role = semanticRole;
				target.MixedParagraph = column == 0;
				target.Text = text;
				Target item = target;
				plan.Add(item);
				if (semanticRole == SemanticRole.IssuingAgency)
				{
					list.Add(item);
				}
			}
		}
		int num2 = list.Count - 1;
		while (num2 > 0 && HeaderRebuilder.IsSingleAgencySplitInTwo(list[num2 - 1].Text, list[num2].Text))
		{
			num2--;
		}
		for (int i = 0; i < num2; i++)
		{
			list[i].Role = SemanticRole.SuperiorAgency;
		}
	}
}
