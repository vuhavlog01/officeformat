using System;
using System.Drawing;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Rules;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Formatters;

public static class UnderlineFormatter
{
	public static void ApplyUnderlineRule(Paragraph p, UnderlineRuleType rule)
	{
		if (p == null)
		{
			return;
		}
		if (rule == UnderlineRuleType.None)
		{
			try
			{
				p.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				p.Range.ParagraphFormat.Borders.Enable = 0;
				p.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
				return;
			}
			catch
			{
				return;
			}
		}
		Range comObj = null;
		try
		{
			comObj = p.Range;
			string value = ((comObj.Text != null) ? comObj.Text.Trim() : "");
			if (string.IsNullOrEmpty(value))
			{
				return;
			}
			switch (rule)
			{
			case UnderlineRuleType.SolidFullWidth:
				p.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
				try
				{
					p.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					break;
				}
				catch
				{
					break;
				}
			case UnderlineRuleType.SolidOneThirdToHalf:
				UnderlineMiddleOneThirdToHalf(p);
				try
				{
					p.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					break;
				}
				catch
				{
					break;
				}
			case UnderlineRuleType.FiveDashes:
				EnsureFiveDashes(p);
				break;
			default:
				try
				{
					p.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					break;
				}
				catch
				{
					break;
				}
			}
		}
		catch
		{
		}
		finally
		{
			ComReleaser.Release(ref comObj);
		}
	}

	public static string CreateSolidSeparator(int targetCharCount = 7)
	{
		if (targetCharCount < 5)
		{
			targetCharCount = 5;
		}
		if (targetCharCount > 25)
		{
			targetCharCount = 25;
		}
		return new string('—', targetCharCount);
	}

	public static void UnderlineMiddleOneThirdToHalf(Paragraph p)
	{
		if (p == null)
		{
			return;
		}
		try
		{
			p.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			try
			{
				p.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
			}
			catch
			{
			}
			string text = p.Range.Text ?? "";
			string text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			if (text2.Length == 0)
			{
				return;
			}
			int targetCharCount = Math.Max(7, Math.Min(16, (int)Math.Round((double)text2.Length * 0.35)));
			string text3 = CreateSolidSeparator(targetCharCount);
			object Count = Missing.Value;
			Paragraph paragraph = p.Next(ref Count);
			if (paragraph != null)
			{
				string text4 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (IsSeparator(text4))
				{
					paragraph.Range.Text = text3 + "\r";
					FormatSeparatorParagraph(paragraph);
				}
			}
		}
		catch
		{
		}
	}

	public static void EnsureSolidSeparator(Paragraph p, int charCount = 7)
	{
		if (p == null)
		{
			return;
		}
		try
		{
			p.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			object Count = Missing.Value;
			Paragraph paragraph = p.Next(ref Count);
			string text = CreateSolidSeparator(charCount);
			if (paragraph != null)
			{
				string text2 = (paragraph.Range.Text ?? "").Trim();
				if (IsSeparator(text2))
				{
					paragraph.Range.Text = text + "\r";
					FormatSeparatorParagraph(paragraph);
					return;
				}
			}
			Range range = p.Range;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range.Collapse(ref Direction);
			range.Text = text + "\r";
			range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			range.ParagraphFormat.FirstLineIndent = 0f;
			range.ParagraphFormat.LeftIndent = 0f;
			range.ParagraphFormat.RightIndent = 0f;
			range.ParagraphFormat.SpaceBefore = 0f;
			range.ParagraphFormat.SpaceAfter = 2f;
			range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			range.Font.Name = "Times New Roman";
			range.Font.Size = 10f;
			range.Font.Bold = -1;
			range.Font.Italic = 0;
			range.Font.Underline = WdUnderline.wdUnderlineNone;
			try
			{
				range.ParagraphFormat.Borders.Enable = 0;
			}
			catch
			{
			}
		}
		catch
		{
		}
	}

	private static bool IsSeparator(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		string text2 = text.Replace("-", "").Replace("_", "").Replace("=", "")
			.Replace("*", "")
			.Replace("—", "")
			.Replace("—", "")
			.Replace("―", "")
			.Replace("─", "")
			.Replace("~", "")
			.Replace(" ", "");
		if (text2.Length == 0)
		{
			return text.Length >= 2;
		}
		return false;
	}

	private static void FormatSeparatorParagraph(Paragraph p)
	{
		try
		{
			p.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			p.Range.ParagraphFormat.FirstLineIndent = 0f;
			p.Range.ParagraphFormat.LeftIndent = 0f;
			p.Range.ParagraphFormat.RightIndent = 0f;
			p.Range.ParagraphFormat.SpaceBefore = 0f;
			p.Range.ParagraphFormat.SpaceAfter = 2f;
			p.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			p.Range.Font.Name = "Times New Roman";
			p.Range.Font.Size = 10f;
			p.Range.Font.Bold = -1;
			p.Range.Font.Italic = 0;
			p.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			try
			{
				p.Range.ParagraphFormat.Borders.Enable = 0;
			}
			catch
			{
			}
			foreach (Border border in p.Borders)
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
	}

	private static void EnsureFiveDashes(Paragraph p)
	{
		try
		{
			object Count = Missing.Value;
			Paragraph paragraph = p.Next(ref Count);
			if (paragraph == null)
			{
				return;
			}
			string text = (paragraph.Range.Text ?? "").Trim();
			if (!text.Contains("- - - - -") && !text.Contains("-----"))
			{
				string text2 = text.ToUpperInvariant();
				if (text2.StartsWith("I.") || text2.StartsWith("1.") || text2.StartsWith("CĂN CỨ") || text2.StartsWith("CAN CU") || text2.StartsWith("ĐIỀU") || text2.StartsWith("DIEU"))
				{
					Range range = p.Range;
					object Direction = WdCollapseDirection.wdCollapseEnd;
					range.Collapse(ref Direction);
					range.Text = "-----\r";
					range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					range.ParagraphFormat.FirstLineIndent = 0f;
					range.ParagraphFormat.SpaceBefore = 2f;
					range.ParagraphFormat.SpaceAfter = 6f;
					range.Font.Name = "Times New Roman";
					range.Font.Size = 14f;
					range.Font.Bold = -1;
					range.Font.Underline = WdUnderline.wdUnderlineNone;
				}
			}
		}
		catch
		{
		}
	}

	private static void DisableShapeEffects(Shape shape)
	{
		if (shape == null)
		{
			return;
		}
		try
		{
			((dynamic)shape).Shadow.Visible = 0;
		}
		catch
		{
		}
		try
		{
			((dynamic)shape).Reflection.Type = 0;
		}
		catch
		{
		}
		try
		{
			((dynamic)shape).Glow.Radius = 0;
		}
		catch
		{
		}
		try
		{
			((dynamic)shape).SoftEdge.Type = 0;
		}
		catch
		{
		}
	}

	public static void EnsureHeaderShapeLines(Table tbl, RulePackType rulePack, float col1Width, float col2Width)
	{
		if (tbl == null)
		{
			return;
		}
		try
		{
			Document document = tbl.Range.Document;
			if (document == null)
			{
				return;
			}
			for (int num = document.Shapes.Count; num >= 1; num--)
			{
				try
				{
					Shapes shapes = document.Shapes;
					object Index = num;
					Shape shape = shapes[ref Index];
					string text = shape.Name ?? "";
					if (text == "VFormat_Motto_Line" || text == "VFormat_Party_Line")
					{
						shape.Delete();
					}
				}
				catch
				{
				}
			}
			float num2 = 56.7f;
			try
			{
				num2 = document.PageSetup.LeftMargin;
			}
			catch
			{
			}
			float num3 = num2 + col1Width;
			float num4 = num3 + col2Width / 2f;
			if (RulePackTraits.UsesOrganizationHeader(rulePack))
			{
				Cell cell = tbl.Cell(1, 2);
				Paragraph paragraph = null;
				for (int i = 1; i <= cell.Range.Paragraphs.Count; i++)
				{
					Paragraph paragraph2 = cell.Range.Paragraphs[i];
					string text2 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
						.Trim();
					if (text2.Contains("ĐẢNG CỘNG SẢN") || text2.Contains("DANG CONG SAN") || RegexPatterns.YouthUnionHeader.IsMatch(text2) || i == 1)
					{
						paragraph = paragraph2;
						break;
					}
				}
				if (paragraph == null)
				{
					return;
				}
				paragraph.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				paragraph.Range.Underline = WdUnderline.wdUnderlineNone;
				string text3 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (string.IsNullOrEmpty(text3))
				{
					return;
				}
				string[] array = text3.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
				float num5 = 0f;
				if (array.Length > 1)
				{
					string[] array2 = array;
					foreach (string text4 in array2)
					{
						float num6 = MeasureTextWidthPt(text4.Trim(), paragraph.Range.Font);
						if (num6 > num5)
						{
							num5 = num6;
						}
					}
				}
				else
				{
					num5 = MeasureTextWidthPt(text3, paragraph.Range.Font);
				}
				float num7 = ((num5 > 0f) ? num5 : MeasureTextWidthPt(text3, paragraph.Range.Font));
				if (num7 > col2Width - 4f)
				{
					num7 = col2Width - 4f;
				}
				float left = num4 - num7 / 2f;
				float right = num4 + num7 / 2f;
				Shape shape2 = AddParagraphUnderline(document, paragraph, left, right);
				shape2.Line.Weight = 0.75f;
				shape2.Line.ForeColor.RGB = 0;
				shape2.Name = "VFormat_Party_Line";
				DisableShapeEffects(shape2);
				return;
			}
			Cell cell2 = tbl.Cell(1, 2);
			Paragraph paragraph3 = null;
			for (int k = 1; k <= cell2.Range.Paragraphs.Count; k++)
			{
				Paragraph paragraph4 = cell2.Range.Paragraphs[k];
				string text5 = (paragraph4.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (text5.Contains("Độc lập") || text5.Contains("Doc lap") || (k == 2 && cell2.Range.Paragraphs.Count >= 2))
				{
					paragraph3 = paragraph4;
					break;
				}
			}
			if (paragraph3 == null)
			{
				return;
			}
			paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			paragraph3.Range.Underline = WdUnderline.wdUnderlineNone;
			string text6 = (paragraph3.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			if (!string.IsNullOrEmpty(text6))
			{
				float num8 = MeasureTextWidthPt(text6, paragraph3.Range.Font);
				if (num8 > col2Width - 4f)
				{
					num8 = col2Width - 4f;
				}
				float left2 = num4 - num8 / 2f;
				float right2 = num4 + num8 / 2f;
				Shape shape3 = AddParagraphUnderline(document, paragraph3, left2, right2);
				shape3.Line.Weight = 0.75f;
				shape3.Line.ForeColor.RGB = 0;
				shape3.Name = "VFormat_Motto_Line";
				DisableShapeEffects(shape3);
			}
		}
		catch
		{
		}
	}

	public static void EnsurePersonalMottoShapeLine(Document doc, Paragraph pMotto)
	{
		if (doc == null || pMotto == null)
		{
			return;
		}
		try
		{
			pMotto.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			pMotto.Range.Underline = WdUnderline.wdUnderlineNone;
			for (int num = doc.Shapes.Count; num >= 1; num--)
			{
				try
				{
					Shapes shapes = doc.Shapes;
					object Index = num;
					Shape shape = shapes[ref Index];
					if (shape.Name == "VFormat_Motto_Line")
					{
						shape.Delete();
					}
				}
				catch
				{
				}
			}
			string text = (pMotto.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				float num2 = MeasureTextWidthPt(text, pMotto.Range.Font);
				float num3 = 595.3f;
				float num4 = 85.05f;
				float num5 = 56.7f;
				try
				{
					num3 = doc.PageSetup.PageWidth;
					num4 = doc.PageSetup.LeftMargin;
					num5 = doc.PageSetup.RightMargin;
				}
				catch
				{
				}
				float num6 = num3 - num4 - num5;
				float num7 = num4 + num6 / 2f;
				float left = num7 - num2 / 2f;
				float right = num7 + num2 / 2f;
				Shape shape2 = AddParagraphUnderline(doc, pMotto, left, right);
				shape2.Line.Weight = 0.75f;
				shape2.Line.ForeColor.RGB = 0;
				shape2.Name = "VFormat_Motto_Line";
				DisableShapeEffects(shape2);
			}
		}
		catch
		{
		}
	}

	private static Shape AddParagraphUnderline(Document doc, Paragraph paragraph, float left, float right)
	{
		Range duplicate = paragraph.Range.Duplicate;
		object Direction = WdCollapseDirection.wdCollapseStart;
		duplicate.Collapse(ref Direction);
		object Anchor = duplicate;
		float num = paragraph.Range.Font.Size;
		if (num <= 0f || num > 200f)
		{
			num = 14f;
		}
		int num2 = Math.Max(1, paragraph.Range.ComputeStatistics(WdStatistic.wdStatisticLines));
		if (paragraph.Range.Tables.Count > 0)
		{
			Cell cell = paragraph.Range.Cells[1];
			float num3 = cell.Width - cell.LeftPadding - cell.RightPadding;
			int num4 = 0;
			string text = (paragraph.Range.Text ?? "").TrimEnd('\r', '\a');
			string[] array = text.Split('\r', '\n', '\v');
			foreach (string text2 in array)
			{
				string text3 = "";
				num4++;
				string[] array2 = text2.Split(new char[2] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string text4 in array2)
				{
					string text5 = ((text3.Length == 0) ? text4 : (text3 + " " + text4));
					if (text3.Length > 0 && MeasureTextWidthPt(text5, paragraph.Range.Font) > num3)
					{
						num4++;
						text3 = text4;
					}
					else
					{
						text3 = text5;
					}
				}
			}
			num2 = Math.Max(num2, num4);
		}
		float num5 = num * 1.15f;
		ParagraphFormat format = paragraph.Format;
		if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly || format.LineSpacingRule == WdLineSpacing.wdLineSpaceAtLeast)
		{
			num5 = Math.Max(num5, format.LineSpacing);
		}
		else if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceDouble)
		{
			num5 *= 2f;
		}
		else if (format.LineSpacingRule == WdLineSpacing.wdLineSpace1pt5)
		{
			num5 *= 1.5f;
		}
		else if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceMultiple)
		{
			num5 *= Math.Max(1f, format.LineSpacing / 12f);
		}
		Shape shape = doc.Shapes.AddLine(0f, 0f, right - left, 0f, ref Anchor);
		shape.WrapFormat.Type = WdWrapType.wdWrapNone;
		shape.LayoutInCell = -1;
		shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
		shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
		shape.LockAnchor = -1;
		shape.Left = left;
		if (paragraph.Range.Tables.Count > 0)
		{
			Cell cell2 = paragraph.Range.Cells[1];
			shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn;
			float num6 = cell2.Width - cell2.LeftPadding - cell2.RightPadding;
			shape.Width = Math.Min(shape.Width, num6);
			shape.Left = Math.Max(0f, (num6 - shape.Width) / 2f);
		}
				int lineIndexForUnderline = num2;
		string pTxtUnsigned = (paragraph.Range.Text ?? "").ToUpperInvariant();
		if (pTxtUnsigned.Contains("DANG CONG SAN") || pTxtUnsigned.Contains("DOAN TNCS") || pTxtUnsigned.Contains("DOAN THANH NIEN"))
		{
			lineIndexForUnderline = 1;
		}
		shape.Top = (float)(lineIndexForUnderline - 1) * num5 + num + 2.5f;
		return shape;
	}

	public static float MeasureTextWidthPt(string text, Microsoft.Office.Interop.Word.Font font)
	{
		if (string.IsNullOrEmpty(text))
		{
			return 0f;
		}
		try
		{
			float emSize = 13f;
			try
			{
				if (font != null && font.Size > 0f)
				{
					emSize = font.Size;
				}
			}
			catch
			{
			}
			string familyName = "Times New Roman";
			try
			{
				if (font != null && !string.IsNullOrEmpty(font.Name))
				{
					familyName = font.Name;
				}
			}
			catch
			{
			}
			FontStyle fontStyle = FontStyle.Regular;
			try
			{
				if (font != null && font.Bold != 0)
				{
					fontStyle |= FontStyle.Bold;
				}
			}
			catch
			{
			}
			try
			{
				if (font != null && font.Italic != 0)
				{
					fontStyle |= FontStyle.Italic;
				}
			}
			catch
			{
			}
			using Bitmap image = new Bitmap(1, 1);
			using Graphics graphics = Graphics.FromImage(image);
			using System.Drawing.Font font2 = new System.Drawing.Font(familyName, emSize, fontStyle, GraphicsUnit.Point);
			StringFormat stringFormat = new StringFormat(StringFormat.GenericTypographic);
			return graphics.MeasureString(text, font2, PointF.Empty, stringFormat).Width * 72f / graphics.DpiX;
		}
		catch
		{
			return (float)text.Length * 5.9f;
		}
	}
}
