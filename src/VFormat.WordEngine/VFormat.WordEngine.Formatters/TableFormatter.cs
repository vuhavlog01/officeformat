using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Models;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Formatters;

public static class TableFormatter
{
	private static readonly Regex ShortColumnHeaderPattern = new Regex("^(stt|số\\s*tt|mã|mã\\s*số|số|số\\s*hiệu|ký\\s*hiệu|đvt|đơn\\s*vị\\s*tính|ngày|thời\\s*gian|năm|quý|bước|kỳ|mức|tỷ\\s*lệ|ghi\\s*chú|no\\.?|code|id)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly WdBorderType[] AllBorderTypes = new WdBorderType[6]
	{
		WdBorderType.wdBorderTop,
		WdBorderType.wdBorderLeft,
		WdBorderType.wdBorderBottom,
		WdBorderType.wdBorderRight,
		WdBorderType.wdBorderHorizontal,
		WdBorderType.wdBorderVertical
	};

	public static void FormatTables(Document doc, TablePreset preset)
	{
		if (doc == null || preset == null)
		{
			return;
		}
		Tables comObj = null;
		try
		{
			comObj = doc.Tables;
			int count = comObj.Count;
			for (int i = 1; i <= count; i++)
			{
				OperationProgress.Step($"Đang chuẩn hóa bảng {i}/{count}...");
				Table comObj2 = null;
				try
				{
					comObj2 = comObj[i];
					FormatSingleTable(comObj2, preset);
				}
				catch (OperationCanceledException)
				{
					throw;
				}
				catch
				{
				}
				finally
				{
					ComReleaser.Release(ref comObj2);
				}
			}
		}
		catch (OperationCanceledException)
		{
			throw;
		}
		catch
		{
		}
		finally
		{
			ComReleaser.Release(ref comObj);
		}
	}

	public static void FormatSingleTable(Table table, TablePreset preset)
	{
		if (table == null || preset == null)
		{
			return;
		}
		try
		{
			if (IsMetadataTable(table))
			{
				return;
			}
			int num = 0;
			int num2 = 0;
			try
			{
				num = table.Rows.Count;
				num2 = table.Columns.Count;
			}
			catch
			{
			}
			if (num2 == 2 && num <= 4 && IsMetadataTable(table))
			{
				return;
			}
			Range comObj = null;
			try
			{
				comObj = table.Range;
				comObj.Font.Name = "Times New Roman";
				comObj.Font.ColorIndex = WdColorIndex.wdBlack;
				try
				{
					comObj.ParagraphFormat.FirstLineIndent = 0f;
					comObj.ParagraphFormat.LeftIndent = 0f;
					comObj.ParagraphFormat.RightIndent = 0f;
				}
				catch
				{
				}
			}
			catch
			{
			}
			finally
			{
				ComReleaser.Release(ref comObj);
			}
			try
			{
				NormalizeBorders(table);
			}
			catch
			{
			}
			HashSet<int> hashSet = new HashSet<int>();
			try
			{
				for (int i = 2; i <= num; i++)
				{
					Row row = table.Rows[i];
					if (row.Cells.Count < 2)
					{
						continue;
					}
					string input = CleanCellText(row.Cells[1].Range.Text);
					string input2 = CleanCellText(row.Cells[2].Range.Text);
					if (Regex.IsMatch(input, "^\\d{1,4}\\.?$") && !Regex.IsMatch(input2, "^(tổng|cộng)", RegexOptions.IgnoreCase))
					{
						int bold = row.Range.Font.Bold;
						if (bold == -1 || bold == 1)
						{
							hashSet.Add(i);
						}
					}
				}
			}
			catch
			{
			}
			try
			{
				table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
			}
			catch
			{
			}
			try
			{
				table.TopPadding = preset.CellPaddingTopPt;
				table.BottomPadding = preset.CellPaddingBottomPt;
				table.LeftPadding = preset.CellPaddingLeftPt;
				table.RightPadding = preset.CellPaddingRightPt;
			}
			catch
			{
			}
			if (preset.CantSplit)
			{
				try
				{
					table.Rows.AllowBreakAcrossPages = 0;
				}
				catch
				{
				}
			}
			if (preset.AutoFit)
			{
				try
				{
					table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
					table.AllowAutoFit = false;
				}
				catch
				{
				}
			}
			HashSet<int> hashSet2 = new HashSet<int>();
			try
			{
				if (num > 0 && num2 > 0)
				{
					for (int j = 1; j <= num2; j++)
					{
						try
						{
							Cell cell = table.Cell(1, j);
							if (cell != null)
							{
								string headerText = CleanCellText(cell.Range.Text);
								if (IsShortColumn(headerText))
								{
									hashSet2.Add(j);
								}
							}
						}
						catch
						{
						}
					}
				}
			}
			catch
			{
			}
			if (preset.RepeatHeader && num > 1)
			{
				try
				{
					table.Rows[1].HeadingFormat = -1;
				}
				catch
				{
				}
			}
			float num3 = ((preset.HeaderFontSizePt > 0f) ? preset.HeaderFontSizePt : 13f);
			float num4 = ((preset.BodyFontSizePt > 0f) ? preset.BodyFontSizePt : 13f);
			foreach (Cell cell2 in table.Range.Cells)
			{
				try
				{
					int num5 = 1;
					int item = 1;
					try
					{
						num5 = cell2.RowIndex;
					}
					catch
					{
					}
					try
					{
						item = cell2.ColumnIndex;
					}
					catch
					{
					}
					bool flag = num5 == 1;
					if (!flag && hashSet.Contains(num5))
					{
						try
						{
							cell2.Range.Font.Bold = 0;
						}
						catch
						{
						}
					}
					try
					{
						cell2.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
					}
					catch
					{
					}
					try
					{
						cell2.Range.Font.Size = (flag ? num3 : num4);
						if (flag)
						{
							cell2.Range.Font.Bold = (preset.HeaderBold ? (-1) : 0);
							cell2.Range.Font.Italic = 0;
						}
					}
					catch
					{
					}
					string text = CleanCellText(cell2.Range.Text);
					bool flag2 = flag || hashSet2.Contains(item) || IsShortDataContent(text);
					foreach (Paragraph paragraph in cell2.Range.Paragraphs)
					{
						try
						{
							StripLeadingTabs(paragraph);
							ParagraphFormat paragraphFormat = paragraph.Range.ParagraphFormat;
							paragraphFormat.FirstLineIndent = 0f;
							paragraphFormat.LeftIndent = 0f;
							paragraphFormat.RightIndent = 0f;
							paragraphFormat.SpaceBefore = (flag ? 2f : 1f);
							paragraphFormat.SpaceAfter = (flag ? 2f : 1.5f);
							paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
							if (flag2)
							{
								paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
							}
							else if (paragraphFormat.Alignment != WdParagraphAlignment.wdAlignParagraphRight || !IsNumericLike(text))
							{
								paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
							}
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
		}
		catch
		{
		}
	}

	private static void StripLeadingTabs(Paragraph p)
	{
		try
		{
			string text = p.Range.Text;
			if (!string.IsNullOrEmpty(text))
			{
				int i;
				for (i = 0; i < text.Length && (text[i] == '\t' || (i < 4 && text[i] == ' ')); i++)
				{
				}
				if (i > 0)
				{
					Range comObj = p.Range.Duplicate;
					comObj.End = comObj.Start + i;
					Range range = comObj;
					object Unit = Missing.Value;
					object Count = Missing.Value;
					range.Delete(ref Unit, ref Count);
					ComReleaser.Release(ref comObj);
				}
			}
		}
		catch
		{
		}
	}

	private static void NormalizeBorders(Table table)
	{
		bool flag = false;
		WdBorderType[] allBorderTypes = AllBorderTypes;
		foreach (WdBorderType index in allBorderTypes)
		{
			try
			{
				if (table.Borders[index].LineStyle != WdLineStyle.wdLineStyleNone)
				{
					flag = true;
					break;
				}
			}
			catch
			{
			}
		}
		if (!flag)
		{
			return;
		}
		Borders[] array = new Borders[2];
		try
		{
			array[0] = table.Borders;
		}
		catch
		{
		}
		try
		{
			array[1] = table.Range.Cells.Borders;
		}
		catch
		{
		}
		Borders[] array2 = array;
		foreach (Borders borders in array2)
		{
			if (borders == null)
			{
				continue;
			}
			WdBorderType[] allBorderTypes2 = AllBorderTypes;
			foreach (WdBorderType index2 in allBorderTypes2)
			{
				try
				{
					Border border = borders[index2];
					border.LineStyle = WdLineStyle.wdLineStyleSingle;
					border.LineWidth = WdLineWidth.wdLineWidth050pt;
					border.Color = WdColor.wdColorAutomatic;
				}
				catch
				{
				}
			}
		}
	}

	private static bool IsNumericLike(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		return Regex.IsMatch(text.Trim(), "^[+\\-]?\\d[\\d.,\\s]*(%|đ|đồng|vnđ|vnd)?$", RegexOptions.IgnoreCase);
	}

	private static bool IsShortColumn(string headerText)
	{
		if (string.IsNullOrEmpty(headerText))
		{
			return false;
		}
		string text = headerText.Trim().ToLowerInvariant();
		if (ShortColumnHeaderPattern.IsMatch(text))
		{
			return true;
		}
		if (text.StartsWith("stt") || text.StartsWith("mã") || text == "số" || text == "đvt")
		{
			return true;
		}
		return false;
	}

	private static bool IsShortDataContent(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string input = text.Trim();
		if (Regex.IsMatch(input, "^\\d+([.,]\\d+)?%?$"))
		{
			return true;
		}
		if (Regex.IsMatch(input, "^[A-Z0-9_.-]{1,15}$", RegexOptions.IgnoreCase))
		{
			return true;
		}
		return false;
	}

	private static string CleanCellText(string rawText)
	{
		if (string.IsNullOrEmpty(rawText))
		{
			return string.Empty;
		}
		return rawText.Replace("\r", "").Replace("\n", "").Replace("\a", "")
			.Replace("\v", "")
			.Trim();
	}

	private static bool IsMetadataTable(Table table)
	{
		try
		{
			string text = table.Range.Text;
			if (string.IsNullOrEmpty(text))
			{
				return false;
			}
			if (RegexPatterns.NationalHeader.IsMatch(text) || RegexPatterns.Motto.IsMatch(text) || RegexPatterns.PartyHeader.IsMatch(text) || RegexPatterns.YouthUnionHeader.IsMatch(text) || RegexPatterns.RecipientHeader.IsMatch(text) || Regex.IsMatch(text, "(^|[\\r\\n\\v\\a])\\s*Nơi\\s+nhận\\s*:?", RegexOptions.IgnoreCase))
			{
				return true;
			}
		}
		catch
		{
		}
		return false;
	}
}
