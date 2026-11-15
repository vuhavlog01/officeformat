using System;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Formatters;

public static class HeaderFooterFormatter
{
	private static bool HasUserContent(HeaderFooter area)
	{
		if (area == null || !area.Exists)
		{
			return false;
		}
		Range range = area.Range;
		if (range.InlineShapes.Count > 0 || area.Shapes.Count > 0 || range.Tables.Count > 0)
		{
			return true;
		}
		string text = range.Text ?? "";
		foreach (Field field in range.Fields)
		{
			if (field.Type != WdFieldType.wdFieldPage)
			{
				return true;
			}
			string text2 = field.Result.Text;
			if (!string.IsNullOrEmpty(text2))
			{
				text = text.Replace(text2, "");
			}
		}
		return !string.IsNullOrWhiteSpace(text);
	}

	public static void ApplyPageNumbers(Document doc, PageNumberPreset preset, bool preserveUserContent = false)
	{
		if (doc == null || preset == null)
		{
			return;
		}
		Sections comObj = null;
		try
		{
			comObj = doc.Sections;
			int count = comObj.Count;
			for (int i = 1; i <= count; i++)
			{
				OperationProgress.Step();
				Section comObj2 = null;
				HeadersFooters comObj3 = null;
				HeaderFooter comObj4 = null;
				HeadersFooters comObj5 = null;
				HeaderFooter comObj6 = null;
				try
				{
					comObj2 = comObj[i];
					if (!preserveUserContent)
					{
						goto IL_00de;
					}
					bool flag = false;
					foreach (HeaderFooter header in comObj2.Headers)
					{
						flag |= HasUserContent(header);
					}
					foreach (HeaderFooter footer in comObj2.Footers)
					{
						flag |= HasUserContent(footer);
					}
					if (!flag)
					{
						goto IL_00de;
					}
					goto end_IL_0032;
					IL_00de:
					if (preset.HideFirstPage)
					{
						comObj2.PageSetup.DifferentFirstPageHeaderFooter = -1;
						try
						{
							comObj2.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
						}
						catch
						{
						}
					}
					if (preset.CenterVerticallyInTopMargin || preset.DistanceFromTopPt > 0f)
					{
						try
						{
							comObj2.PageSetup.HeaderDistance = (preset.CenterVerticallyInTopMargin ? Math.Max(0f, (comObj2.PageSetup.TopMargin - preset.SizePt) / 2f) : preset.DistanceFromTopPt);
						}
						catch
						{
						}
					}
					comObj5 = comObj2.Footers;
					comObj6 = comObj5[WdHeaderFooterIndex.wdHeaderFooterPrimary];
					if (comObj6 != null && comObj6.Range != null && comObj6.Range.Fields.Count > 0)
					{
						foreach (Field field in comObj6.Range.Fields)
						{
							if (field.Type == WdFieldType.wdFieldPage || field.Type == WdFieldType.wdFieldNumPages)
							{
								try
								{
									field.Delete();
								}
								catch
								{
								}
							}
						}
					}
					comObj3 = comObj2.Headers;
					comObj4 = comObj3[WdHeaderFooterIndex.wdHeaderFooterPrimary];
					if (comObj4 == null || comObj4.Range == null)
					{
						continue;
					}
					try
					{
						if (i == 1)
						{
							comObj4.PageNumbers.RestartNumberingAtSection = true;
							comObj4.PageNumbers.StartingNumber = 1;
						}
						else
						{
							comObj4.PageNumbers.RestartNumberingAtSection = false;
						}
					}
					catch
					{
					}
					Range comObj7 = comObj4.Range;
					comObj7.Text = "";
					Paragraph comObj8 = comObj7.Paragraphs[1];
					WdParagraphAlignment alignment = ConvertPositionToAlignment(preset.Position);
					comObj8.Format.Alignment = alignment;
					comObj8.Format.SpaceBefore = 0f;
					comObj8.Format.SpaceAfter = 0f;
					comObj8.Format.LeftIndent = 0f;
					comObj8.Format.RightIndent = 0f;
					comObj8.Format.FirstLineIndent = 0f;
					try
					{
						comObj8.Format.TabStops.ClearAll();
					}
					catch
					{
					}
					comObj8.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					if (preset.CenterVerticallyInTopMargin)
					{
						comObj8.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
						comObj8.Format.LineSpacing = preset.SizePt;
					}
					Fields fields = comObj4.Range.Fields;
					Range range = comObj8.Range;
					object Type = WdFieldType.wdFieldPage;
					object Text = Missing.Value;
					object PreserveFormatting = Missing.Value;
					Field comObj9 = fields.Add(range, ref Type, ref Text, ref PreserveFormatting);
					try
					{
						comObj9.Update();
					}
					catch
					{
					}
					comObj7.Font.Name = "Times New Roman";
					comObj7.Font.Size = preset.SizePt;
					comObj7.Font.Bold = (preset.Bold ? (-1) : 0);
					comObj7.Font.Italic = (preset.Italic ? (-1) : 0);
					comObj7.Font.ColorIndex = WdColorIndex.wdBlack;
					foreach (Paragraph paragraph in comObj4.Range.Paragraphs)
					{
						try
						{
							paragraph.Format.Alignment = alignment;
							paragraph.Format.LeftIndent = 0f;
							paragraph.Format.RightIndent = 0f;
							paragraph.Format.FirstLineIndent = 0f;
							paragraph.Format.SpaceBefore = 0f;
							paragraph.Format.SpaceAfter = 0f;
							paragraph.Format.TabStops.ClearAll();
							Range range2;
							object Unit;
							object Count;
							for (int num = 0; paragraph.Range.Characters.Count > 1 && num++ < 20; range2 = paragraph.Range.Characters[1], Unit = Missing.Value, Count = Missing.Value, range2.Delete(ref Unit, ref Count))
							{
								switch (paragraph.Range.Characters[1].Text)
								{
								case " ":
								case "\t":
								case "\u00a0":
									continue;
								}
								break;
							}
						}
						catch
						{
						}
					}
					ComReleaser.Release(ref comObj9);
					ComReleaser.Release(ref comObj8);
					ComReleaser.Release(ref comObj7);
					end_IL_0032:;
				}
				catch
				{
				}
				finally
				{
					ComReleaser.Release(ref comObj6);
					ComReleaser.Release(ref comObj5);
					ComReleaser.Release(ref comObj4);
					ComReleaser.Release(ref comObj3);
					ComReleaser.Release(ref comObj2);
				}
			}
			try
			{
				doc.Fields.Update();
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
	}

	private static WdParagraphAlignment ConvertPositionToAlignment(PageNumberPosition pos)
	{
		switch (pos)
		{
		case PageNumberPosition.TopRight:
		case PageNumberPosition.BottomRight:
			return WdParagraphAlignment.wdAlignParagraphRight;
		default:
			return WdParagraphAlignment.wdAlignParagraphCenter;
		}
	}
}
