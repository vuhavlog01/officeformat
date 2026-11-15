using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;
using VFormat.Core.Serialization;
using VFormat.WordEngine.Formatters;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine;

public static class WordToolEngine
{
	public static void ApplyPageSetupA4(Document doc, RulePackType rulePack)
	{
		ApplyPageSetupA4(doc, rulePack, PrintLayoutMode.SingleSided);
	}

	public static void ApplyPageSetupA4(Document doc, RulePackType rulePack, PrintLayoutMode printLayoutMode)
	{
		ApplyPageSetupA4(doc, rulePack, printLayoutMode, null);
	}

	public static void ApplyPageSetupA4(Document doc, RulePackType rulePack, PrintLayoutMode printLayoutMode, UserPersonalization personalization)
	{
		if (doc == null)
		{
			return;
		}
		float mm = 20f;
		float mm2 = 20f;
		float mm3 = 30f;
		float mm4 = 15f;
		if (personalization == null)
		{
			try
			{
				JsonProfileRepository jsonProfileRepository = new JsonProfileRepository();
				personalization = jsonProfileRepository.LoadPersonalization();
			}
			catch
			{
			}
		}
		if (personalization != null)
		{
			MarginProfile activeMargins = personalization.GetActiveMargins(rulePack);
			if (activeMargins != null)
			{
				if (activeMargins.MarginTopMm > 0f)
				{
					mm = activeMargins.MarginTopMm;
				}
				if (activeMargins.MarginBottomMm > 0f)
				{
					mm2 = activeMargins.MarginBottomMm;
				}
				if (activeMargins.MarginLeftMm > 0f)
				{
					mm3 = activeMargins.MarginLeftMm;
				}
				if (activeMargins.MarginRightMm > 0f)
				{
					mm4 = activeMargins.MarginRightMm;
				}
			}
		}
		bool flag = printLayoutMode == PrintLayoutMode.TwoSidedMirror;
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Căn Lề Trang A4"))
			{
				foreach (Section section in doc.Sections)
				{
					try
					{
						PageSetup comObj = section.PageSetup;
						comObj.PaperSize = WdPaperSize.wdPaperA4;
						comObj.PageWidth = UnitConverter.MmToPt(210f);
						comObj.PageHeight = UnitConverter.MmToPt(297f);
						comObj.Orientation = WdOrientation.wdOrientPortrait;
						comObj.TopMargin = UnitConverter.MmToPt(mm);
						comObj.BottomMargin = UnitConverter.MmToPt(mm2);
						comObj.LeftMargin = UnitConverter.MmToPt(mm3);
						comObj.RightMargin = UnitConverter.MmToPt(mm4);
						comObj.HeaderDistance = UnitConverter.MmToPt(10f);
						comObj.FooterDistance = UnitConverter.MmToPt(10f);
						comObj.MirrorMargins = (flag ? (-1) : 0);
						comObj.DifferentFirstPageHeaderFooter = -1;
						ComReleaser.Release(ref comObj);
					}
					catch
					{
					}
				}
			}
		}
	}

	public static void ApplyCustomPageMargins(Document doc, float topMm, float botMm, float leftMm, float rightMm, float hdrDistMm, float ftrDistMm, bool isLandscape, bool forCurrentSectionOnly)
	{
		ApplyCustomPageMargins(doc, topMm, botMm, leftMm, rightMm, hdrDistMm, ftrDistMm, isLandscape, forCurrentSectionOnly, mirrorMargins: false);
	}

	public static void ApplyCustomPageMargins(Document doc, float topMm, float botMm, float leftMm, float rightMm, float hdrDistMm, float ftrDistMm, bool isLandscape, bool forCurrentSectionOnly, bool mirrorMargins)
	{
		if (doc == null)
		{
			return;
		}
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Cài Đặt Căn Lề Trang A4"))
			{
				Sections sections = ((forCurrentSectionOnly && doc.Application.Selection != null && doc.Application.Selection.Sections != null) ? doc.Application.Selection.Sections : doc.Sections);
				foreach (Section item in sections)
				{
					try
					{
						PageSetup comObj = item.PageSetup;
						comObj.PaperSize = WdPaperSize.wdPaperA4;
						comObj.Orientation = (isLandscape ? WdOrientation.wdOrientLandscape : WdOrientation.wdOrientPortrait);
						if (isLandscape)
						{
							comObj.PageWidth = UnitConverter.MmToPt(297f);
							comObj.PageHeight = UnitConverter.MmToPt(210f);
						}
						else
						{
							comObj.PageWidth = UnitConverter.MmToPt(210f);
							comObj.PageHeight = UnitConverter.MmToPt(297f);
						}
						comObj.TopMargin = UnitConverter.MmToPt(topMm);
						comObj.BottomMargin = UnitConverter.MmToPt(botMm);
						comObj.LeftMargin = UnitConverter.MmToPt(leftMm);
						comObj.RightMargin = UnitConverter.MmToPt(rightMm);
						comObj.HeaderDistance = UnitConverter.MmToPt(hdrDistMm);
						comObj.FooterDistance = UnitConverter.MmToPt(ftrDistMm);
						comObj.MirrorMargins = (mirrorMargins ? (-1) : 0);
						ComReleaser.Release(ref comObj);
					}
					catch
					{
					}
				}
			}
		}
	}

	public static void ApplyCustomPageNumbering(Document doc, int fromPage, int startNum, bool hideFirst, bool isHeader, float fontSize = 14f)
	{
		if (doc == null)
		{
			return;
		}
		int num = 1;
		try
		{
			object IncludeFootnotesAndEndnotes = Missing.Value;
			num = doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref IncludeFootnotesAndEndnotes);
			if (num < 1)
			{
				num = 1;
			}
		}
		catch
		{
		}
		if (fromPage < 1)
		{
			fromPage = 1;
		}
		if (startNum < 1)
		{
			startNum = 1;
		}
		if (fromPage > num)
		{
			throw new InvalidOperationException($"Trang bắt đầu vượt quá tổng số trang của tài liệu (hiện có {num} trang)!");
		}
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Đánh Số Trang"))
			{
				if (fromPage <= 1)
				{
					foreach (Section section4 in doc.Sections)
					{
						try
						{
							section4.PageSetup.DifferentFirstPageHeaderFooter = (hideFirst ? (-1) : 0);
							HeaderFooter headerFooter;
							if (isHeader)
							{
								section4.PageSetup.HeaderDistance = UnitConverter.MmToPt(10f);
								headerFooter = section4.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
								section4.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
								section4.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
							}
							else
							{
								section4.PageSetup.FooterDistance = UnitConverter.MmToPt(10f);
								headerFooter = section4.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
								section4.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
								section4.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
							}
							headerFooter.Range.Text = "";
							headerFooter.Range.ParagraphFormat.TabStops.ClearAll();
							headerFooter.Range.ParagraphFormat.LeftIndent = 0f;
							headerFooter.Range.ParagraphFormat.RightIndent = 0f;
							headerFooter.Range.ParagraphFormat.FirstLineIndent = 0f;
							headerFooter.Range.Font.Name = "Times New Roman";
							headerFooter.Range.Font.Size = fontSize;
							headerFooter.Range.Font.Bold = 0;
							headerFooter.Range.Font.Italic = 0;
							headerFooter.Range.Font.ColorIndex = WdColorIndex.wdBlack;
							headerFooter.PageNumbers.RestartNumberingAtSection = true;
							headerFooter.PageNumbers.StartingNumber = startNum;
							object Type = WdFieldType.wdFieldPage;
							Fields fields = headerFooter.Range.Fields;
							Range range = headerFooter.Range;
							object Text = Missing.Value;
							object PreserveFormatting = Missing.Value;
							fields.Add(range, ref Type, ref Text, ref PreserveFormatting);
							headerFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
							if (hideFirst)
							{
								if (isHeader)
								{
									section4.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
								}
								else
								{
									section4.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
								}
							}
						}
						catch
						{
						}
					}
					return;
				}
				object What = WdGoToItem.wdGoToPage;
				object Which = WdGoToDirection.wdGoToFirst;
				object Count = fromPage;
				object Name = Missing.Value;
				Range range2 = doc.GoTo(ref What, ref Which, ref Count, ref Name);
				int index = range2.Sections[1].Index;
				bool flag = false;
				if (index > 1 && range2.Start == range2.Sections[1].Range.Start)
				{
					flag = true;
				}
				if (!flag)
				{
					try
					{
						object Start = range2.Start - 1;
						object End = range2.Start;
						Range range3 = doc.Range(ref Start, ref End);
						if (range3.Text == "\f")
						{
							object Unit = Missing.Value;
							object Count2 = Missing.Value;
							range3.Delete(ref Unit, ref Count2);
						}
					}
					catch
					{
					}
					object Type2 = WdBreakType.wdSectionBreakNextPage;
					range2.InsertBreak(ref Type2);
					object Name2 = Missing.Value;
					range2 = doc.GoTo(ref What, ref Which, ref Count, ref Name2);
					index = range2.Sections[1].Index;
				}
				Section section2 = doc.Sections[index];
				try
				{
					section2.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
					section2.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
					section2.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
					section2.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
					section2.PageSetup.DifferentFirstPageHeaderFooter = 0;
				}
				catch
				{
				}
				for (int i = 1; i < index; i++)
				{
					try
					{
						doc.Sections[i].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
						doc.Sections[i].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
						doc.Sections[i].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
						doc.Sections[i].Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
					}
					catch
					{
					}
				}
				HeaderFooter headerFooter2;
				if (isHeader)
				{
					section2.PageSetup.HeaderDistance = UnitConverter.MmToPt(10f);
					headerFooter2 = section2.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
					section2.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
				}
				else
				{
					section2.PageSetup.FooterDistance = UnitConverter.MmToPt(10f);
					headerFooter2 = section2.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
					section2.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
				}
				headerFooter2.PageNumbers.RestartNumberingAtSection = true;
				headerFooter2.PageNumbers.StartingNumber = startNum;
				headerFooter2.Range.Text = "";
				headerFooter2.Range.ParagraphFormat.TabStops.ClearAll();
				headerFooter2.Range.ParagraphFormat.LeftIndent = 0f;
				headerFooter2.Range.ParagraphFormat.RightIndent = 0f;
				headerFooter2.Range.ParagraphFormat.FirstLineIndent = 0f;
				headerFooter2.Range.Font.Name = "Times New Roman";
				headerFooter2.Range.Font.Size = fontSize;
				headerFooter2.Range.Font.Bold = 0;
				headerFooter2.Range.Font.Italic = 0;
				headerFooter2.Range.Font.ColorIndex = WdColorIndex.wdBlack;
				object Type3 = WdFieldType.wdFieldPage;
				Fields fields2 = headerFooter2.Range.Fields;
				Range range4 = headerFooter2.Range;
				object Text2 = Missing.Value;
				object PreserveFormatting2 = Missing.Value;
				fields2.Add(range4, ref Type3, ref Text2, ref PreserveFormatting2);
				headerFooter2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				for (int j = index + 1; j <= doc.Sections.Count; j++)
				{
					try
					{
						Section section3 = doc.Sections[j];
						HeaderFooter headerFooter3;
						if (isHeader)
						{
							headerFooter3 = section3.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
							section3.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
						}
						else
						{
							headerFooter3 = section3.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
							section3.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
						}
						headerFooter3.PageNumbers.RestartNumberingAtSection = false;
					}
					catch
					{
					}
				}
			}
		}
	}

	public static void RemoveAllPageNumbers(Document doc)
	{
		if (doc == null)
		{
			return;
		}
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Xóa Số Trang"))
			{
				foreach (Section section in doc.Sections)
				{
					try
					{
						section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
						section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
						section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "";
						section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = "";
					}
					catch
					{
					}
				}
			}
		}
	}

	public static void InsertLandscapePage(Application app)
	{
		InsertIndependentPage(app, WdOrientation.wdOrientLandscape);
	}

	public static void InsertPortraitPage(Application app)
	{
		InsertIndependentPage(app, WdOrientation.wdOrientPortrait);
	}

	private static void InsertIndependentPage(Application app, WdOrientation orientation)
	{
		if (app == null || app.Documents.Count == 0 || app.Selection == null)
		{
			return;
		}
		Document activeDocument = app.ActiveDocument;
		Selection selection = app.Selection;
		using (new ScreenFreezeScope(app))
		{
			using (new SafeUndoScope(app, "OfficeFormat Chèn Trang"))
			{
				object Direction = WdCollapseDirection.wdCollapseStart;
				selection.Collapse(ref Direction);
				int start = selection.Start;
				PageSetup comObj = selection.PageSetup;
				float topMargin = comObj.TopMargin;
				float bottomMargin = comObj.BottomMargin;
				float leftMargin = comObj.LeftMargin;
				float rightMargin = comObj.RightMargin;
				float gutter = comObj.Gutter;
				int mirrorMargins = comObj.MirrorMargins;
				ComReleaser.Release(ref comObj);
				bool flag = start < activeDocument.Content.End - 1;
				object Start = start;
				object End = start;
				Range comObj2 = activeDocument.Range(ref Start, ref End);
				try
				{
					Range range = comObj2;
					object Type = WdBreakType.wdSectionBreakNextPage;
					range.InsertBreak(ref Type);
					int num = start + 1;
					if (flag)
					{
						comObj2.SetRange(num, num);
						Range range2 = comObj2;
						object Type2 = WdBreakType.wdSectionBreakNextPage;
						range2.InsertBreak(ref Type2);
					}
					object Start2 = num;
					object End2 = num;
					Range comObj3 = activeDocument.Range(ref Start2, ref End2);
					try
					{
						comObj3.Sections[1].PageSetup.Orientation = orientation;
						comObj3.Sections[1].PageSetup.PageWidth = UnitConverter.MmToPt((orientation == WdOrientation.wdOrientLandscape) ? 297f : 210f);
						comObj3.Sections[1].PageSetup.PageHeight = UnitConverter.MmToPt((orientation == WdOrientation.wdOrientLandscape) ? 210f : 297f);
						comObj3.Sections[1].PageSetup.TopMargin = topMargin;
						comObj3.Sections[1].PageSetup.BottomMargin = bottomMargin;
						comObj3.Sections[1].PageSetup.LeftMargin = leftMargin;
						comObj3.Sections[1].PageSetup.RightMargin = rightMargin;
						comObj3.Sections[1].PageSetup.Gutter = gutter;
						comObj3.Sections[1].PageSetup.MirrorMargins = mirrorMargins;
						selection.SetRange(num, num);
					}
					finally
					{
						ComReleaser.Release(ref comObj3);
					}
				}
				finally
				{
					ComReleaser.Release(ref comObj2);
				}
			}
		}
	}

	public static int DeleteBlankPages(Document doc)
	{
		if (doc == null)
		{
			return 0;
		}
		int num = 0;
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat XÃ³a Trang Tráº¯ng Thá»«a"))
			{
				int initialPages = 0;
				try
				{
					initialPages = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
				}
				catch
				{
				}

				// Phase 1: XÃ³a cÃ¡c Section trá»‘ng thá»«a á»Ÿ cuá»‘i tÃ i liá»‡u
				for (int sIdx = doc.Sections.Count; sIdx > 1; sIdx--)
				{
					Section sec = null;
					Range secRange = null;
					Section prevSec = null;
					try
					{
						sec = doc.Sections[sIdx];
						secRange = sec.Range;
						if (IsEmptySection(doc, secRange))
						{
							if (sIdx == doc.Sections.Count)
							{
								prevSec = doc.Sections[sIdx - 1];
								CopySectionPageSetup(prevSec, sec);
								WdHeaderFooterIndex[] array = new WdHeaderFooterIndex[3]
								{
									WdHeaderFooterIndex.wdHeaderFooterPrimary,
									WdHeaderFooterIndex.wdHeaderFooterFirstPage,
									WdHeaderFooterIndex.wdHeaderFooterEvenPages
								};
								foreach (WdHeaderFooterIndex index in array)
								{
									try
									{
										sec.Headers[index].LinkToPrevious = true;
										sec.Footers[index].LinkToPrevious = true;
									}
									catch
									{
									}
								}
								secRange.SetRange(Math.Max(0, secRange.Start - 1), doc.Content.End - 1);
							}
							object Unit = Missing.Value;
							object Count = Missing.Value;
							secRange.Delete(ref Unit, ref Count);
							num++;
						}
					}
					catch
					{
					}
					finally
					{
						ComReleaser.Release(ref prevSec);
						ComReleaser.Release(ref secRange);
						ComReleaser.Release(ref sec);
					}
				}

				// Phase 2: XÃ³a ngáº¯t trang liÃªn tiáº¿p (^m^m -> ^m)
				try
				{
					Range rngFind = doc.Content;
					Find find = rngFind.Find;
					find.ClearFormatting();
					find.Replacement.ClearFormatting();
					find.Text = "^m^m";
					find.Replacement.Text = "^m";
					object findText = "^m^m";
					object replText = "^m";
					object replaceAll = WdReplace.wdReplaceAll;
					object missing = Missing.Value;
					while (find.Execute(ref findText, ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing, ref replText, ref replaceAll,
						ref missing, ref missing, ref missing, ref missing))
					{
						num++;
					}
				}
				catch
				{
				}

				// Phase 3: XÃ³a ngáº¯t trang thá»«a Ä‘Æ¡n Ä‘á»™c á»Ÿ cuá»‘i tÃ i liá»‡u (trailing Form Feed \f)
				try
				{
					string contentText = doc.Content.Text ?? "";
					int lastF = contentText.LastIndexOf('\f');
					if (lastF >= 0)
					{
						string trailing = contentText.Substring(lastF + 1);
						string trailingClean = (trailing ?? "").Replace("\r", "").Replace("\n", "").Replace("\f", "")
							.Replace("\v", "").Replace("\t", "").Replace(" ", "").Replace("\u00a0", "");
						if (string.IsNullOrEmpty(trailingClean))
						{
							object fStart = lastF;
							object fEnd = lastF + 1;
							Range fRange = doc.Range(ref fStart, ref fEnd);
							try
							{
								if (IsEmptySection(doc, fRange))
								{
									object Unit = Missing.Value;
									object Count = Missing.Value;
									fRange.Delete(ref Unit, ref Count);
									num++;
								}
							}
							finally
							{
								ComReleaser.Release(ref fRange);
							}
						}
					}
				}
				catch
				{
				}

				// Phase 4: XÃ³a toÃ n bá»™ Ä‘oáº¡n vÄƒn trá»‘ng thá»«a á»Ÿ cuá»‘i tÃ i liá»‡u (tá»« ná»™i dung thá»±c táº¿ cuá»‘i cÃ¹ng Ä‘áº¿n cuá»‘i tÃ i liá»‡u)
				try
				{
					int lastContentEnd = 0;
					int tableCount = doc.Tables.Count;
					if (tableCount > 0)
					{
						Table lastTable = doc.Tables[tableCount];
						try
						{
							lastContentEnd = lastTable.Range.End;
						}
						finally
						{
							Marshal.ReleaseComObject(lastTable);
						}
					}

					for (int pIdx = doc.Paragraphs.Count; pIdx >= 1; pIdx--)
					{
						Paragraph p = null;
						Range pr = null;
						try
						{
							p = doc.Paragraphs[pIdx];
							pr = p.Range;
							if (pr.Tables.Count > 0)
							{
								if (pr.End > lastContentEnd)
								{
									lastContentEnd = pr.End;
								}
								break;
							}
							string pt = (pr.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\f", "")
								.Replace("\v", "").Replace("\t", "").Replace(" ", "").Replace("\u00a0", "");
							if (!string.IsNullOrEmpty(pt) || pr.InlineShapes.Count > 0 || pr.Fields.Count > 0 || pr.OMaths.Count > 0 || pr.ContentControls.Count > 0)
							{
								if (pr.End > lastContentEnd)
								{
									lastContentEnd = pr.End;
								}
								break;
							}
						}
						catch
						{
						}
						finally
						{
							ComReleaser.Release(ref pr);
							ComReleaser.Release(ref p);
						}
					}

					if (lastContentEnd > 0 && lastContentEnd < doc.Content.End - 1)
					{
						object cStart = lastContentEnd;
						object cEnd = doc.Content.End - 1;
						Range delRange = doc.Range(ref cStart, ref cEnd);
						try
						{
							object Unit = Missing.Value;
							object Count = Missing.Value;
							delRange.Delete(ref Unit, ref Count);
							num++;
						}
						finally
						{
							ComReleaser.Release(ref delRange);
						}
					}
				}
				catch
				{
				}

				// Phase 5: Xá»­ lÃ½ kÃ½ tá»± ngáº¯t Ä‘oáº¡n báº¯t buá»™c sau Báº£ng cuá»‘i tÃ i liá»‡u hoáº·c sau vÄƒn báº£n cuá»‘i tÃ i liá»‡u
				// (TrÆ°á»ng há»£p Word: Báº£ng cuá»‘i trang Ä‘áº©y Ä‘oáº¡n Â¶ báº¯t buá»™c sang trang má»›i táº¡o trang tráº¯ng thá»«a)
				try
				{
					int tableCount = doc.Tables.Count;
					bool endsWithTable = false;
					Table lastTable = null;
					if (tableCount > 0)
					{
						lastTable = doc.Tables[tableCount];
						if (lastTable.Range.End >= doc.Content.End - 2)
						{
							endsWithTable = true;
						}
					}

					int pCount = doc.Paragraphs.Count;
					if (pCount > 0)
					{
						Paragraph lastP = doc.Paragraphs[pCount];
						try
						{
							string lastPText = (lastP.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\f", "")
								.Replace("\v", "").Replace("\t", "").Replace(" ", "").Replace("\u00a0", "");
							if (string.IsNullOrEmpty(lastPText) && (endsWithTable || pCount > 1))
							{
								// Thu nhá» kÃ­ch thÆ°á»›c tá»‘i Ä‘a: 1pt, áº©n, dÃ£n dÃ²ng chÃ­nh xÃ¡c 1pt, loáº¡i bá» má»i khoáº£ng cÃ¡ch Ä‘á»‡m
								lastP.Range.Font.Size = 1f;
								try { lastP.Range.Font.Hidden = -1; } catch { }
								lastP.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
								lastP.Format.LineSpacing = 1f;
								lastP.Format.SpaceBefore = 0f;
								lastP.Format.SpaceAfter = 0f;
								lastP.Format.FirstLineIndent = 0f;
								lastP.Format.LeftIndent = 0f;
								lastP.Format.RightIndent = 0f;

								if (endsWithTable && lastTable != null)
								{
									try
									{
										lastTable.BottomPadding = 0f;
										lastTable.TopPadding = 0f;
									}
									catch { }
								}
							}
						}
						finally
						{
							ComReleaser.Release(ref lastP);
						}
					}

					if (lastTable != null)
					{
						Marshal.ReleaseComObject(lastTable);
					}
				}
				catch
				{
				}

				// Phase 6: Náº¿u con trá» chuá»™t/lá»±a chá»n Ä‘ang Ä‘á»©ng trÃªn má»™t trang hoÃ n toÃ n tráº¯ng, xÃ³a trang Ä‘Ã³
				try
				{
					if (doc.Bookmarks.Exists("\\Page"))
					{
						Range pageRange = doc.Bookmarks["\\Page"].Range;
						try
						{
							if (IsEmptySection(doc, pageRange) && pageRange.End < doc.Content.End - 1)
							{
								object Unit = Missing.Value;
								object Count = Missing.Value;
								pageRange.Delete(ref Unit, ref Count);
								num++;
							}
						}
						finally
						{
							ComReleaser.Release(ref pageRange);
						}
					}
				}
				catch
				{
				}

				// Phase 7: TÃ­nh toÃ¡n láº¡i trang vÃ  Ä‘á»“ng bá»™ sá»‘ trang Ä‘Ã£ xÃ³a
				try
				{
					doc.Repaginate();
					int finalPages = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
					if (initialPages > 0 && finalPages < initialPages)
					{
						int diff = initialPages - finalPages;
						if (diff > num)
						{
							num = diff;
						}
					}
				}
				catch
				{
				}
			}
		}
		return num;
	}
	private static bool IsEmptySection(Document doc, Range range)
	{
		if (range.Tables.Count != 0 || range.InlineShapes.Count != 0 || range.Fields.Count != 0 || range.ContentControls.Count != 0 || range.OMaths.Count != 0)
		{
			return false;
		}
		foreach (Shape shape in doc.Shapes)
		{
			try
			{
				if (shape.Anchor.Start >= range.Start && shape.Anchor.Start < range.End)
				{
					return false;
				}
			}
			finally
			{
				Marshal.ReleaseComObject(shape);
			}
		}
		return string.IsNullOrWhiteSpace((range.Text ?? "").Replace("\f", "").Replace("\v", ""));
	}

	private static void CopySectionPageSetup(Section source, Section target)
	{
		PageSetup comObj = source.PageSetup;
		PageSetup comObj2 = target.PageSetup;
		try
		{
			comObj2.PaperSize = comObj.PaperSize;
			comObj2.Orientation = comObj.Orientation;
			comObj2.PageWidth = comObj.PageWidth;
			comObj2.PageHeight = comObj.PageHeight;
			comObj2.MirrorMargins = comObj.MirrorMargins;
			comObj2.GutterPos = comObj.GutterPos;
			comObj2.Gutter = comObj.Gutter;
			comObj2.TopMargin = comObj.TopMargin;
			comObj2.BottomMargin = comObj.BottomMargin;
			comObj2.LeftMargin = comObj.LeftMargin;
			comObj2.RightMargin = comObj.RightMargin;
			comObj2.HeaderDistance = comObj.HeaderDistance;
			comObj2.FooterDistance = comObj.FooterDistance;
			comObj2.SectionStart = comObj.SectionStart;
			comObj2.VerticalAlignment = comObj.VerticalAlignment;
			comObj2.DifferentFirstPageHeaderFooter = comObj.DifferentFirstPageHeaderFooter;
			comObj2.OddAndEvenPagesHeaderFooter = comObj.OddAndEvenPagesHeaderFooter;
			comObj2.SuppressEndnotes = comObj.SuppressEndnotes;
			comObj2.TwoPagesOnOne = comObj.TwoPagesOnOne;
			comObj2.LineNumbering.Active = comObj.LineNumbering.Active;
			if (comObj.LineNumbering.Active != 0)
			{
				comObj2.LineNumbering.StartingNumber = comObj.LineNumbering.StartingNumber;
				comObj2.LineNumbering.CountBy = comObj.LineNumbering.CountBy;
				comObj2.LineNumbering.RestartMode = comObj.LineNumbering.RestartMode;
				comObj2.LineNumbering.DistanceFromText = comObj.LineNumbering.DistanceFromText;
			}
			comObj2.TextColumns.SetCount(comObj.TextColumns.Count);
			int num = 1;
			while (comObj.TextColumns.Count > 1 && num <= comObj.TextColumns.Count)
			{
				comObj2.TextColumns[num].Width = comObj.TextColumns[num].Width;
				comObj2.TextColumns[num].SpaceAfter = comObj.TextColumns[num].SpaceAfter;
				num++;
			}
		}
		finally
		{
			ComReleaser.Release(ref comObj);
			ComReleaser.Release(ref comObj2);
		}
	}

	public static void CleanSpacingErrors(Document doc, Selection sel = null, Action<int, string> onProgress = null)
	{
		if (doc == null)
		{
			return;
		}
		Range rng = ((sel != null && sel.Type == WdSelectionType.wdSelectionNormal && sel.Range.Characters.Count > 2) ? sel.Range : doc.Content);
		onProgress?.Invoke(84, "Đang dọn dẹp khoảng trắng thừa và ký tự rác...");
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Xóa Khoảng Trắng Lỗi"))
			{
				ReplaceText(rng, " ... ", " ___VF_BLANK___ ");
				for (int num = 15; num >= 2; num--)
				{
					string text = new string(' ', num);
					ReplaceText(rng, "ngày" + text + "tháng", "ngày___VF_SP1_" + num + "___tháng");
					ReplaceText(rng, "tháng" + text + "năm", "tháng___VF_SP2_" + num + "___năm");
					ReplaceText(rng, "Số:" + text + "/", "Số:___VF_NUMSP_" + num + "___/");
					ReplaceText(rng, "Số: " + text + "/", "Số: ___VF_NUMSP_" + num + "___/");
					ReplaceText(rng, "Số" + text + "-", "Số___VF_PTYNUM_" + num + "___-");
					ReplaceText(rng, "Số " + text + "-", "Số ___VF_PTYNUM_" + num + "___-");
				}
				ReplaceText(rng, "\u00a0", " ");
				for (int i = 0; i < 5; i++)
				{
					ReplaceText(rng, "  ", " ");
				}
				ReplaceText(rng, ".^t", ". ");
				ReplaceText(rng, ". ^t", ". ");
				ReplaceText(rng, " ,", ",");
				ReplaceText(rng, " .", ".");
				ReplaceText(rng, " ;", ";");
				ReplaceText(rng, " :", ":");
				ReplaceText(rng, " ?", "?");
				ReplaceText(rng, " !", "!");
				ReplaceText(rng, " )", ")");
				ReplaceText(rng, " ]", "]");
				ReplaceText(rng, " }", "}");
				ReplaceText(rng, " ”", "”");
				ReplaceText(rng, " ’", "’");
				ReplaceText(rng, "( ", "(");
				ReplaceText(rng, "[ ", "[");
				ReplaceText(rng, "{ ", "{");
				ReplaceText(rng, "“ ", "“");
				ReplaceText(rng, "‘ ", "‘");
				TrimParagraphBoundarySpaces(doc, rng);
				for (int num2 = 15; num2 >= 2; num2--)
				{
					string replaceText = new string(' ', num2);
					ReplaceText(rng, "___VF_SP1_" + num2 + "___", replaceText);
					ReplaceText(rng, "___VF_SP2_" + num2 + "___", replaceText);
					ReplaceText(rng, "___VF_NUMSP_" + num2 + "___", replaceText);
					ReplaceText(rng, "___VF_PTYNUM_" + num2 + "___", replaceText);
				}
				ReplaceText(rng, " ___VF_BLANK___ ", " ... ");
			}
		}
	}

	private static void TrimParagraphBoundarySpaces(Document doc, Range rng)
	{
		OperationProgress.Step();
		OperationProgress.Step();
		OperationProgress.Step();
		if (doc == null || rng == null)
		{
			return;
		}
		try
		{
			Paragraphs paras = rng.Paragraphs;
			int pCount = paras.Count;
			for (int i = pCount; i >= 1; i--)
			{
				Paragraph p = null;
				Range pRange = null;
				try
				{
					p = paras[i];
					pRange = p.Range;
					string t = pRange.Text ?? "";
					if (t.Length < 2 || pRange.Fields.Count > 0)
					{
						continue;
					}
					int pStart = pRange.Start;
					int contentEndIdx = t.Length;
					while (contentEndIdx > 0 && (t[contentEndIdx - 1] == '\r' || t[contentEndIdx - 1] == '\n' || t[contentEndIdx - 1] == '\a'))
					{
						contentEndIdx--;
					}
					int trimEndIdx = contentEndIdx;
					while (trimEndIdx > 0 && (t[trimEndIdx - 1] == ' ' || t[trimEndIdx - 1] == '\u00a0'))
					{
						trimEndIdx--;
					}
					if (trimEndIdx < contentEndIdx && trimEndIdx > 0)
					{
						object s1 = pStart + trimEndIdx;
						object e1 = pStart + contentEndIdx;
						Range delEnd = doc.Range(ref s1, ref e1);
						if ((delEnd.Text ?? "").Trim(' ', '\u00a0').Length == 0)
						{
							object u = Missing.Value;
							object c = Missing.Value;
							delEnd.Delete(ref u, ref c);
						}
						ComReleaser.Release(ref delEnd);
					}
					int trimStartIdx = 0;
					while (trimStartIdx < trimEndIdx && (t[trimStartIdx] == ' ' || t[trimStartIdx] == '\u00a0'))
					{
						trimStartIdx++;
					}
					if (trimStartIdx > 0 && trimStartIdx < trimEndIdx)
					{
						object s2 = pStart;
						object e2 = pStart + trimStartIdx;
						Range delStart = doc.Range(ref s2, ref e2);
						if ((delStart.Text ?? "").Trim(' ', '\u00a0').Length == 0)
						{
							object u2 = Missing.Value;
							object c2 = Missing.Value;
							delStart.Delete(ref u2, ref c2);
						}
						ComReleaser.Release(ref delStart);
					}
				}
				catch
				{
				}
				finally
				{
					ComReleaser.Release(ref pRange);
					ComReleaser.Release(ref p);
				}
			}
			ComReleaser.Release(ref paras);
		}
		catch
		{
		}
	}

	private static void ReplaceText(Range rng, string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false)
	{
		OperationProgress.Step();
		if (rng == null || string.IsNullOrEmpty(findText))
		{
			return;
		}
		Range searchRange = null;
		Find comObj = null;
		try
		{
			string currentText = rng.Text ?? "";
			if (currentText.Length == 0)
			{
				return;
			}
			string literalProbe = findText.Replace("^t", "\t").Replace("^p", "\r").Replace("^l", "\v");
			StringComparison cmp = (matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
			if (currentText.IndexOf(literalProbe, cmp) < 0)
			{
				return;
			}
			searchRange = rng.Duplicate;
			int rangeStart = rng.Start;
			int rangeEnd = rng.End;
			comObj = searchRange.Find;
			comObj.ClearFormatting();
			comObj.Replacement.ClearFormatting();
			comObj.Text = findText;
			comObj.Replacement.Text = replaceText;
			comObj.Forward = true;
			comObj.Wrap = WdFindWrap.wdFindStop;
			comObj.Format = false;
			comObj.MatchCase = matchCase;
			comObj.MatchWholeWord = matchWholeWord;
			comObj.MatchWildcards = false;
			int delta = (replaceText?.Length ?? 0) - literalProbe.Length;
			int guard = 0;
			while (guard++ < 5000)
			{
				object Replace = WdReplace.wdReplaceOne;
				Find find = comObj;
				object FindText = Missing.Value;
				object MatchCase = Missing.Value;
				object MatchWholeWord = Missing.Value;
				object MatchWildcards = Missing.Value;
				object MatchSoundsLike = Missing.Value;
				object MatchAllWordForms = Missing.Value;
				object Forward = Missing.Value;
				object Wrap = Missing.Value;
				object Format = Missing.Value;
				object ReplaceWith = Missing.Value;
				object MatchKashida = Missing.Value;
				object MatchDiacritics = Missing.Value;
				object MatchAlefHamza = Missing.Value;
				object MatchControl = Missing.Value;
				if (!find.Execute(ref FindText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl))
				{
					break;
				}
				rangeEnd += delta;
				int nextStart = searchRange.End;
				int actualDocEnd = searchRange.Document.Content.End;
				if (rangeEnd > actualDocEnd)
				{
					rangeEnd = actualDocEnd;
				}
				if (nextStart < rangeStart || nextStart >= rangeEnd)
				{
					break;
				}
				searchRange.SetRange(nextStart, rangeEnd);
			}
		}
		catch
		{
		}
		finally
		{
			ComReleaser.Release(ref comObj);
			ComReleaser.Release(ref searchRange);
		}
	}

	public static void RemoveToneSelection(Selection sel)
	{
		if (sel == null || sel.Type != WdSelectionType.wdSelectionNormal || string.IsNullOrEmpty(sel.Text))
		{
			return;
		}
		using (new SafeUndoScope(sel.Application, "OfficeFormat Bỏ Dấu Tiếng Việt"))
		{
			string text = sel.Text;
			sel.Text = RemoveVietnameseDiacritics(text);
		}
	}

	public static void ConvertOldFontsToUnicode(Document doc)
	{
		if (doc == null)
		{
			return;
		}
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Chuyển Mã Unicode"))
			{
				try
				{
					ConvertTcvn3InDocument(doc);
				}
				catch
				{
				}
				try
				{
					doc.Content.Font.Name = "Times New Roman";
					doc.Content.Font.NameAscii = "Times New Roman";
					doc.Content.Font.NameOther = "Times New Roman";
					doc.Content.Font.NameBi = "Times New Roman";
				}
				catch
				{
				}
				foreach (Table table in doc.Tables)
				{
					OperationProgress.Step();
					try
					{
						table.Range.Font.Name = "Times New Roman";
						table.Range.Font.NameAscii = "Times New Roman";
						table.Range.Font.NameOther = "Times New Roman";
					}
					catch
					{
					}
				}
				foreach (Section section in doc.Sections)
				{
					try
					{
						section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = "Times New Roman";
						section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = "Times New Roman";
					}
					catch
					{
					}
				}
			}
		}
	}

	private static void ConvertTcvn3InDocument(Document doc)
	{
		string[,] array = new string[30, 2]
		{
			{ "CÃ©NG HÃŸA XÂ· HÃ©I CHÃ± NGHÃœA VIÃ–T NAM", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" },
			{ "THÂ¤NG B\u00b8O", "THÔNG BÁO" },
			{ "QUYÃ•T Â§ÃžNH", "QUYẾT ĐỊNH" },
			{ "NGHÃž QUYÃ•T", "NGHỊ QUYẾT" },
			{ "KÃ• HOÂ¹CH", "KẾ HOẠCH" },
			{ "B\u00b8O C\u00b8O", "BÁO CÁO" },
			{ "TÃª TRÃ—NH", "TỜ TRÌNH" },
			{ "BIÂ£N BÂ¶N", "BIÊN BẢN" },
			{ "CÂ¤NG VIÃ–C", "CÔNG VIỆC" },
			{ "HÂ¦Ã­NG DÃ‰N", "HƯỚNG DẪN" },
			{ "GIÃŠN VIÂ£N", "GIẢNG VIÊN" },
			{ "Â§Ã©c lÃ‹p - TÃ¹ do - HÂ¹nh phÃ³c", "Độc lập - Tự do - Hạnh phúc" },
			{ "SÃ« GI\u00b8O DÃ´C VÂµ Â§ÂµO TÂ¹O", "SỞ GIÁO DỤC VÀ ĐÀO TẠO" },
			{ "VÂµ Â§ÂµO TÂ¹O", "VÀ ĐÀO TẠO" },
			{ "ngÂµy", "ngày" },
			{ "th\u00b8ng", "tháng" },
			{ "n\u00a8m", "năm" },
			{ "KÃnh gÃ¶i", "Kính gửi" },
			{ "NÂ¬i nhÃ‹n", "Nơi nhận" },
			{ "ThÃ± trÃ«ng", "Thủ trưởng" },
			{ "Gi\u00b8m Â®Ã¨c", "Giám đốc" },
			{ "PhÃ£ Gi\u00b8m Â®Ã¨c", "Phó Giám đốc" },
			{ "ChÃ± tÃžch", "Chủ tịch" },
			{ "PhÃ£ ChÃ± tÃžch", "Phó Chủ tịch" },
			{ "BÃ th", "Bí thư" },
			{ "PhÃ£ BÃ th", "Phó Bí thư" },
			{ "TÃ¦ng cÃ´m", "Tổng cục" },
			{ "UÃ» ban", "Ủy ban" },
			{ "HÃ©i Â®Ã¥ng", "Hội đồng" },
			{ "ChÃnh phÃ±", "Chính phủ" }
		};
		Range content = doc.Content;
		for (int i = 0; i < array.GetLength(0); i++)
		{
			ReplaceText(content, array[i, 0], array[i, 1]);
		}
	}

	public static void StandardizeToneMarks(Document doc, Action<int, string> onProgress = null)
	{
		if (doc == null)
		{
			return;
		}
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Chuẩn Hóa Dấu Thanh"))
			{
				Range content = doc.Content;
				onProgress?.Invoke(89, "Đang chuẩn hóa dấu thanh vần oa, oe (hòa, thỏa, xòe)...");
				string[,] array = new string[121, 2]
				{
					{ "hoà", "hòa" },
					{ "Hoà", "Hòa" },
					{ "HOÀ", "HÒA" },
					{ "toà", "tòa" },
					{ "Toà", "Tòa" },
					{ "TOÀ", "TÒA" },
					{ "thoả", "thỏa" },
					{ "Thoả", "Thỏa" },
					{ "THOẢ", "THỎA" },
					{ "xoá", "xóa" },
					{ "Xoá", "Xóa" },
					{ "XOÁ", "XÓA" },
					{ "khoá", "khóa" },
					{ "Khoá", "Khóa" },
					{ "KHOÁ", "KHÓA" },
					{ "hoá", "hóa" },
					{ "Hoá", "Hóa" },
					{ "HOÁ", "HÓA" },
					{ "toả", "tỏa" },
					{ "Toả", "Tỏa" },
					{ "TOẢ", "TỎA" },
					{ "doạ", "dọa" },
					{ "Doạ", "Dọa" },
					{ "DOẠ", "DỌA" },
					{ "toạ", "tọa" },
					{ "Toạ", "Tọa" },
					{ "TOẠ", "TỌA" },
					{ "hoạ", "họa" },
					{ "Hoạ", "Họa" },
					{ "HOẠ", "HỌA" },
					{ "loà", "lòa" },
					{ "Loà", "Lòa" },
					{ "LOÀ", "LÒA" },
					{ "loá", "lóa" },
					{ "Loá", "Lóa" },
					{ "LOÁ", "LÓA" },
					{ "xoà", "xòa" },
					{ "Xoà", "Xòa" },
					{ "XOÀ", "XÒA" },
					{ "choã", "chõa" },
					{ "Choã", "Chõa" },
					{ "CHOÃ", "CHÕA" },
					{ "đoá", "đóa" },
					{ "Đoá", "Đóa" },
					{ "ĐOÁ", "ĐÓA" },
					{ "soá", "sóa" },
					{ "Soá", "Sóa" },
					{ "SOÁ", "SÓA" },
					{ "thoa", "thoa" },
					{ "Thoa", "Thoa" },
					{ "thoá", "thóa" },
					{ "Thoá", "Thóa" },
					{ "thoạ", "thọa" },
					{ "Thoạ", "Thọa" },
					{ "chòa", "chòa" },
					{ "Chòa", "Chòa" },
					{ "hỏa", "hỏa" },
					{ "Hỏa", "Hỏa" },
					{ "hoè", "hòe" },
					{ "Hoè", "Hòe" },
					{ "HOÈ", "HÒE" },
					{ "khoẻ", "khỏe" },
					{ "Khoẻ", "Khỏe" },
					{ "KHOẺ", "KHỎE" },
					{ "loè", "lòe" },
					{ "Loè", "Lòe" },
					{ "LOÈ", "LÒE" },
					{ "loé", "lóe" },
					{ "Loé", "Lóe" },
					{ "LOÉ", "LÓE" },
					{ "xoè", "xòe" },
					{ "Xoè", "Xòe" },
					{ "XOÈ", "XÒE" },
					{ "xoé", "xóe" },
					{ "Xoé", "Xóe" },
					{ "XOÉ", "XÓE" },
					{ "chòe", "chòe" },
					{ "Chòe", "Chòe" },
					{ "CHÒE", "CHÒE" },
					{ "chóe", "chóe" },
					{ "Chóe", "Chóe" },
					{ "CHÓE", "CHÓE" },
					{ "nhòe", "nhòe" },
					{ "Nhòe", "Nhòe" },
					{ "NHÒE", "NHÒE" },
					{ "tóe", "tóe" },
					{ "Tóe", "Tóe" },
					{ "TÓE", "TÓE" },
					{ "tòe", "tòe" },
					{ "Tòe", "Tòe" },
					{ "TÒE", "TÒE" },
					{ "oà", "òa" },
					{ "Oà", "Òa" },
					{ "OÀ", "ÒA" },
					{ "oá", "óa" },
					{ "Oá", "Óa" },
					{ "OÁ", "ÓA" },
					{ "oả", "ỏa" },
					{ "Oả", "Ỏa" },
					{ "OẢ", "ỎA" },
					{ "oã", "õa" },
					{ "Oã", "Õa" },
					{ "OÃ", "ÕA" },
					{ "oạ", "ọa" },
					{ "Oạ", "Ọa" },
					{ "OẠ", "ỌA" },
					{ "oè", "òe" },
					{ "Oè", "Òe" },
					{ "OÈ", "ÒE" },
					{ "oé", "óe" },
					{ "Oé", "Óe" },
					{ "OÉ", "ÓE" },
					{ "oẻ", "ỏe" },
					{ "Oẻ", "Ỏe" },
					{ "OẺ", "ỎE" },
					{ "oẽ", "õe" },
					{ "Oẽ", "Õe" },
					{ "OẼ", "ÕE" },
					{ "oẹ", "ọe" },
					{ "Oẹ", "Ọe" },
					{ "OẸ", "ỌE" }
				};
				for (int i = 0; i < array.GetLength(0); i++)
				{
					ReplaceText(content, array[i, 0], array[i, 1], matchCase: true, matchWholeWord: true);
				}
				onProgress?.Invoke(91, "Đang chuẩn hóa dấu thanh vần uy (thủy, hủy, ủy, tùy)...");
				string[,] array2 = new string[30, 2]
				{
					{ "thuỷ", "thủy" },
					{ "Thuỷ", "Thủy" },
					{ "THUỶ", "THỦY" },
					{ "huỷ", "hủy" },
					{ "Huỷ", "Hủy" },
					{ "HUỶ", "HỦY" },
					{ "uỷ", "ủy" },
					{ "Uỷ", "Ủy" },
					{ "UỶ", "ỦY" },
					{ "luỹ", "lũy" },
					{ "Luỹ", "Lũy" },
					{ "LUỸ", "LŨY" },
					{ "thuỵ", "thụy" },
					{ "Thuỵ", "Thụy" },
					{ "THUỴ", "THỤY" },
					{ "tuỳ", "tùy" },
					{ "Tuỳ", "Tùy" },
					{ "TUỲ", "TÙY" },
					{ "tuý", "túy" },
					{ "Tuý", "Túy" },
					{ "TUÝ", "TÚY" },
					{ "nhuỵ", "nhụy" },
					{ "Nhuỵ", "Nhụy" },
					{ "NHUỴ", "NHỤY" },
					{ "luỵ", "lụy" },
					{ "Luỵ", "Lụy" },
					{ "LUỴ", "LỤY" },
					{ "thuỳ", "thùy" },
					{ "Thuỳ", "Thùy" },
					{ "THUỲ", "THÙY" }
				};
				for (int j = 0; j < array2.GetLength(0); j++)
				{
					ReplaceText(content, array2[j, 0], array2[j, 1], matchCase: true, matchWholeWord: true);
				}
				onProgress?.Invoke(94, "Đang sửa lỗi dấu thanh âm đệm (hoàn, toàn, đoàn, khoản)...");
				ApplyClosedSyllableFixes(content);
				onProgress?.Invoke(96, "Đang sửa triệt để dấu thanh âm đầu qu (quản, Quý, quả)...");
				FixCorruptedQuWords(content);
			}
		}
	}

	public static void ConvertToTraditionalToneMarks(Document doc)
	{
		if (doc == null)
		{
			return;
		}
		using (new ScreenFreezeScope(doc.Application))
		{
			using (new SafeUndoScope(doc.Application, "OfficeFormat Dấu Thanh Kiểu Cũ"))
			{
				Range content = doc.Content;
				string[,] array = new string[115, 2]
				{
					{ "hòa", "hoà" },
					{ "Hòa", "Hoà" },
					{ "HÒA", "HOÀ" },
					{ "tòa", "toà" },
					{ "Tòa", "Toà" },
					{ "TÒA", "TOÀ" },
					{ "thỏa", "thoả" },
					{ "Thỏa", "Thoả" },
					{ "THỎA", "THOẢ" },
					{ "xóa", "xoá" },
					{ "Xóa", "Xoá" },
					{ "XÓA", "XOÁ" },
					{ "khóa", "khoá" },
					{ "Khóa", "Khoá" },
					{ "KHÓA", "KHOÁ" },
					{ "hóa", "hoá" },
					{ "Hóa", "Hoá" },
					{ "HÓA", "HOÁ" },
					{ "tỏa", "toả" },
					{ "Tỏa", "Toả" },
					{ "TỎA", "TOẢ" },
					{ "dọa", "doạ" },
					{ "Dọa", "Doạ" },
					{ "DỌA", "DOẠ" },
					{ "tọa", "toạ" },
					{ "Tọa", "Toạ" },
					{ "TỌA", "TOẠ" },
					{ "họa", "hoạ" },
					{ "Họa", "Hoạ" },
					{ "HỌA", "HOẠ" },
					{ "lòa", "loà" },
					{ "Lòa", "Loà" },
					{ "LÒA", "LOÀ" },
					{ "lóa", "loá" },
					{ "Lóa", "Loá" },
					{ "LÓA", "LOÁ" },
					{ "xòa", "xoà" },
					{ "Xòa", "Xoà" },
					{ "XÒA", "XOÀ" },
					{ "chõa", "choã" },
					{ "Chõa", "Choã" },
					{ "CHÕA", "CHOÃ" },
					{ "đóa", "đoá" },
					{ "Đóa", "Đoá" },
					{ "ĐÓA", "ĐOÁ" },
					{ "sóa", "soá" },
					{ "Sóa", "Soá" },
					{ "SÓA", "SOÁ" },
					{ "thóa", "thoá" },
					{ "Thóa", "Thoá" },
					{ "thọa", "thoạ" },
					{ "Thọa", "Thoạ" },
					{ "hòe", "hoè" },
					{ "Hòe", "Hoè" },
					{ "HÒE", "HOÈ" },
					{ "khỏe", "khoẻ" },
					{ "Khỏe", "Khoẻ" },
					{ "KHỎE", "KHOẺ" },
					{ "lòe", "loè" },
					{ "Lòe", "Loè" },
					{ "LÒE", "LOÈ" },
					{ "lóe", "loé" },
					{ "Lóe", "Loé" },
					{ "LÓE", "LOÉ" },
					{ "xòe", "xoè" },
					{ "Xòe", "Xoè" },
					{ "XÒE", "XOÈ" },
					{ "xóe", "xoé" },
					{ "Xóe", "Xoé" },
					{ "XÓE", "XOÉ" },
					{ "chòe", "chòe" },
					{ "Chòe", "Chòe" },
					{ "CHÒE", "CHÒE" },
					{ "chóe", "chóe" },
					{ "Chóe", "Chóe" },
					{ "CHÓE", "CHÓE" },
					{ "nhòe", "nhòe" },
					{ "Nhòe", "Nhòe" },
					{ "NHÒE", "NHÒE" },
					{ "tóe", "tóe" },
					{ "Tóe", "Tóe" },
					{ "TÓE", "TÓE" },
					{ "tòe", "tòe" },
					{ "Tòe", "Tòe" },
					{ "TÒE", "TÒE" },
					{ "òa", "oà" },
					{ "Òa", "Oà" },
					{ "ÒA", "OÀ" },
					{ "óa", "oá" },
					{ "Óa", "Oá" },
					{ "ÓA", "OÁ" },
					{ "ỏa", "oả" },
					{ "Ỏa", "Oả" },
					{ "ỎA", "OẢ" },
					{ "õa", "oã" },
					{ "Õa", "Oã" },
					{ "ÕA", "OÃ" },
					{ "ọa", "oạ" },
					{ "Ọa", "Oạ" },
					{ "ỌA", "OẠ" },
					{ "òe", "oè" },
					{ "Òe", "Oè" },
					{ "ÒE", "OÈ" },
					{ "óe", "oé" },
					{ "Óe", "Oé" },
					{ "ÓE", "OÉ" },
					{ "ỏe", "oẻ" },
					{ "Ỏe", "Oẻ" },
					{ "ỎE", "OẺ" },
					{ "õe", "oẽ" },
					{ "Õe", "Oẽ" },
					{ "ÕE", "OẼ" },
					{ "ọe", "oẹ" },
					{ "Ọe", "Oẹ" },
					{ "ỌE", "OẸ" }
				};
				for (int i = 0; i < array.GetLength(0); i++)
				{
					ReplaceText(content, array[i, 0], array[i, 1], matchCase: true, matchWholeWord: true);
				}
				string[,] array2 = new string[30, 2]
				{
					{ "thủy", "thuỷ" },
					{ "Thủy", "Thuỷ" },
					{ "THỦY", "THUỶ" },
					{ "hủy", "huỷ" },
					{ "Hủy", "Huỷ" },
					{ "HỦY", "HUỶ" },
					{ "ủy", "uỷ" },
					{ "Ủy", "Uỷ" },
					{ "ỦY", "UỶ" },
					{ "lũy", "luỹ" },
					{ "Lũy", "Luỹ" },
					{ "LŨY", "LUỸ" },
					{ "thụy", "thuỵ" },
					{ "Thụy", "Thuỵ" },
					{ "THỤY", "THUỴ" },
					{ "tùy", "tuỳ" },
					{ "Tùy", "Tuỳ" },
					{ "TÙY", "TUỲ" },
					{ "túy", "tuý" },
					{ "Túy", "Tuý" },
					{ "TÚY", "TUÝ" },
					{ "nhụy", "nhuỵ" },
					{ "Nhụy", "Nhuỵ" },
					{ "NHỤY", "NHUỴ" },
					{ "lụy", "luỵ" },
					{ "Lụy", "Luỵ" },
					{ "LỤY", "LUỴ" },
					{ "thùy", "thuỳ" },
					{ "Thùy", "Thuỳ" },
					{ "THÙY", "THUỲ" }
				};
				for (int j = 0; j < array2.GetLength(0); j++)
				{
					ReplaceText(content, array2[j, 0], array2[j, 1], matchCase: true, matchWholeWord: true);
				}
				ApplyClosedSyllableFixes(content);
				FixCorruptedQuWords(content);
			}
		}
	}

	private static void ApplyClosedSyllableFixes(Range rng)
	{
		string[,] array = new string[62, 2]
		{
			{ "hòan", "hoàn" },
			{ "Hòan", "Hoàn" },
			{ "HÒAN", "HOÀN" },
			{ "tòan", "toàn" },
			{ "Tòan", "Toàn" },
			{ "TÒAN", "TOÀN" },
			{ "đòan", "đoàn" },
			{ "Đòan", "Đoàn" },
			{ "ĐÒAN", "ĐOÀN" },
			{ "khỏan", "khoản" },
			{ "Khỏan", "Khoản" },
			{ "KHỎAN", "KHOẢN" },
			{ "lọai", "loại" },
			{ "Lọai", "Loại" },
			{ "LỌAI", "LOẠI" },
			{ "ngọai", "ngoại" },
			{ "Ngọai", "Ngoại" },
			{ "NGỌAI", "NGOẠI" },
			{ "tóan", "toán" },
			{ "Tóan", "Toán" },
			{ "TÓAN", "TOÁN" },
			{ "thóat", "thoát" },
			{ "Thóat", "Thoát" },
			{ "THÓAT", "THOÁT" },
			{ "họat", "hoạt" },
			{ "Họat", "Hoạt" },
			{ "HỌAT", "HOẠT" },
			{ "thọai", "thoại" },
			{ "Thọai", "Thoại" },
			{ "THỌAI", "THOẠI" },
			{ "lọat", "loạt" },
			{ "Lọat", "Loạt" },
			{ "LỌAT", "LOẠT" },
			{ "đọan", "đoạn" },
			{ "Đọan", "Đoạn" },
			{ "ĐỌAN", "ĐOẠN" },
			{ "khóan", "khoán" },
			{ "Khóan", "Khoán" },
			{ "KHÓAN", "KHOÁN" },
			{ "xòai", "xoài" },
			{ "Xòai", "Xoài" },
			{ "XÒAI", "XOÀI" },
			{ "khóai", "khoái" },
			{ "Khóai", "Khoái" },
			{ "KHÓAI", "KHOÁI" },
			{ "hóan", "hoán" },
			{ "Hóan", "Hoán" },
			{ "HÓAN", "HOÁN" },
			{ "hóang", "hoáng" },
			{ "Hóang", "Hoáng" },
			{ "khóang", "khoáng" },
			{ "Khóang", "Khoáng" },
			{ "thóang", "thoáng" },
			{ "Thóang", "Thoáng" },
			{ "khỏang", "khoảng" },
			{ "Khỏang", "Khoảng" },
			{ "hỏang", "hoảng" },
			{ "Hỏang", "Hoảng" },
			{ "túyp", "tuýp" },
			{ "Túyp", "Tuýp" },
			{ "súyt", "suýt" },
			{ "Súyt", "Suýt" }
		};
		for (int i = 0; i < array.GetLength(0); i++)
		{
			ReplaceText(rng, array[i, 0], array[i, 1], matchCase: true);
		}
	}

	private static void FixCorruptedQuWords(Range rng)
	{
		string[,] array = new string[77, 2]
		{
			{ "qủan", "quản" },
			{ "Qủan", "Quản" },
			{ "QỦAN", "QUẢN" },
			{ "Qúy", "Quý" },
			{ "qúy", "quý" },
			{ "QÚY", "QUÝ" },
			{ "Qùy", "Quỳ" },
			{ "qùy", "quỳ" },
			{ "QÙY", "QUỲ" },
			{ "Qủy", "Quỷ" },
			{ "qủy", "quỷ" },
			{ "QỦY", "QUỶ" },
			{ "Qũy", "Quỹ" },
			{ "qũy", "quỹ" },
			{ "QŨY", "QUỸ" },
			{ "Qụy", "Quỵ" },
			{ "qụy", "quỵ" },
			{ "QỤY", "QUỴ" },
			{ "qủa", "quả" },
			{ "Qủa", "Quả" },
			{ "QỦA", "QUẢ" },
			{ "qúan", "quán" },
			{ "Qúan", "Quán" },
			{ "QÚAN", "QUÁN" },
			{ "qùa", "quà" },
			{ "Qùa", "Quà" },
			{ "QÙA", "QUÀ" },
			{ "qúa", "quá" },
			{ "Qúa", "Quá" },
			{ "QÚA", "QUÁ" },
			{ "qũa", "quã" },
			{ "Qũa", "Quã" },
			{ "QŨA", "QUÃ" },
			{ "qụa", "quạ" },
			{ "Qụa", "Quạ" },
			{ "QỤA", "QUẠ" },
			{ "qủang", "quảng" },
			{ "Qủang", "Quảng" },
			{ "QỦANG", "QUẢNG" },
			{ "qúang", "quáng" },
			{ "Qúang", "Quáng" },
			{ "QÚANG", "QUÁNG" },
			{ "qùang", "quàng" },
			{ "Qùang", "Quàng" },
			{ "QÙANG", "QUÀNG" },
			{ "qũang", "quãng" },
			{ "Qũang", "Quãng" },
			{ "QŨANG", "QUÃNG" },
			{ "qụang", "quạng" },
			{ "Qụang", "Quạng" },
			{ "QỤANG", "QUẠNG" },
			{ "qúat", "quát" },
			{ "Qúat", "Quát" },
			{ "QÚAT", "QUÁT" },
			{ "qụat", "quạt" },
			{ "Qụat", "Quạt" },
			{ "QỤAT", "QUẠT" },
			{ "qúanh", "quánh" },
			{ "Qúanh", "Quánh" },
			{ "qùanh", "quành" },
			{ "Qùanh", "Quành" },
			{ "qủanh", "quảnh" },
			{ "Qủanh", "Quảnh" },
			{ "qụanh", "quạnh" },
			{ "Qụanh", "Quạnh" },
			{ "qúam", "quám" },
			{ "qủam", "quảm" },
			{ "qúap", "quáp" },
			{ "qụap", "quạp" },
			{ "qúe", "qué" },
			{ "Qúe", "Qué" },
			{ "qùe", "què" },
			{ "Qùe", "Què" },
			{ "qủe", "quẻ" },
			{ "Qủe", "Quẻ" },
			{ "qụe", "quẹ" },
			{ "Qụe", "Quẹ" }
		};
		for (int i = 0; i < array.GetLength(0); i++)
		{
			ReplaceText(rng, array[i, 0], array[i, 1], matchCase: true);
		}
	}

	private static string RemoveVietnameseDiacritics(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return text;
		}
		string[] array = new string[15]
		{
			"aAeEoOuUiIdDyY", "áàạảãâấầậẩẫăắằặẳẵ", "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ", "éèẹẻẽêếềệểễ", "ÉÈẸẺẼÊẾỀỆỂỄ", "óòọỏõôốồộổỗơớờợởỡ", "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ", "úùụủũưứừựửữ", "ÚÙỤỦŨƯỨỪỰỬỮ", "íìịỉĩ",
			"ÍÌỊỈĨ", "đ", "Đ", "ýỳỵỷỹ", "ÝỲỴỶỸ"
		};
		for (int i = 1; i < array.Length; i++)
		{
			for (int j = 0; j < array[i].Length; j++)
			{
				text = text.Replace(array[i][j], array[0][i - 1]);
			}
		}
		return text;
	}

	public static void InsertCalloutBox(Application app, string text, bool isRed, float widthMm = 36f, float heightMm = 8.5f)
	{
		if (app == null || app.Documents.Count == 0)
		{
			return;
		}
		Selection selection = app.Selection;
		if (selection == null)
		{
			return;
		}
		using (new SafeUndoScope(app, "OfficeFormat Chèn Dấu Chỉ Độ Mật/Khẩn"))
		{
			try
			{
				float num = 0f;
				float num2 = 0f;
				try
				{
					num = Convert.ToSingle((dynamic)selection.get_Information(WdInformation.wdHorizontalPositionRelativeToPage));
					num2 = Convert.ToSingle((dynamic)selection.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
				}
				catch
				{
				}
				if (num <= 0f)
				{
					num = UnitConverter.MmToPt(30f);
				}
				if (num2 <= 0f)
				{
					num2 = UnitConverter.MmToPt(25f);
				}
				Shapes shapes = app.ActiveDocument.Shapes;
				float left = num;
				float top = num2;
				float width = UnitConverter.MmToPt(widthMm);
				float height = UnitConverter.MmToPt(heightMm);
				object Anchor = Missing.Value;
				Shape shape = shapes.AddShape(1, left, top, width, height, ref Anchor);
				shape.WrapFormat.Type = WdWrapType.wdWrapNone;
				shape.Fill.Solid();
				shape.Fill.ForeColor.RGB = 16777215;
				shape.Line.ForeColor.RGB = (isRed ? 255 : 0);
				shape.Line.Weight = 1f;
				TextFrame textFrame = shape.TextFrame;
				textFrame.MarginTop = 0f;
				textFrame.MarginBottom = 0f;
				textFrame.MarginLeft = 0f;
				textFrame.MarginRight = 0f;
				Range textRange = textFrame.TextRange;
				textRange.Text = text;
				textRange.Font.Name = "Times New Roman";
				textRange.Font.Size = 13f;
				textRange.Font.Bold = 1;
				textRange.Font.Color = (isRed ? WdColor.wdColorRed : WdColor.wdColorBlack);
				textRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				textRange.ParagraphFormat.SpaceBefore = 0f;
				textRange.ParagraphFormat.SpaceAfter = 0f;
				textRange.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				object Replace = Missing.Value;
				shape.Select(ref Replace);
			}
			catch
			{
			}
		}
	}

	public static void InsertDocTitle(Selection sel, string title, string subject)
	{
		if (sel == null)
		{
			return;
		}
		using (new SafeUndoScope(sel.Application, "OfficeFormat Chèn Tên Loại Văn Bản"))
		{
			Range range = sel.Range;
			range.Text = title + "\r\n" + (string.IsNullOrEmpty(subject) ? "" : (subject + "\r\n"));
			Paragraph paragraph = range.Paragraphs[1];
			paragraph.Range.Font.Name = "Times New Roman";
			paragraph.Range.Font.Size = 14f;
			paragraph.Range.Font.Bold = 1;
			paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph.Range.ParagraphFormat.SpaceBefore = 6f;
			paragraph.Range.ParagraphFormat.SpaceAfter = 3f;
			if (!string.IsNullOrEmpty(subject) && range.Paragraphs.Count >= 2)
			{
				Paragraph paragraph2 = range.Paragraphs[2];
				paragraph2.Range.Font.Name = "Times New Roman";
				paragraph2.Range.Font.Size = 13f;
				paragraph2.Range.Font.Bold = 1;
				paragraph2.Range.Font.Italic = 1;
				paragraph2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				paragraph2.Range.ParagraphFormat.SpaceBefore = 0f;
				paragraph2.Range.ParagraphFormat.SpaceAfter = 6f;
			}
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range.Collapse(ref Direction);
		}
	}

	private static void SetTableColumnWidths(Table table, float col1Width, float col2Width)
	{
		if (table == null) return;
		try
		{
			table.AllowAutoFit = false;
			table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
		}
		catch { }

		try
		{
			foreach (Row row in table.Rows)
			{
				if (row.Cells.Count >= 2)
				{
					try { row.Cells[1].Width = col1Width; } catch { }
					try { row.Cells[2].Width = col2Width; } catch { }
				}
			}
		}
		catch { }

		try { table.Columns[1].Width = col1Width; } catch { }
		try { table.Columns[2].Width = col2Width; } catch { }
	}

	public static void InsertCustomSignature(Selection sel, string noiNhanText, string signerTitle, string signerName)
	{
		string upperTitle = (signerTitle ?? "").ToUpperInvariant();
		string upperRec = (noiNhanText ?? "").ToUpperInvariant();
		bool isParty = upperTitle.Contains("CHI BỘ") || upperTitle.Contains("BAN CHẤP HÀNH") || upperTitle.Contains("BAN THƯỜNG VỤ") || upperTitle.Contains("BÍ THƯ") || upperTitle.Contains("ĐẢNG ỦY") || upperTitle.Contains("TỈNH ỦY") || upperTitle.Contains("THÀNH ỦY") || upperTitle.Contains("ĐẢNG") || upperTitle.Contains("T/M") || upperTitle.Contains("K/T") || upperTitle.Contains("T/L") || upperRec.Contains("CHI BỘ") || upperRec.Contains("ĐẢNG ỦY") || upperRec.Contains("ĐẢNG");
		InsertCustomSignature(sel, noiNhanText, signerTitle, signerName, isParty);
	}

	public static void InsertCustomSignature(Selection sel, string noiNhanText, string signerTitle, string signerName, bool isParty)
	{
		if (sel == null)
		{
			return;
		}
		using (new SafeUndoScope(sel.Application, "OfficeFormat Chèn Khung Chữ Ký & Nơi Nhận"))
		{
			Range range = sel.Range;
			Tables tables = range.Tables;
			object DefaultTableBehavior = Missing.Value;
			object AutoFitBehavior = Missing.Value;
			Table table = tables.Add(range, 1, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
			table.Borders.Enable = 0;
			SetTableColumnWidths(table, UnitConverter.MmToPt(75f), UnitConverter.MmToPt(95f));
			Cell cell = table.Cell(1, 1);
			cell.Range.Text = "Nơi nhận:\r\n" + noiNhanText;
			cell.Range.Font.Name = "Times New Roman";
			cell.Range.Font.Size = (isParty ? 12f : 11f);
			cell.Range.Font.Bold = 0;
			cell.Range.Font.Italic = 0;
			cell.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
			cell.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			cell.Range.ParagraphFormat.SpaceBefore = 0f;
			cell.Range.ParagraphFormat.SpaceAfter = 0f;
			cell.Range.ParagraphFormat.FirstLineIndent = 0f;
			int pCount = cell.Range.Paragraphs.Count;
			if (pCount > 0)
			{
				Paragraph paragraph = cell.Range.Paragraphs[1];
				paragraph.Range.Font.Bold = (isParty ? 0 : (-1));
				paragraph.Range.Font.Italic = (isParty ? 0 : (-1));
				paragraph.Range.Font.Underline = (isParty ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
				paragraph.Range.Font.Size = (isParty ? 14f : 12f);
				for (int i = 2; i <= pCount; i++)
				{
					Paragraph itemPara = cell.Range.Paragraphs[i];
					itemPara.Range.Font.Bold = 0;
					itemPara.Range.Font.Italic = 0;
					itemPara.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					itemPara.Range.Font.Size = (isParty ? 12f : 11f);
				}
			}
			Cell cell2 = table.Cell(1, 2);
			cell2.Range.Text = signerTitle + "\r\n\r\n\r\n\r\n" + signerName;
			cell2.Range.Font.Name = "Times New Roman";
			cell2.Range.Font.Size = (isParty ? 14f : 13f);
			cell2.Range.Font.Bold = -1;
			cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			cell2.Range.ParagraphFormat.SpaceBefore = 0f;
			cell2.Range.ParagraphFormat.SpaceAfter = 0f;
			if (isParty)
			{
				int sigPCount = cell2.Range.Paragraphs.Count;
				for (int j = 1; j < sigPCount; j++)
				{
					Paragraph sigP = cell2.Range.Paragraphs[j];
					string sigLine = (sigP.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (sigLine.Length > 0 && !RegexPatterns.SignAuthority.IsMatch(sigLine))
					{
						sigP.Range.Font.Bold = 0;
					}
				}
			}
			Range range2 = table.Range;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range2.Collapse(ref Direction);
			table.Range.InsertParagraphAfter();
		}
	}

	public static void InsertCustomHeader(Selection sel, string upperOrgan, string lowerOrgan, string docNum, string locationDate, bool isParty, bool isMilitary)
	{
		if (sel == null)
		{
			return;
		}
		using (new SafeUndoScope(sel.Application, "OfficeFormat Chèn Khung Tiêu Đề Cơ Quan"))
		{
			Range range = sel.Range;
			Tables tables = range.Tables;
			object DefaultTableBehavior = Missing.Value;
			object AutoFitBehavior = Missing.Value;
			Table table = tables.Add(range, 2, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
			table.Borders.Enable = 0;
			SetTableColumnWidths(table, UnitConverter.MmToPt(isParty ? 80f : 65f), UnitConverter.MmToPt(isParty ? 90f : 105f));
			Cell cell = table.Cell(1, 1);
			cell.Range.Text = (string.IsNullOrEmpty(upperOrgan) ? "" : (upperOrgan + "\r\n")) + lowerOrgan;
			cell.Range.Font.Name = "Times New Roman";
			cell.Range.Font.Size = 12f;
			cell.Range.Font.Bold = 1;
			cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			Cell cell2 = table.Cell(2, 1);
			Range range2 = cell2.Range;
			string text;
			if (string.IsNullOrEmpty(docNum))
			{
				text = (isParty ? "Số     -..." : "Số: ... / ...");
			}
			else
			{
				text = docNum;
			}
			range2.Text = text;
			cell2.Range.Font.Name = "Times New Roman";
			cell2.Range.Font.Size = (isParty ? 14f : 13f);
			cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			Cell cell3 = table.Cell(1, 2);
			if (isParty)
			{
				cell3.Range.Text = "ĐẢNG CỘNG SẢN VIỆT NAM";
				cell3.Range.Font.Name = "Times New Roman";
				cell3.Range.Font.Size = 14f;
				cell3.Range.Font.Bold = 1;
			}
			else
			{
				cell3.Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
				cell3.Range.Font.Name = "Times New Roman";
				if (cell3.Range.Paragraphs.Count > 0)
				{
					cell3.Range.Paragraphs[1].Range.Font.Size = 12f;
					cell3.Range.Paragraphs[1].Range.Font.Bold = 1;
				}
				if (cell3.Range.Paragraphs.Count > 1)
				{
					cell3.Range.Paragraphs[2].Range.Font.Size = 13f;
					cell3.Range.Paragraphs[2].Range.Font.Bold = 1;
				}
			}
			cell3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell3.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			Cell cell4 = table.Cell(2, 2);
			cell4.Range.Text = (string.IsNullOrEmpty(locationDate) ? "..., ngày ... tháng ... năm 20..." : locationDate);
			cell4.Range.Font.Name = "Times New Roman";
			cell4.Range.Font.Size = 13f;
			cell4.Range.Font.Italic = 1;
			cell4.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell4.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			VFormat.WordEngine.Builders.HeaderRebuilder.FormatExistingHeaderTable(table, isParty ? RulePackType.HD05_VPTW : (isMilitary ? RulePackType.TT68_2023_BQP : RulePackType.ND30_2020), profile: null, skipRebuildCheck: true);
			Range range3 = table.Range;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range3.Collapse(ref Direction);
			table.Range.InsertParagraphAfter();
		}
	}

	public static void InsertYouthUnionHeader(Selection sel, string upperOrgan, string lowerOrgan, string docNum, string locationDate)
	{
		if (sel == null)
		{
			return;
		}
		using (new SafeUndoScope(sel.Application, "OfficeFormat Chèn Khối Đầu Văn Bản Đoàn"))
		{
			Range range = sel.Range;
			Tables tables = range.Tables;
			object DefaultTableBehavior = Missing.Value;
			object AutoFitBehavior = Missing.Value;
			Table table = tables.Add(range, 2, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
			table.Borders.Enable = 0;
			float num = 0f;
			try
			{
				num = sel.Document.PageSetup.PageWidth - sel.Document.PageSetup.LeftMargin - sel.Document.PageSetup.RightMargin;
			}
			catch
			{
			}
			if (num < UnitConverter.MmToPt(120f) || num > UnitConverter.MmToPt(200f))
			{
				num = UnitConverter.MmToPt(160f);
			}
			float num2 = UnitConverter.MmToPt(78f);
			SetTableColumnWidths(table, num - num2, num2);
			Cell cell = table.Cell(1, 1);
			bool flag = !string.IsNullOrEmpty(upperOrgan);
			string text = (string.IsNullOrEmpty(lowerOrgan) ? "BCH ĐOÀN..." : lowerOrgan);
			cell.Range.Text = (flag ? (upperOrgan.ToUpperInvariant() + "\r") : "") + text.Replace("\r\n", "\r").Replace("\n", "\r").ToUpperInvariant() + "\r***";
			cell.Range.Font.Name = "Times New Roman";
			cell.Range.Font.Size = 14f;
			cell.Range.Font.Bold = -1;
			cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			cell.Range.ParagraphFormat.SpaceBefore = 0f;
			cell.Range.ParagraphFormat.SpaceAfter = 0f;
			if (flag && cell.Range.Paragraphs.Count > 0)
			{
				cell.Range.Paragraphs[1].Range.Font.Bold = 0;
			}
			int count = cell.Range.Paragraphs.Count;
			if (count > 0)
			{
				cell.Range.Paragraphs[count].Range.Font.Bold = 0;
			}
			Cell cell2 = table.Cell(2, 1);
			cell2.Range.Text = (string.IsNullOrEmpty(docNum) ? "Số:     -.../..." : docNum);
			cell2.Range.Font.Name = "Times New Roman";
			cell2.Range.Font.Size = 14f;
			cell2.Range.Font.Bold = 0;
			cell2.Range.Font.Italic = 0;
			cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			Cell cell3 = table.Cell(1, 2);
			cell3.Range.Text = "ĐOÀN TNCS HỒ CHÍ MINH";
			cell3.Range.Font.Name = "Times New Roman";
			cell3.Range.Font.Size = 15f;
			cell3.Range.Font.Bold = -1;
			cell3.Range.Font.Italic = 0;
			cell3.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
			cell3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell3.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			Cell cell4 = table.Cell(2, 2);
			cell4.Range.Text = (string.IsNullOrEmpty(locationDate) ? "..., ngày ... tháng ... năm 20..." : locationDate);
			cell4.Range.Font.Name = "Times New Roman";
			cell4.Range.Font.Size = 13f;
			cell4.Range.Font.Bold = 0;
			cell4.Range.Font.Italic = -1;
			cell4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			cell4.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell4.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			VFormat.WordEngine.Builders.HeaderRebuilder.FormatExistingHeaderTable(table, RulePackType.HD29_TWDTN, profile: null, skipRebuildCheck: true);
			Range range2 = table.Range;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range2.Collapse(ref Direction);
			table.Range.InsertParagraphAfter();
		}
	}

	public static void InsertPersonalMottoHeader(Selection sel, string locationDate = null)
	{
		if (sel == null)
		{
			return;
		}
		using (new SafeUndoScope(sel.Application, "OfficeFormat Chèn Tiêu Ngữ Đơn Từ Cá Nhân"))
		{
			Range range = sel.Range;
			string text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc\r\n";
			if (!string.IsNullOrEmpty(locationDate))
			{
				text = text + locationDate + "\r\n";
			}
			range.Text = text;
			Paragraph paragraph = range.Paragraphs[1];
			paragraph.Range.Font.Name = "Times New Roman";
			paragraph.Range.Font.Size = 12f;
			paragraph.Range.Font.Bold = 1;
			paragraph.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			paragraph.Range.ParagraphFormat.SpaceBefore = 0f;
			paragraph.Range.ParagraphFormat.SpaceAfter = 0f;
			Paragraph paragraph2 = range.Paragraphs[2];
			paragraph2.Range.Font.Name = "Times New Roman";
			paragraph2.Range.Font.Size = 13f;
			paragraph2.Range.Font.Bold = 1;
			paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			paragraph2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			paragraph2.Range.ParagraphFormat.SpaceBefore = 0f;
			paragraph2.Range.ParagraphFormat.SpaceAfter = 6f;
			UnderlineFormatter.EnsurePersonalMottoShapeLine(sel.Document, paragraph2);
			if (!string.IsNullOrEmpty(locationDate) && range.Paragraphs.Count >= 3)
			{
				Paragraph paragraph3 = range.Paragraphs[3];
				paragraph3.Range.Font.Name = "Times New Roman";
				paragraph3.Range.Font.Size = 13f;
				paragraph3.Range.Font.Italic = 1;
				paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				paragraph3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				paragraph3.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				paragraph3.Range.ParagraphFormat.SpaceBefore = 0f;
				paragraph3.Range.ParagraphFormat.SpaceAfter = 6f;
			}
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range.Collapse(ref Direction);
		}
	}

	public static void InsertFooterAgencyInfo(Document doc, Selection sel = null, string customInfo = null)
	{
		if (doc == null)
		{
			return;
		}
		using (new SafeUndoScope(doc.Application, "OfficeFormat Chèn Thông Tin Chân Trang Cơ Quan"))
		{
			string text = customInfo;
			if (string.IsNullOrEmpty(text))
			{
				text = "Địa chỉ: ...; Điện thoại: ...; Fax: ...; Email: ...; Website: ...";
			}
			try
			{
				if (doc.Sections.Count > 0)
				{
					Section section = doc.Sections[1];
					section.PageSetup.DifferentFirstPageHeaderFooter = -1;
					HeaderFooter headerFooter = section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
					Range range = headerFooter.Range;
					range.Text = text;
					range.Font.Name = "Times New Roman";
					range.Font.Size = 11f;
					range.Font.Bold = 0;
					range.Font.Italic = 0;
					range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
					range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					range.ParagraphFormat.SpaceBefore = 4f;
					range.ParagraphFormat.SpaceAfter = 0f;
					try
					{
						Border border = range.ParagraphFormat.Borders[WdBorderType.wdBorderTop];
						border.LineStyle = WdLineStyle.wdLineStyleSingle;
						border.LineWidth = WdLineWidth.wdLineWidth050pt;
						border.Color = WdColor.wdColorAutomatic;
						return;
					}
					catch
					{
						return;
					}
				}
			}
			catch
			{
			}
			if (sel != null)
			{
				Range range2 = sel.Range;
				range2.Text = text + "\r\n";
				range2.Font.Name = "Times New Roman";
				range2.Font.Size = 11f;
				range2.Font.Bold = 0;
				range2.Font.Italic = 0;
				range2.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				range2.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				range2.ParagraphFormat.SpaceBefore = 4f;
				range2.ParagraphFormat.SpaceAfter = 0f;
				try
				{
					Border border2 = range2.ParagraphFormat.Borders[WdBorderType.wdBorderTop];
					border2.LineStyle = WdLineStyle.wdLineStyleSingle;
					border2.LineWidth = WdLineWidth.wdLineWidth050pt;
					border2.Color = WdColor.wdColorAutomatic;
					return;
				}
				catch
				{
					return;
				}
			}
		}
	}

	public static bool OpenTemplate(Application app, string templateFileName, out string errorMessage)
	{
		errorMessage = null;
		if (app == null)
		{
			errorMessage = "Ứng dụng Word chưa sẵn sàng.";
			return false;
		}
		string text = null;
		try
		{
			if (!string.IsNullOrEmpty(templateFileName) && Path.IsPathRooted(templateFileName) && File.Exists(templateFileName))
			{
				text = Path.GetFullPath(templateFileName);
			}
		}
		catch
		{
		}
		if (string.IsNullOrEmpty(text))
		{
			string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			string path = Path.Combine(Constants.AppDataDirectory, "user_templates");
			string[] array = new string[8]
			{
				Path.Combine(path, templateFileName),
				Path.Combine(directoryName, "templates", templateFileName),
				Path.Combine(Constants.ProgramDirectory, "templates", templateFileName),
				Path.Combine(Constants.AppDataDirectory, "templates", templateFileName),
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates", templateFileName),
				Path.Combine(directoryName, "..", "templates", templateFileName),
				Path.Combine(directoryName, "..", "..", "templates", templateFileName),
				Path.Combine("d:\\Tool\\V-Format\\templates", templateFileName)
			};
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				try
				{
					if (!string.IsNullOrEmpty(text2) && File.Exists(text2))
					{
						text = Path.GetFullPath(text2);
						break;
					}
				}
				catch
				{
				}
			}
		}
		if (!string.IsNullOrEmpty(text) && File.Exists(text))
		{
			try
			{
				Documents documents = app.Documents;
				object Template = text;
				object NewTemplate = Missing.Value;
				object DocumentType = Missing.Value;
				object Visible = Missing.Value;
				documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
				return true;
			}
			catch (Exception ex)
			{
				errorMessage = "Không thể mở biểu mẫu: " + ex.Message;
				return false;
			}
		}
		errorMessage = $"Không tìm thấy tệp biểu mẫu '{templateFileName}' trong thư mục templates.";
		return false;
	}

	private static bool HasNonAuditReply(Comment comment)
	{
		Comments comments = null;
		try
		{
			try
			{
				comments = comment.GetType().InvokeMember("Replies", BindingFlags.GetProperty, null, comment, null) as Comments;
			}
			catch (TargetInvocationException ex)
			{
				if (ex.InnerException is COMException ex2 && (ex2.ErrorCode == -2147352573 || ex2.ErrorCode == -2147352570))
				{
					return false;
				}
				throw;
			}
			catch (COMException ex3)
			{
				if (ex3.ErrorCode == -2147352573 || ex3.ErrorCode == -2147352570)
				{
					return false;
				}
				throw;
			}
			catch (MissingMemberException)
			{
				return false;
			}
			if (comments == null)
			{
				throw new InvalidOperationException("Không đọc được phần trả lời; giữ lại nhận xét để tránh mất nội dung.");
			}
			int count = comments.Count;
			for (int i = 1; i <= count; i++)
			{
				Comment comment2 = null;
				Range range = null;
				try
				{
					comment2 = comments[i];
					range = comment2.Range;
					if (!(range.Text ?? "").StartsWith("[OFFICEFORMAT] ", StringComparison.OrdinalIgnoreCase) && !(range.Text ?? "").StartsWith("[V-FORMAT] ", StringComparison.OrdinalIgnoreCase))
					{
						return true;
					}
				}
				finally
				{
					if (range != null)
					{
						Marshal.ReleaseComObject(range);
					}
					if (comment2 != null)
					{
						Marshal.ReleaseComObject(comment2);
					}
				}
			}
			return false;
		}
		finally
		{
			if (comments != null)
			{
				Marshal.ReleaseComObject(comments);
			}
		}
	}

	public static int DeleteDocumentComments(Document doc, bool onlyVFormatAudit, out int failed)
	{
		int keptWithReplies;
		return DeleteDocumentComments(doc, onlyVFormatAudit, out failed, out keptWithReplies);
	}

	public static int DeleteDocumentComments(Document doc, bool onlyVFormatAudit, out int failed, out int keptWithReplies)
	{
		failed = 0;
		keptWithReplies = 0;
		if (doc == null)
		{
			throw new ArgumentNullException("doc");
		}
		if (doc.ProtectionType != WdProtectionType.wdNoProtection)
		{
			throw new InvalidOperationException("Tài liệu đang được bảo vệ. Hãy bỏ bảo vệ tài liệu trước khi xóa nhận xét.");
		}
		int num = 0;
		Comments comments = null;
		try
		{
			comments = doc.Comments;
			int count = comments.Count;
			if (count == 0)
			{
				return 0;
			}
			using (new SafeUndoScope(doc.Application, "OfficeFormat Xóa Nhận Xét"))
			{
				OperationProgress.Expect(count);
				for (int num2 = count; num2 >= 1; num2--)
				{
					OperationProgress.Step("Đang xóa nhận xét...");
					Comment comment = null;
					Range range = null;
					try
					{
						comment = comments[num2];
						if (!onlyVFormatAudit)
						{
							goto IL_00b9;
						}
						range = comment.Range;
						if ((range.Text ?? "").StartsWith("[OFFICEFORMAT] ", StringComparison.OrdinalIgnoreCase) || (range.Text ?? "").StartsWith("[V-FORMAT] ", StringComparison.OrdinalIgnoreCase))
						{
							if (!HasNonAuditReply(comment))
							{
								goto IL_00b9;
							}
							keptWithReplies++;
						}
						goto end_IL_0074;
						IL_00b9:
						comment.Delete();
						num++;
						end_IL_0074:;
					}
					catch
					{
						failed++;
					}
					finally
					{
						if (range != null)
						{
							Marshal.ReleaseComObject(range);
						}
						if (comment != null)
						{
							Marshal.ReleaseComObject(comment);
						}
					}
				}
			}
			return num;
		}
		finally
		{
			if (comments != null)
			{
				Marshal.ReleaseComObject(comments);
			}
		}
	}

	public static string AuditDocument(Document doc, RulePackType rulePack)
	{
		int issueCount;
		int incompleteCount;
		return AuditDocument(doc, rulePack, out issueCount, out incompleteCount);
	}

	public static string AuditDocument(Document doc, RulePackType rulePack, out int issueCount, out int incompleteCount, UserPersonalization personalization = null)
	{
		issueCount = (incompleteCount = 0);
		if (doc == null)
		{
			incompleteCount = 1;
			return "Không có tài liệu nào đang mở.";
		}
		RuleEngine ruleEngine = new RuleEngine(ND30Rules.CreateProfile(), personalization ?? new JsonProfileRepository().LoadPersonalization());
		ruleEngine.SetActiveProfile(rulePack);
		ProfileModel activeProfile = ruleEngine.ActiveProfile;
		StringBuilder sb = new StringBuilder();
		sb.AppendLine("=== BÁO CÁO RÀ SOÁT THỂ THỨC OFFICEFORMAT ===");
		sb.AppendLine("Quy chuẩn kiểm tra: " + RulePackTraits.GetSourceTitle(rulePack));
		sb.AppendLine("Thời gian rà soát: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
		sb.AppendLine("--------------------------------------------------");
		int passCount = 0;
		int warningCount = 0;
		int commentsAdded = 0;
		int unreadable = 0;
		int commentFailures = 0;
		int num = 0;
		int num2 = 0;
		sb.AppendLine("Đối chiếu lề và kiểu chữ với cấu hình hiệu lực của quy chuẩn đang chọn (bao gồm cá nhân hóa).");
		sb.AppendLine("Phạm vi: khổ giấy/lề từng section, font các đoạn trong thân tài liệu và các dấu hiệu thể thức được nhận diện. Không kết luận thay cho kiểm định pháp lý đầy đủ.");
		HashSet<string> existingComments = new HashSet<string>(StringComparer.Ordinal);
		try
		{
			foreach (Comment comment in doc.Comments)
			{
				try
				{
					existingComments.Add(comment.Scope.Start + "|" + (comment.Range.Text ?? "").TrimEnd('\r', '\a'));
				}
				finally
				{
					Marshal.ReleaseComObject(comment);
				}
			}
		}
		catch
		{
			unreadable++;
		}
		Action<Range, string> action = (Range rng, string msg) =>
		{
			try
			{
				sb.AppendLine(" [CẦN KIỂM TRA] Vị trí " + rng.Start + ": " + msg);
				string text6 = "[OFFICEFORMAT] " + msg;
				string item = rng.Start + "|" + text6;
				if (!existingComments.Contains(item))
				{
					Comments comments = rng.Document.Comments;
					object Text = text6;
					Comment o = comments.Add(rng, ref Text);
					Marshal.ReleaseComObject(o);
					existingComments.Add(item);
					commentsAdded++;
				}
			}
			catch
			{
				commentFailures++;
			}
		};
		for (int num3 = 1; num3 <= doc.Sections.Count; num3++)
		{
			OperationProgress.Step();
			Section comObj = null;
			PageSetup comObj2 = null;
			Range comObj3 = null;
			try
			{
				comObj = doc.Sections[num3];
				comObj2 = comObj.PageSetup;
				comObj3 = comObj.Range.Duplicate;
				Range range = comObj3;
				object Direction = WdCollapseDirection.wdCollapseStart;
				range.Collapse(ref Direction);
				bool flag = Math.Abs(Math.Min(comObj2.PageWidth, comObj2.PageHeight) - UnitConverter.MmToPt(210f)) < 5f && Math.Abs(Math.Max(comObj2.PageWidth, comObj2.PageHeight) - UnitConverter.MmToPt(297f)) < 5f;
				bool flag2 = Math.Abs(Math.Min(comObj2.PageWidth, comObj2.PageHeight) - UnitConverter.MmToPt(148f)) < 5f && Math.Abs(Math.Max(comObj2.PageWidth, comObj2.PageHeight) - UnitConverter.MmToPt(210f)) < 5f;
				bool flag3 = rulePack switch
				{
					RulePackType.HD05_VPTW => string.Equals(activeProfile.PageSetup.PaperSizeName, "A5", StringComparison.OrdinalIgnoreCase), 
					RulePackType.HD29_TWDTN => true, 
					_ => false, 
				};
				if (flag || (flag2 && flag3))
				{
					passCount++;
				}
				else
				{
					warningCount++;
					action(comObj3, "Section " + num3 + ": khổ giấy chưa khớp cấu hình A4/A5 áp dụng.");
				}
				LayoutPreset pageSetup = activeProfile.PageSetup;
				bool flag4 = Math.Abs(comObj2.TopMargin - pageSetup.MarginTopPt) <= UnitConverter.MmToPt(2f) && Math.Abs(comObj2.BottomMargin - pageSetup.MarginBottomPt) <= UnitConverter.MmToPt(2f) && Math.Abs(comObj2.LeftMargin - pageSetup.MarginLeftPt) <= UnitConverter.MmToPt(2f) && Math.Abs(comObj2.RightMargin - pageSetup.MarginRightPt) <= UnitConverter.MmToPt(2f);
				string text = $"Section {num3}: lề trên/dưới/trái/phải {UnitConverter.PtToMm(comObj2.TopMargin):F1}/{UnitConverter.PtToMm(comObj2.BottomMargin):F1}/{UnitConverter.PtToMm(comObj2.LeftMargin):F1}/{UnitConverter.PtToMm(comObj2.RightMargin):F1} mm; cấu hình hiệu lực {UnitConverter.PtToMm(pageSetup.MarginTopPt):F1}/{UnitConverter.PtToMm(pageSetup.MarginBottomPt):F1}/{UnitConverter.PtToMm(pageSetup.MarginLeftPt):F1}/{UnitConverter.PtToMm(pageSetup.MarginRightPt):F1} mm.";
				if (flag4)
				{
					passCount++;
					sb.AppendLine(" [ĐẠT] " + text);
				}
				else
				{
					warningCount++;
					action(comObj3, text);
				}
				num2++;
			}
			catch (OperationCanceledException)
			{
				throw;
			}
			catch
			{
				unreadable++;
			}
			finally
			{
				ComReleaser.Release(ref comObj3);
				ComReleaser.Release(ref comObj2);
				ComReleaser.Release(ref comObj);
			}
		}
		try
		{
			foreach (Table table in doc.Tables)
			{
				OperationProgress.Step();
				for (int num4 = 1; num4 <= table.Rows.Count; num4++)
				{
					for (int num5 = 1; num5 <= table.Columns.Count; num5++)
					{
						try
						{
							Cell cell = table.Cell(num4, num5);
							string text2 = (cell.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
							string text3 = RemoveVietnameseDiacritics(text2).ToUpperInvariant();
							if ((text3.Contains("CONG HOA XA HOI") || text3.Contains("DANG CONG SAN") || text3.Contains("DOAN TNCS") || text3.Contains("DOAN THANH NIEN CONG SAN")) && cell.Range.Font.Bold == 0 && activeProfile.GetComponentStyle(SemanticRole.NationalHeader).Bold)
							{
								warningCount++;
								action(cell.Range, "Quốc hiệu / Tiêu đề Đảng, Đoàn phải in đậm (Bold), in hoa toàn bộ.");
							}
							if ((text3.Contains("DOC LAP - TU DO") || text3.Contains("DOC LAP")) && cell.Range.Font.Bold == 0 && activeProfile.GetComponentStyle(SemanticRole.Motto).Bold)
							{
								warningCount++;
								action(cell.Range, "Tiêu ngữ phải in đậm, chữ cái đầu viết hoa, có gạch nối cách đều (Độc lập - Tự do - Hạnh phúc).");
							}
							if (text3.Contains("SO :"))
							{
								warningCount++;
								action(cell.Range, "Không để dấu cách trước dấu hai chấm ở số văn bản (viết đúng: 'Số:').");
							}
							if (text3.Contains("NGAY") && text3.Contains("THANG") && text3.Contains("NAM") && cell.Range.Font.Italic == 0 && activeProfile.GetComponentStyle(SemanticRole.DocumentDate).Italic)
							{
								warningCount++;
								action(cell.Range, "Địa danh và ngày tháng năm phải in nghiêng (Italic), chữ thường, cỡ 13-14pt.");
							}
						}
						catch (OperationCanceledException)
						{
							throw;
						}
						catch
						{
							unreadable++;
						}
					}
				}
			}
		}
		catch (OperationCanceledException)
		{
			throw;
		}
		catch
		{
			unreadable++;
		}
		try
		{
			int count = doc.Paragraphs.Count;
			for (int num6 = 1; num6 <= count; num6++)
			{
				OperationProgress.Step();
				try
				{
					Paragraph paragraph = doc.Paragraphs[num6];
					string text4 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					string text5 = RemoveVietnameseDiacritics(text4).ToUpperInvariant();
					num++;
					if (text5.Length > 2)
					{
						if (!AuditFontMatches(paragraph.Range, activeProfile.Typography.FontName))
						{
							warningCount++;
							action(paragraph.Range, $"Đoạn {num6}: phông chữ chưa khớp cấu hình {activeProfile.Typography.FontName}.");
						}
						if ((text5 == "QUYET DINH" || text5.StartsWith("QUYET DINH ") || text5 == "NGHI QUYET" || text5.StartsWith("NGHI QUYET ") || text5 == "TO TRINH" || text5.StartsWith("TO TRINH ") || text5 == "KE HOACH" || text5.StartsWith("KE HOACH ")) && paragraph.Range.Font.Bold == 0 && activeProfile.GetComponentStyle(SemanticRole.DocumentType).Bold)
						{
							warningCount++;
							action(paragraph.Range, "Tên loại văn bản phải in đậm (Bold), căn giữa.");
						}
						if (rulePack == RulePackType.HD05_VPTW && text5.StartsWith("NOI NHAN") && (paragraph.Range.Font.Bold != 0 || paragraph.Range.Font.Italic != 0 || paragraph.Range.Font.Underline == WdUnderline.wdUnderlineNone))
						{
							warningCount++;
							action(paragraph.Range, "Theo HD 05-HD/VPTW, từ 'Nơi nhận:' phải gạch chân và không in đậm, không in nghiêng.");
						}
						else if (rulePack != RulePackType.HD29_TWDTN && rulePack != RulePackType.HD05_VPTW && text5.StartsWith("NOI NHAN") && ((paragraph.Range.Font.Italic == 0 && activeProfile.GetComponentStyle(SemanticRole.RecipientHeader).Italic) || (paragraph.Range.Font.Bold == 0 && activeProfile.GetComponentStyle(SemanticRole.RecipientHeader).Bold)))
						{
							warningCount++;
							action(paragraph.Range, "Từ 'Nơi nhận:' phải in đậm và nghiêng (Italic).");
						}
					}
				}
				catch (OperationCanceledException)
				{
					throw;
				}
				catch
				{
					unreadable++;
				}
			}
		}
		catch (OperationCanceledException)
		{
			throw;
		}
		catch
		{
			unreadable++;
		}
		if (rulePack == RulePackType.HD29_TWDTN)
		{
			try
			{
				sb.AppendLine("Đối chiếu bổ sung HD 29 gốc: các cảnh báo dưới đây có thể khác lựa chọn cá nhân hóa.");
				AuditYouthUnionDocument(doc, sb, action, ref passCount, ref warningCount, ref unreadable);
			}
			catch (OperationCanceledException)
			{
				throw;
			}
			catch
			{
				unreadable++;
			}
		}
		sb.AppendLine("--------------------------------------------------");
		sb.AppendLine($"Tổng kết: {passCount} tiêu chuẩn ĐẠT, {warningCount} lỗi/cảnh báo được phát hiện.");
		if (commentsAdded > 0)
		{
			sb.AppendLine($"Đã đính kèm {commentsAdded} ghi chú (Word Comments) trực tiếp vào các dòng bị lỗi thể thức trên văn bản!");
		}
		if (warningCount > 0)
		{
			sb.AppendLine("Gợi ý: dùng các lệnh chuẩn hóa phù hợp rồi rà soát lại; nội dung và các trường hợp đặc biệt cần kiểm tra thủ công.");
		}
		else
		{
			sb.AppendLine((unreadable == 0) ? "Không phát hiện lỗi trong các mục đã kiểm tra." : "Chưa thể kết luận: còn mục không đọc được.");
		}
		sb.AppendLine($"Phạm vi đã đọc: {num2} section, {num} đoạn thân tài liệu. Mục không đọc được: {unreadable}. Ghi chú không tạo được: {commentFailures}.");
		sb.AppendLine("Chưa rà soát đầy đủ nội dung đầu/chân trang, hộp văn bản và mọi yêu cầu thể thức. Các ghi chú cũ được giữ lại; kết quả lượt này là số liệu trong báo cáo.");
		issueCount = warningCount;
		incompleteCount = unreadable;
		return sb.ToString();
	}

	private static bool AuditFontMatches(Range range, string expected)
	{
		string name = range.Font.Name;
		if (!string.IsNullOrEmpty(name))
		{
			return string.Equals(name, expected, StringComparison.OrdinalIgnoreCase);
		}
		for (int i = 1; i <= range.Characters.Count; i++)
		{
			if (i % 128 == 0)
			{
				OperationProgress.Step();
			}
			Range range2 = range.Characters[i];
			try
			{
				string text = range2.Text;
				if (!string.IsNullOrWhiteSpace(text) && text != "\a" && !string.Equals(range2.Font.Name, expected, StringComparison.OrdinalIgnoreCase))
				{
					return false;
				}
			}
			finally
			{
				Marshal.ReleaseComObject(range2);
			}
		}
		return true;
	}

	private static void AuditYouthUnionDocument(Document doc, StringBuilder sb, Action<Range, string> addComment, ref int passCount, ref int warningCount, ref int unreadable)
	{
		string text = "";
		Range range = null;
		try
		{
			if (doc.Tables.Count > 0 && doc.Tables[1].Range.Start < 1500)
			{
				range = doc.Tables[1].Range;
				text = range.Text ?? "";
			}
			else
			{
				int num = ((doc.Paragraphs.Count >= 15) ? doc.Paragraphs[15].Range.End : doc.Content.End);
				object Start = 0;
				object End = num;
				range = doc.Range(ref Start, ref End);
				text = range.Text ?? "";
			}
		}
		catch (OperationCanceledException)
		{
			throw;
		}
		catch
		{
			unreadable++;
		}
		string text2 = RemoveVietnameseDiacritics(text).ToUpperInvariant();
		Range range2 = range ?? ((doc.Paragraphs.Count > 0) ? doc.Paragraphs[1].Range : null);
		if (RegexPatterns.YouthUnionHeader.IsMatch(text))
		{
			sb.AppendLine(" [ĐẠT] DTN-01 Tiêu đề: Có \"ĐOÀN TNCS HỒ CHÍ MINH\" ở khối đầu.");
			passCount++;
		}
		else
		{
			sb.AppendLine(" [LỖI] DTN-01 Tiêu đề: Thiếu tiêu đề \"ĐOÀN TNCS HỒ CHÍ MINH\" (góc phải, dòng đầu, có kẻ ngang dài bằng tiêu đề).");
			warningCount++;
			if (range2 != null)
			{
				addComment(range2, "HD 29 mục III.1: Văn bản của Đoàn phải có tiêu đề \"ĐOÀN TNCS HỒ CHÍ MINH\" ở góc phải, dòng đầu.");
			}
		}
		if (text2.Contains("CONG HOA XA HOI") || text2.Contains("DANG CONG SAN VIET NAM"))
		{
			sb.AppendLine(" [LỖI] DTN-02 Khối đầu: Có Quốc hiệu/Tiêu ngữ hoặc tiêu đề Đảng - không đúng thể thức Đoàn (trừ văn bản liên tịch, đơn vị dùng dấu Nhà nước).");
			warningCount++;
			if (range2 != null)
			{
				addComment(range2, "HD 29: Văn bản của Đoàn dùng tiêu đề \"ĐOÀN TNCS HỒ CHÍ MINH\", không dùng Quốc hiệu hoặc \"ĐẢNG CỘNG SẢN VIỆT NAM\" (trừ trường hợp mục II.10).");
			}
		}
		else
		{
			passCount++;
		}
		bool flag = false;
		string[] array = text.Split(new char[4] { '\r', '\n', '\a', '\v' }, StringSplitOptions.RemoveEmptyEntries);
		foreach (string text3 in array)
		{
			if (text3.Replace(" ", "").Trim() == "***")
			{
				flag = true;
				break;
			}
		}
		if (flag)
		{
			sb.AppendLine(" [ĐẠT] DTN-03 Dấu phân cách: Có \"***\" dưới tên cơ quan ban hành.");
			passCount++;
		}
		else
		{
			sb.AppendLine(" [CẢNH BÁO] DTN-03 Dấu phân cách: Chưa có \"***\" dưới tên cơ quan ban hành (HD 29 mục III.3.1).");
			warningCount++;
			if (range2 != null)
			{
				addComment(range2, "HD 29 mục III.3.1: Số văn bản phân cách với tên cơ quan ban hành bởi 3 dấu sao (***), không kẻ đường ngang.");
			}
		}
		Match match = Regex.Match(text, "S[á»‘o]\\s*:?\\s*[^\\r\\n\\a\\v]*", RegexOptions.IgnoreCase);
		if (match.Success)
		{
			string text4 = match.Value.Trim();
			bool flag2 = Regex.IsMatch(text4, "^S[ốo]\\s*:?\\s*(\\d{1,4}|\\s*)\\s*-\\s*[A-Za-zĐđ]+\\s*/\\s*[A-ZĐ0-9\\-]+", RegexOptions.IgnoreCase) || Regex.IsMatch(text4, "/\\s*ĐH\\s*$", RegexOptions.IgnoreCase);
			bool flag3 = Regex.IsMatch(text4, "^S[á»‘o]\\s*:", RegexOptions.IgnoreCase);
			if (flag2 && flag3)
			{
				sb.AppendLine(" [ĐẠT] DTN-04 Số ký hiệu: \"" + text4 + "\".");
				passCount++;
			}
			else
			{
				sb.AppendLine(" [CẢNH BÁO] DTN-04 Số ký hiệu: \"" + text4 + "\" - nên viết dạng \"Số: 245-CV/TWĐTN-VP\" (gạch nối giữa số và ký hiệu, gạch chéo giữa loại và cơ quan).");
				warningCount++;
				if (range2 != null)
				{
					addComment(range2, "HD 29 mục III.3.2: Giữa số và ký hiệu có dấu nối (-), giữa tên loại và cơ quan ban hành có dấu gạch chéo (/), giữa cơ quan ban hành và đơn vị tham mưu có dấu nối (-). Ví dụ: Số: 245-CV/TWĐTN-VP.");
				}
			}
		}
		Match match2 = Regex.Match(text, "ngày\\s+(\\d{1,2})\\s+tháng\\s+(\\d{1,2})\\s+năm\\s+\\d{4}", RegexOptions.IgnoreCase);
		if (match2.Success)
		{
			bool flag4 = int.TryParse(match2.Groups[1].Value, out var result) && result < 10 && match2.Groups[1].Value.Length == 1;
			bool flag5 = int.TryParse(match2.Groups[2].Value, out var result2) && result2 <= 2 && match2.Groups[2].Value.Length == 1;
			if (flag4 || flag5)
			{
				sb.AppendLine(" [CẢNH BÁO] DTN-05 Ngày tháng: Ngày dưới 10 và tháng 1, 2 phải thêm số 0 phía trước (ví dụ: ngày 09 tháng 02 năm 2013).");
				warningCount++;
				if (range2 != null)
				{
					addComment(range2, "HD 29 mục III.4.2: Ngày dưới 10 và tháng 1, 2 phải ghi thêm số 0 đứng trước.");
				}
			}
			else
			{
				passCount++;
			}
		}
		int num2 = 0;
		int num3 = 0;
		Range range3 = null;
		int count = doc.Paragraphs.Count;
		for (int j = 1; j <= count; j++)
		{
			OperationProgress.Step();
			Paragraph paragraph = null;
			try
			{
				paragraph = doc.Paragraphs[j];
				string text5 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				if (text5.Length == 0)
				{
					continue;
				}
				string text6 = RemoveVietnameseDiacritics(text5).ToUpperInvariant();
				if (Regex.IsMatch(text5, "^\\s*(T/M|K/T|T/L)\\b"))
				{
					warningCount++;
					sb.AppendLine(" [CẢNH BÁO] DTN-08 Thể thức đề ký: \"" + text5 + "\" - văn bản Đoàn dùng \"TM.\", \"KT.\", \"TL.\".");
					addComment(paragraph.Range, "HD 29 mục III.7.2: Thể thức đề ký của Đoàn là TM. (thay mặt), KT. (ký thay), TL. (thừa lệnh), không viết T/M, K/T, T/L.");
				}
				if (text6.StartsWith("NOI NHAN"))
				{
					if (paragraph.Range.Font.Bold == 0 || paragraph.Range.Font.Italic == -1 || (paragraph.Range.Font.Size > 0f && Math.Abs(paragraph.Range.Font.Size - 13f) > 0.5f))
					{
						warningCount++;
						sb.AppendLine(" [CẢNH BÁO] DTN-09 Nơi nhận: Từ \"Nơi nhận:\" phải đứng, đậm, cỡ 13.");
						addComment(paragraph.Range, "HD 29 mục III.8 / Mẫu 3: \"Nơi nhận:\" in thường, đứng, đậm, cỡ 13; danh sách nơi nhận cỡ 11, đứng.");
					}
					else
					{
						passCount++;
					}
				}
				if (paragraph.Range.Tables.Count > 0 || text5.Length < 60 || paragraph.Format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
				{
					continue;
				}
				if (paragraph.Format.SpaceAfter > 6.1f || paragraph.Format.SpaceBefore > 6.1f)
				{
					num2++;
					if (range3 == null)
					{
						range3 = paragraph.Range;
					}
				}
				if (paragraph.Format.LineSpacingRule == WdLineSpacing.wdLineSpaceMultiple || paragraph.Format.LineSpacingRule == WdLineSpacing.wdLineSpace1pt5 || paragraph.Format.LineSpacingRule == WdLineSpacing.wdLineSpaceDouble)
				{
					float num4 = paragraph.Format.LineSpacing / 12f;
					if (num4 > 1.21f)
					{
						num3++;
					}
				}
			}
			catch (OperationCanceledException)
			{
				throw;
			}
			catch
			{
				unreadable++;
			}
		}
		if (num2 > 0)
		{
			warningCount++;
			sb.AppendLine($" [CẢNH BÁO] DTN-07 Cách đoạn: {num2} đoạn có khoảng cách đoạn lớn hơn 6 pt (HD 29 mục III.6: tối đa 6 pt).");
			if (range3 != null)
			{
				addComment(range3, "HD 29 mục III.6: Khoảng cách giữa các đoạn văn đặt tối đa là 6 pt.");
			}
		}
		if (num3 > 0)
		{
			warningCount++;
			sb.AppendLine($" [CẢNH BÁO] DTN-07 Cách dòng: {num3} đoạn có cách dòng lớn hơn 1,2 (HD 29 mục III.6: từ single đến tối đa 1,2).");
		}
		if (num2 == 0 && num3 == 0)
		{
			sb.AppendLine(" [ĐẠT] DTN-07 Cách đoạn tối đa 6 pt, cách dòng 1,0 - 1,2.");
			passCount++;
		}
		try
		{
			if (doc.Sections.Count <= 0)
			{
				return;
			}
			object IncludeFootnotesAndEndnotes = Missing.Value;
			if (doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref IncludeFootnotesAndEndnotes) > 1)
			{
				if (doc.Sections[1].PageSetup.DifferentFirstPageHeaderFooter == -1)
				{
					sb.AppendLine(" [ĐẠT] DTN-10 Số trang: Không đánh số trang thứ nhất.");
					passCount++;
				}
				else
				{
					warningCount++;
					sb.AppendLine(" [CẢNH BÁO] DTN-10 Số trang: Văn bản nhiều trang đánh số từ trang 2, không đánh số trang thứ nhất (HD 29 mục II.4).");
				}
			}
		}
		catch (OperationCanceledException)
		{
			throw;
		}
		catch
		{
			unreadable++;
		}
	}
}
