using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Interfaces;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;
using VFormat.Core.Serialization;
using VFormat.WordEngine.Builders;
using VFormat.WordEngine.Formatters;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine;

public class WordFormatEngine
{
	private enum ComponentOutcome
	{
		Formatted,
		Rebuilt,
		Converted,
		InsertedTemplate,
		AppendixEnding,
		KeptRawBlock,
		Missing
	}

	private sealed class SymbolFontScope : IDisposable
	{
		private readonly Document _doc;

		private List<Tuple<Range, string>> _runs;

		public SymbolFontScope(Document doc)
		{
			_doc = doc;
			_runs = CaptureSymbolFontRuns(doc);
		}

		public void Dispose()
		{
			if (_runs == null)
			{
				return;
			}
			Restore();
			foreach (Tuple<Range, string> run in _runs)
			{
				Range comObj = run.Item1;
				ComReleaser.Release(ref comObj);
			}
			_runs = null;
		}

		public void Restore()
		{
			if (_runs != null)
			{
				RestoreSymbolFontRuns(_doc, _runs);
			}
		}
	}

	private readonly IRuleEngine _ruleEngine;

	private static readonly Regex YouthMultiLevelNumber = new Regex("^\\s*(\\d{1,2}(?:\\.\\d{1,2}){1,5})\\.?(?=\\s)", RegexOptions.Compiled);

	private static readonly Regex LeadInEnding = new Regex("(như\\s+sau|dưới\\s+đây|sau\\s+đây|cụ\\s+thể(\\s+là|\\s+như\\s+sau)?|bao\\s+gồm|gồm\\s+có|gồm|như\\s+bảng|theo\\s+bảng|tại\\s+bảng|kèm\\s+theo)\\s*:\\s*$", RegexOptions.IgnoreCase);

	private static readonly Regex MissingSpaceAfterMarker = new Regex("^(?:[a-zđA-ZĐ][\\)\\.](?=\\p{L})|\\d{1,2}(?:\\.\\d{1,2})*\\.(?=\\p{L})|[IVXLCDM]{1,6}\\.(?=\\p{L}))", RegexOptions.Compiled);

	private static readonly Regex HeadingOrListMarkerRun = new Regex("^(?<marker>(?:[IVXLCDM]{1,6}|[A-Za-zĐđ]|\\d{1,2}(?:\\.\\d{1,2})*)\\.|[a-zđ]\\)|[-+])[ \\t\\u00A0]+(?=\\S)", RegexOptions.Compiled);

	private static readonly Regex MultiLevelNoDotMarkerRun = new Regex("^(?<marker>\\d{1,2}(?:\\.\\d{1,2}){1,5})[ \\t\\u00A0]+(?=\\S)", RegexOptions.Compiled);

	private static readonly Regex MultiSpaceRun = new Regex("(?<=\\S)[ \\u00A0]{2,}(?=\\S)");

	private static readonly Regex SpaceBeforePunct = new Regex("(?<=\\S)[ \\u00A0]+(?=[,;:.)\\]?!])");

	private static readonly Regex SpaceAfterOpenParen = new Regex("(?<=[(\\[])[ \\u00A0]+(?=\\S)");

	internal static readonly string[] SymbolFontNames = new string[7] { "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings", "Marlett", "MT Extra" };

	public IRuleEngine RuleEngine => _ruleEngine;

	public event Action<int, int, string> ProgressChanged;

	public WordFormatEngine()
	{
		_ruleEngine = new RuleEngine();
	}

	public WordFormatEngine(IRuleEngine ruleEngine)
	{
		_ruleEngine = ruleEngine ?? new RuleEngine();
	}

	private void ReportProgress(int current, int total, string message)
	{
		OperationProgress.Report(current, total, message);
		if (ProgressChanged != null)
		{
			try
			{
				ProgressChanged(current, total, message);
			}
			catch
			{
			}
		}
	}

	public FormattingResult StandardizeDocument(Document doc, RulePackType? forcedRulePack)
	{
		return StandardizeDocument(doc, forcedRulePack, null);
	}

	public FormattingResult StandardizeDocument(Document doc, RulePackType? forcedRulePack = null, PrintLayoutMode? printLayoutMode = null)
	{
		return StandardizeDocument(doc, forcedRulePack, printLayoutMode, safeMode: false);
	}

	public FormattingResult StandardizeDocument(Document doc, RulePackType? forcedRulePack, PrintLayoutMode? printLayoutMode, bool safeMode)
	{
		return StandardizeDocument(doc, forcedRulePack, printLayoutMode, safeMode, null);
	}

	public FormattingResult StandardizeDocument(Document doc, RulePackType? forcedRulePack, PrintLayoutMode? printLayoutMode, bool safeMode, StandardizeOptions options)
	{
		StandardizeOptions standardizeOptions = options ?? StandardizeOptions.All;
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc", "Tài liệu Word không được để trống."));
		}
		FormattingResult formattingResult = CheckLegacyEncoding(doc, forcedRulePack);
		if (formattingResult != null)
		{
			return formattingResult;
		}
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application comObj = null;
		try
		{
			comObj = doc.Application;
			EnsureEditablePrintView(doc);
			using (new ScreenFreezeScope(comObj))
			{
				using (new SafeUndoScope(comObj))
				{
					using SymbolFontScope symbolFontScope = new SymbolFontScope(doc);
					ReportProgress(0, 100, "Đang phân tích cấu trúc tài liệu...");
					RulePackType rulePackType;
					if (forcedRulePack.HasValue)
					{
						rulePackType = forcedRulePack.Value;
					}
					else
					{
						string documentSampleText = GetDocumentSampleText(doc);
						rulePackType = _ruleEngine.DetectRulePack(documentSampleText);
					}
					_ruleEngine.SetActiveProfile(rulePackType);
					ProfileModel activeProfile = _ruleEngine.ActiveProfile;
					if (activeProfile != null && activeProfile.PageSetup != null && printLayoutMode.HasValue)
					{
						activeProfile.PageSetup.MirrorMargins = printLayoutMode.Value == PrintLayoutMode.TwoSidedMirror;
					}
					ReportProgress(5, 100, "Đang thiết lập phông chữ chuẩn Times New Roman...");
					if (standardizeOptions.Typography)
					{
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
						try
						{
							doc.Content.Font.ColorIndex = WdColorIndex.wdAuto;
							doc.Content.HighlightColorIndex = WdColorIndex.wdAuto;
						}
						catch
						{
						}
						symbolFontScope.Restore();
					}
					ReportProgress(15, 100, "Đang chuẩn hóa lề trang và khổ giấy A4...");
					if (standardizeOptions.PageSetup)
					{
						PageSetupFormatter.ApplyPageSetup(doc, activeProfile);
					}
					ReportProgress(25, 100, "Đang chuẩn hóa đánh số trang...");
					if (standardizeOptions.PageNumbers)
					{
						HeaderFooterFormatter.ApplyPageNumbers(doc, activeProfile.PageNumber, preserveUserContent: true);
					}
					ReportProgress(35, 100, "Đang chuẩn hóa khối Quốc hiệu & Tiêu đề...");
					ComponentOutcome outcome = ((standardizeOptions.Typography && standardizeOptions.Tables) ? FormatHeaderComponent(doc, rulePackType, activeProfile, safeMode, !safeMode) : ComponentOutcome.Missing);
					ReportProgress(40, 100, DescribeOutcome("Tiêu đề", outcome));
					ReportProgress(45, 100, "Đang chuẩn hóa Tên loại văn bản & Trích yếu...");
					if (standardizeOptions.Typography)
					{
						StandardizeTitleAndAbstract(doc, rulePackType, preserveLayout: true);
					}
					ReportProgress(50, 100, "Đang chuẩn hóa font chữ, cỡ chữ và giãn dòng các đoạn...");
					int paragraphs = 0;
					if (standardizeOptions.Typography)
					{
						paragraphs = FormatAllParagraphs(doc, activeProfile, (int cur, int tot) =>
						{
							int current = 50 + (int)Math.Round(22.0 * (double)cur / (double)tot);
							ReportProgress(current, 100, $"Đang chuẩn hóa đoạn văn ({cur}/{tot})...");
						});
					}
					ReportProgress(74, 100, "Đang chuẩn hóa khối Chữ ký & Nơi nhận...");
					ComponentOutcome outcome2 = ((standardizeOptions.Typography && standardizeOptions.Tables) ? FormatSignatureComponent(doc, rulePackType, safeMode, !safeMode) : ComponentOutcome.Missing);
					ReportProgress(76, 100, DescribeOutcome("Chữ ký & Nơi nhận", outcome2));
					ReportProgress(78, 100, "Đang chuẩn hóa bảng biểu nội dung...");
					int tables = 0;
					if (standardizeOptions.Tables)
					{
						TableFormatter.FormatTables(doc, activeProfile.TablePreset);
						tables = doc.Tables.Count;
					}
					ReportProgress(80, 100, "Đang dọn khoảng trắng thừa và trang trắng cuối văn bản...");
					if (standardizeOptions.Typography)
					{
						try
						{
							CleanupInlineSpacing(doc);
						}
						catch
						{
						}
					}
					if (standardizeOptions.PageSetup)
					{
						try
						{
							RemoveTrailingBlankPages(doc);
						}
						catch
						{
						}
					}
					ReportProgress(82, 100, "Bảo toàn khoảng trắng và tab dùng để căn cột...");
					ReportProgress(88, 100, "Bảo toàn nội dung, tên riêng và trích dẫn...");
					ReportProgress(98, 100, "Bảo toàn ngắt phần (section) giữa văn bản...");
					ReportProgress(99, 100, "Đang hoàn tất và đồng bộ cấu trúc văn bản...");
					if (standardizeOptions.Typography)
					{
						try
						{
							Table table = HeaderRebuilder.FindExistingHeaderTable(doc);
							if (table != null && table.Rows.Count >= 2 && table.Columns.Count >= 2)
							{
								Cell cell = table.Cell(2, 1);
								if (cell.Range.Paragraphs.Count > 1)
								{
									Paragraph paragraph = cell.Range.Paragraphs[1];
									string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
									if (HeaderRebuilder.IsDocNumberLine(text))
									{
										string text2 = HeaderRebuilder.NormalizeDocNumber(text, rulePackType);
										if (text != text2)
										{
											paragraph.Range.Text = text2 + "\r";
										}
									}
								}
								else
								{
									string text3 = (cell.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
									if (HeaderRebuilder.IsDocNumberLine(text3))
									{
										string text4 = HeaderRebuilder.NormalizeDocNumber(text3, rulePackType);
										if (text3 != text4)
										{
											cell.Range.Text = text4;
											cell.Range.Font.Name = "Times New Roman";
											cell.Range.Font.Size = GetComponentSizeOrDefault(activeProfile, SemanticRole.DocumentNumber, RulePackTraits.UsesOrganizationHeader(rulePackType) ? 14f : 13f);
											cell.Range.Font.Bold = 0;
											cell.Range.Font.Italic = 0;
											cell.Range.Font.Underline = WdUnderline.wdUnderlineNone;
										}
									}
								}
							}
						}
						catch
						{
						}
					}
					ReportProgress(99, 100, "Đang nhận diện và định dạng từng thành phần trong bảng và khối ký...");
					int componentsFormatted = (standardizeOptions.Typography ? SemanticComponentFormatter.Apply(doc, activeProfile, null, standardizeOptions.Tables) : 0);
					ReportProgress(100, 100, "Hoàn thành chuẩn hóa văn bản.");
					stopwatch.Stop();
					FormattingResult formattingResult2 = FormattingResult.Succeeded(rulePackType, paragraphs, tables, stopwatch.ElapsedMilliseconds);
					formattingResult2.ComponentsFormatted = componentsFormatted;
					formattingResult2.SafeModeApplied = safeMode;
					formattingResult2.HeaderStatus = DescribeOutcome("Tiêu đề", outcome);
					formattingResult2.SignatureStatus = DescribeOutcome("Chữ ký & Nơi nhận", outcome2);
					return formattingResult2;
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
		finally
		{
			comObj = null;
		}
	}

	public FormattingResult StandardizeDocument(Document doc, ProfileModel profile)
	{
		if (profile == null)
		{
			return StandardizeDocument(doc, (RulePackType?)null);
		}
		_ruleEngine.SetActiveProfile(profile.RulePack);
		return StandardizeDocument(doc, profile.RulePack);
	}

	public FormattingResult StandardizeSelection(Selection selection, RulePackType? forcedRulePack = null, bool includeAllRoles = true)
	{
		if (selection == null || selection.Range == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("selection", "Vùng chọn không hợp lệ."));
		}
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application comObj = null;
		try
		{
			comObj = selection.Application;
			if (selection.Range.StoryType != WdStoryType.wdMainTextStory)
			{
				return FormattingResult.Failed(new InvalidOperationException("Chỉ có thể chuẩn hóa vùng chọn trong nội dung chính của văn bản."));
			}
			using (new ScreenFreezeScope(comObj))
			{
				using (new SafeUndoScope(comObj, "OfficeFormat Chuẩn Hóa Vùng Chọn"))
				{
					RulePackType rulePackType = (forcedRulePack.HasValue ? forcedRulePack.Value : _ruleEngine.ActiveProfile.RulePack);
					try
					{
						JsonProfileRepository jsonProfileRepository = new JsonProfileRepository();
						UserPersonalization userPersonalization = jsonProfileRepository.LoadPersonalization();
						if (userPersonalization != null)
						{
							_ruleEngine.UpdatePersonalization(userPersonalization);
						}
					}
					catch
					{
					}
					_ruleEngine.SetActiveProfile(rulePackType);
					ProfileModel activeProfile = _ruleEngine.ActiveProfile;
					int tablesProcessed = 0;
					try
					{
						if (selection.Tables != null && selection.Tables.Count > 0)
						{
							Tables selTables = selection.Tables;
							int tCount = selTables.Count;
							for (int t = 1; t <= tCount; t++)
							{
								Table tbl = null;
								try
								{
									tbl = selTables[t];
									if (TableFormatter.IsDirectRecipientTable(tbl))
									{
										TableFormatter.FormatDirectRecipientTable(tbl, activeProfile);
										tablesProcessed++;
										continue;
									}
									if (!TableFormatter.IsMetadataTable(tbl))
									{
										TableFormatter.FormatSingleTable(tbl, activeProfile.TablePreset);
										tablesProcessed++;
									}
								}
								catch
								{
								}
								finally
								{
									ComReleaser.Release(ref tbl);
								}
							}
							ComReleaser.Release(ref selTables);
						}
					}
					catch
					{
					}
					Paragraphs comObj2 = selection.Paragraphs;
					int count = comObj2.Count;
					int num = 0;
					Document document = selection.Document;
					int totalParagraphs = document?.Paragraphs.Count ?? count;
					string documentType = DetectPrimaryDocumentType(document);
					bool selFlagDirectRecipient = false;
					for (int i = 1; i <= count; i++)
					{
						OperationProgress.Step();
						Paragraph comObj3 = null;
						try
						{
							comObj3 = comObj2[i];
							string text = comObj3.Range.Text;
							string trimmedSelectionText = (text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
							if (string.IsNullOrEmpty(trimmedSelectionText))
							{
								continue;
							}
							bool isInTable = comObj3.Range.Tables.Count > 0;
							if (isInTable)
							{
								continue;
							}
							NormalizeHeadingAndListMarkers(comObj3);
							try
							{
								string selInput = comObj3.Range.Text ?? "";
								Match selMatch = MissingSpaceAfterMarker.Match(selInput);
								if (selMatch.Success && comObj3.Range.Fields.Count == 0)
								{
									Range selDup = comObj3.Range.Duplicate;
									selDup.SetRange(comObj3.Range.Start + selMatch.Length, comObj3.Range.Start + selMatch.Length);
									selDup.InsertAfter(" ");
									ComReleaser.Release(ref selDup);
								}
							}
							catch
							{
							}
							text = comObj3.Range.Text;
							trimmedSelectionText = (text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
							SemanticRole semanticRole = SemanticRole.BodyText;
							if (RegexPatterns.DirectRecipient.IsMatch(trimmedSelectionText))
							{
								selFlagDirectRecipient = true;
								semanticRole = SemanticRole.DirectRecipient;
							}
							else if (selFlagDirectRecipient)
							{
								if (trimmedSelectionText.StartsWith("-") || trimmedSelectionText.StartsWith("–") || trimmedSelectionText.StartsWith("—"))
								{
									semanticRole = SemanticRole.DirectRecipient;
								}
								else
								{
									selFlagDirectRecipient = false;
								}
							}
							if (!selFlagDirectRecipient)
							{
								if (IsRomanNumeralHeading(trimmedSelectionText))
								{
									semanticRole = SemanticRole.SectionHeader;
								}
								else if (RegexPatterns.PartHeader.IsMatch(trimmedSelectionText))
								{
									semanticRole = SemanticRole.PartHeader;
								}
								else if (RegexPatterns.ChapterHeader.IsMatch(trimmedSelectionText))
								{
									semanticRole = SemanticRole.ChapterHeader;
								}
								else if (RegexPatterns.SectionHeader.IsMatch(trimmedSelectionText))
								{
									semanticRole = SemanticRole.SectionHeader;
								}
								else if (RegexPatterns.ArticleTitle.IsMatch(trimmedSelectionText))
								{
									semanticRole = SemanticRole.ArticleTitle;
								}
								else if (IsLetterHeading(trimmedSelectionText))
								{
									semanticRole = (char.IsLower(trimmedSelectionText[0]) ? SemanticRole.PointText : SemanticRole.SectionHeader);
								}
								else if (RegexPatterns.PointLetter.IsMatch(trimmedSelectionText))
								{
									semanticRole = SemanticRole.PointText;
								}
								else if (IsClauseHeading(trimmedSelectionText) || RegexPatterns.ClauseNumber.IsMatch(trimmedSelectionText))
								{
									semanticRole = SemanticRole.ClauseText;
								}
								else if (RegexPatterns.LegalGround.IsMatch(trimmedSelectionText) || trimmedSelectionText.StartsWith("Căn cứ", StringComparison.OrdinalIgnoreCase))
								{
									semanticRole = SemanticRole.LegalGround;
								}
								else
								{
									semanticRole = SemanticRole.BodyText;
								}
							}
							ComponentStyle componentStyle = DocumentStyleResolver.Resolve(activeProfile, semanticRole, documentType);
							ParagraphFormatter.FormatParagraph(comObj3, componentStyle, activeProfile.Typography);
							float bodyIndentPt = activeProfile.GetComponentStyle(SemanticRole.BodyText)?.FirstLineIndentPt ?? activeProfile.Typography.FirstLineIndentPt;
							if (bodyIndentPt <= 0f)
							{
								bodyIndentPt = UnitConverter.MmToPt(10f);
							}
							if ((semanticRole == SemanticRole.SectionHeader || semanticRole == SemanticRole.PartHeader) && (IsRomanNumeralHeading(trimmedSelectionText) || IsLetterHeading(trimmedSelectionText)))
							{
								comObj3.Range.Font.Bold = -1;
								comObj3.Range.Font.Italic = 0;
								comObj3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
								comObj3.Format.FirstLineIndent = bodyIndentPt;
								comObj3.Format.LeftIndent = 0f;
								comObj3.Format.RightIndent = 0f;
								comObj3.Format.KeepWithNext = -1;
								comObj3.Format.KeepTogether = -1;
							}
							if ((semanticRole == SemanticRole.SectionHeader || semanticRole == SemanticRole.PartHeader) && IsRomanNumeralHeading(trimmedSelectionText) && trimmedSelectionText.Length <= 200 && IsTitledClause(Regex.Replace(trimmedSelectionText, "^\\s*[IVXivx]{1,4}\\s*[.\\-]\\s*", "")))
							{
								try
								{
									string text6 = comObj3.Range.Text ?? "";
									int num4 = text6.Length;
									while (num4 > 0 && (text6[num4 - 1] == '\r' || text6[num4 - 1] == '\a' || text6[num4 - 1] == '\n'))
									{
										num4--;
									}
									if (num4 > 0 && comObj3.Range.Fields.Count == 0)
									{
										string text7 = text6.Substring(0, num4);
										string text8 = text7.ToUpper(new CultureInfo("vi-VN"));
										if (text8 != text7)
										{
											Range comObj6 = comObj3.Range.Duplicate;
											comObj6.SetRange(comObj3.Range.Start, comObj3.Range.Start + num4);
											comObj6.Text = text8;
											ComReleaser.Release(ref comObj6);
											comObj3.Range.Font.Bold = -1;
										}
									}
									comObj3.Range.Font.AllCaps = 0;
								}
								catch
								{
								}
							}
							if (GetYouthMultiLevelDepth(trimmedSelectionText) >= 2)
							{
								comObj3.Format.FirstLineIndent = bodyIndentPt;
								comObj3.Format.LeftIndent = 0f;
								comObj3.Format.RightIndent = 0f;
								ApplyYouthMultiLevelStyle(comObj3, trimmedSelectionText);
							}
							else if (semanticRole == SemanticRole.ClauseText || semanticRole == SemanticRole.ArticleTitle)
							{
								comObj3.Format.FirstLineIndent = bodyIndentPt;
								comObj3.Format.LeftIndent = 0f;
								comObj3.Format.RightIndent = 0f;
								string text9 = comObj3.Range.Text ?? "";
								if (text9.Length > 2 && char.IsDigit(text9[0]))
								{
									int num6 = text9.IndexOf('.');
									if (num6 > 0 && num6 <= 3)
									{
										int num7 = num6 + 1;
										int j;
										for (j = 0; num7 + j < text9.Length && (text9[num7 + j] == ' ' || text9[num7 + j] == '\t' || text9[num7 + j] == '\u00a0'); j++)
										{
										}
										if (j > 1 || (j == 1 && (text9[num7] == '\t' || text9[num7] == '\u00a0')))
										{
											try
											{
												Range comObj8 = comObj3.Range.Duplicate;
												comObj8.End = comObj8.Start + num7 + j;
												comObj8.Text = text9.Substring(0, num7) + " ";
												ComReleaser.Release(ref comObj8);
												trimmedSelectionText = text9.Substring(0, num7) + " " + text9.Substring(num7 + j).Replace("\r", "").Replace("\a", "").Trim();
											}
											catch
											{
											}
										}
									}
								}
								int colonIdx = trimmedSelectionText.IndexOf(':');
								if (IsLeadInSentence(trimmedSelectionText))
								{
									comObj3.Range.Font.Bold = 0;
								}
								else if (colonIdx > 0 && colonIdx < 120)
								{
									try
									{
										Range rColon = comObj3.Range.Duplicate;
										rColon.End = rColon.Start + colonIdx + 1;
										rColon.Font.Bold = -1;
										rColon.Font.Italic = 0;
										ComReleaser.Release(ref rColon);
									}
									catch
									{
									}
								}
								else if (IsTitledClause(trimmedSelectionText))
								{
									comObj3.Range.Font.Bold = -1;
									comObj3.Format.KeepWithNext = -1;
									comObj3.Format.KeepTogether = -1;
								}
							}
							if (semanticRole == SemanticRole.PointText && IsLetterHeading(trimmedSelectionText) && IsTitledClause(trimmedSelectionText.TrimEnd(':')))
							{
								comObj3.Range.Font.Bold = -1;
								comObj3.Format.KeepWithNext = -1;
								comObj3.Format.KeepTogether = -1;
							}
							if (semanticRole == SemanticRole.PointText && RegexPatterns.PointLetter.IsMatch(trimmedSelectionText) && trimmedSelectionText.Length <= 160 && !IsLeadInSentence(trimmedSelectionText) && IsTitledClause(Regex.Replace(trimmedSelectionText, "^\\s*[a-zđA-ZĐ]\\)\\s*", "").TrimEnd(':')))
							{
								try
								{
									string input3 = comObj3.Range.Text ?? "";
									Match match3 = Regex.Match(input3, "^\\s*[a-zđ]\\)\\s+(\\p{Ll})");
									if (match3.Success && comObj3.Range.Fields.Count == 0)
									{
										Range comObj7 = comObj3.Range.Duplicate;
										int num5 = comObj3.Range.Start + match3.Groups[1].Index;
										comObj7.SetRange(num5, num5 + 1);
										comObj7.Text = match3.Groups[1].Value.ToUpper(CultureInfo.CurrentCulture);
										ComReleaser.Release(ref comObj7);
									}
								}
								catch
								{
								}
								comObj3.Range.Font.Bold = -1;
								comObj3.Format.KeepWithNext = -1;
								comObj3.Format.KeepTogether = -1;
							}
							if (semanticRole == SemanticRole.DirectRecipient || RegexPatterns.DirectRecipient.IsMatch(trimmedSelectionText))
							{
								try
								{
									comObj3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
									comObj3.Format.FirstLineIndent = 0f;
									comObj3.Format.LeftIndent = 0f;
									comObj3.Format.RightIndent = 0f;
									comObj3.Range.Font.Size = GetComponentSizeOrDefault(activeProfile, SemanticRole.DirectRecipient, 14f);
									if (trimmedSelectionText.StartsWith("-") || trimmedSelectionText.StartsWith("–") || trimmedSelectionText.StartsWith("—"))
									{
										comObj3.Range.Font.Bold = 0;
										comObj3.Range.Font.Italic = 0;
									}
									else if (RulePackTraits.IsYouthUnion(activeProfile.RulePack))
									{
										int num9 = trimmedSelectionText.IndexOf(':');
										int num10 = ((num9 > 0) ? (num9 + 1) : Math.Min(trimmedSelectionText.Length, 9));
										Range comObj10 = comObj3.Range.Duplicate;
										comObj10.End = comObj10.Start + num10;
										comObj10.Font.Bold = -1;
										comObj10.Font.Italic = -1;
										ComReleaser.Release(ref comObj10);
										if (trimmedSelectionText.Length > num10)
										{
											Range comObj11 = comObj3.Range.Duplicate;
											comObj11.Start += num10;
											comObj11.Font.Bold = -1;
											comObj11.Font.Italic = 0;
											ComReleaser.Release(ref comObj11);
										}
									}
									else if (activeProfile.RulePack == RulePackType.HD05_VPTW)
									{
										int num11 = trimmedSelectionText.IndexOf(':');
										int num12 = ((num11 > 0) ? (num11 + 1) : Math.Min(trimmedSelectionText.Length, 9));
										Range comObj12 = comObj3.Range.Duplicate;
										comObj12.End = comObj12.Start + num12;
										comObj12.Font.Bold = 0;
										comObj12.Font.Italic = -1;
										ComReleaser.Release(ref comObj12);
										if (trimmedSelectionText.Length > num12)
										{
											Range comObj13 = comObj3.Range.Duplicate;
											comObj13.Start += num12;
											comObj13.Font.Bold = 0;
											comObj13.Font.Italic = 0;
											ComReleaser.Release(ref comObj13);
										}
									}
									else
									{
										int colonPos = trimmedSelectionText.IndexOf(':');
										if (colonPos > 0 && colonPos < 40 && comObj3.Range.Fields.Count == 0)
										{
											Range rDirect = comObj3.Range.Duplicate;
											rDirect.End = rDirect.Start + colonPos + 1;
											rDirect.Font.Bold = -1;
											rDirect.Font.Italic = 0;
											ComReleaser.Release(ref rDirect);
										}
									}
								}
								catch
								{
								}
							}
							if (semanticRole == SemanticRole.BodyText)
							{
								if (comObj3.Range.Font.Italic != 0)
								{
									comObj3.Range.Font.Italic = 0;
								}
								if (comObj3.Range.Font.Bold != 0 && (trimmedSelectionText.Length > 80 || IsBodyStarter(trimmedSelectionText)))
								{
									comObj3.Range.Font.Bold = 0;
								}
							}
							if (componentStyle != null && componentStyle.UnderlineRule != UnderlineRuleType.None)
							{
								UnderlineFormatter.ApplyUnderlineRule(comObj3, componentStyle.UnderlineRule);
							}
							ParagraphFormatter.FixOrphanWord(comObj3);
							num++;
						}
						catch
						{
						}
						finally
						{
							ComReleaser.Release(ref comObj3);
						}
					}
					ComReleaser.Release(ref comObj2);
					stopwatch.Stop();
					return FormattingResult.Succeeded(rulePackType, num, tablesProcessed, stopwatch.ElapsedMilliseconds);
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
		finally
		{
			comObj = null;
		}
	}

	private static int FindParagraphIndex(Document doc, int rangeStart, int fallback)
	{
		if (doc == null)
		{
			return fallback;
		}
		try
		{
			int count = doc.Paragraphs.Count;
			for (int i = 1; i <= count; i++)
			{
				OperationProgress.Step();
				Paragraph comObj = doc.Paragraphs[i];
				try
				{
					if (comObj.Range.Start >= rangeStart || comObj.Range.End > rangeStart)
					{
						return i;
					}
				}
				finally
				{
					ComReleaser.Release(ref comObj);
				}
			}
		}
		catch
		{
		}
		return fallback;
	}

	public static void EnsureEditablePrintView(Document doc)
	{
		if (doc == null)
		{
			return;
		}
		try
		{
			if (doc.ActiveWindow != null && doc.ActiveWindow.View != null)
			{
				if (doc.ActiveWindow.View.ReadingLayout)
				{
					doc.ActiveWindow.View.ReadingLayout = false;
				}
				if (doc.ActiveWindow.View.Type == WdViewType.wdReadingView)
				{
					doc.ActiveWindow.View.Type = WdViewType.wdPrintView;
				}
			}
		}
		catch
		{
		}
	}

	private static bool IsHeaderRebuildSafe(Document doc, Table existingHeader, RulePackType rulePack)
	{
		if (existingHeader != null)
		{
			return true;
		}
		if (doc == null)
		{
			return false;
		}
		int num = 0;
		int num2 = Math.Min(doc.Paragraphs.Count, 20);
		bool flag = false;
		bool flag2 = false;
		try
		{
			for (int i = 1; i <= num2; i++)
			{
				OperationProgress.Step();
				Paragraph comObj = doc.Paragraphs[i];
				try
				{
					if (comObj.Range.Tables.Count > 0)
					{
						continue;
					}
					string text = (comObj.Range.Text ?? string.Empty).Replace("\r", "").Replace("\n", "").Replace("\a", "")
						.Trim();
					if (text.Length != 0)
					{
						if (RegexPatterns.NationalHeader.IsMatch(text) || RegexPatterns.Motto.IsMatch(text) || RegexPatterns.PartyHeader.IsMatch(text) || RegexPatterns.YouthUnionHeader.IsMatch(text))
						{
							num++;
						}
						if (RegexPatterns.AdminAgency.IsMatch(text) || RegexPatterns.PartyAgency.IsMatch(text) || RegexPatterns.MilitaryAgency.IsMatch(text) || RegexPatterns.YouthUnionAgency.IsMatch(text))
						{
							flag = true;
						}
						if (RegexPatterns.DocumentNumberGeneral.IsMatch(text) || RegexPatterns.DocumentDate.IsMatch(text))
						{
							flag2 = true;
						}
					}
				}
				finally
				{
					ComReleaser.Release(ref comObj);
				}
			}
		}
		catch
		{
			return false;
		}
		if (flag)
		{
			num++;
		}
		if (flag2)
		{
			num++;
		}
		if (num >= 2)
		{
			if (!flag)
			{
				return RulePackTraits.UsesOrganizationHeader(rulePack);
			}
			return true;
		}
		return false;
	}

	private static bool IsLikelyDocumentTitle(string text, int paragraphIndex, bool hasEnteredBodyStructure)
	{
		if (string.IsNullOrWhiteSpace(text) || paragraphIndex > 30 || hasEnteredBodyStructure)
		{
			return false;
		}
		string text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
			.Trim();
		if (text2.Length < 3 || text2.Length > 120)
		{
			return false;
		}
		if (text2.EndsWith(".") || text2.EndsWith(";"))
		{
			return false;
		}
		string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
		if (text3.StartsWith("CAN CU") || text3.StartsWith("THEO ") || (text3.StartsWith("BAO CAO ") && text2.Length > 45))
		{
			return false;
		}
		if (!HeaderRebuilder.IsDocumentTitle(text2) && !RegexPatterns.DocumentType.IsMatch(text2))
		{
			return false;
		}
		bool flag = text2.ToUpperInvariant() == text2;
		bool result = text2.Length <= 35 && !text2.Contains(",");
		if (!flag)
		{
			return result;
		}
		return true;
	}

		private static bool IsMetadataRole(SemanticRole role)
	{
		switch (role)
		{
		case SemanticRole.NationalHeader:
		case SemanticRole.Motto:
		case SemanticRole.PartyHeader:
		case SemanticRole.SuperiorAgency:
		case SemanticRole.IssuingAgency:
		case SemanticRole.DocumentNumber:
		case SemanticRole.DocumentDate:
		case SemanticRole.RecipientHeader:
		case SemanticRole.RecipientItem:
		case SemanticRole.SignAuthority:
		case SemanticRole.SignTitle:
		case SemanticRole.SignerName:
		case SemanticRole.SignRankName:
			return true;
		default:
			return false;
		}
	}

	private static bool IsProtectedFromBodySelection(SemanticRole role)
	{
		switch (role)
		{
		case SemanticRole.NationalHeader:
		case SemanticRole.Motto:
		case SemanticRole.PartyHeader:
		case SemanticRole.SuperiorAgency:
		case SemanticRole.IssuingAgency:
		case SemanticRole.DocumentNumber:
		case SemanticRole.DocumentDate:
		case SemanticRole.DocumentType:
		case SemanticRole.SubjectExcerpt:
		case SemanticRole.DispatchSubject:
		case SemanticRole.SignAuthority:
		case SemanticRole.SignTitle:
		case SemanticRole.SignerName:
		case SemanticRole.SignRankName:
		case SemanticRole.UrgencyIndicator:
		case SemanticRole.SecurityIndicator:
		case SemanticRole.CirculationScope:
		case SemanticRole.DraftIndicator:
			return true;
		default:
			return false;
		}
	}

	public FormattingResult ApplyPageSetupOnly(Document doc, ProfileModel profile)
	{
		return ApplyPageSetupOnly(doc, profile, null);
	}

	public FormattingResult ApplyPageSetupOnly(Document doc, ProfileModel profile = null, PrintLayoutMode? printLayoutMode = null)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application comObj = null;
		try
		{
			comObj = doc.Application;
			using (new ScreenFreezeScope(comObj))
			{
				using (new SafeUndoScope(comObj, "OfficeFormat Chuẩn Hóa Lề Trang"))
				{
					ProfileModel profileModel = profile ?? _ruleEngine.ActiveProfile;
					if (profileModel != null && profileModel.PageSetup != null && printLayoutMode.HasValue)
					{
						profileModel.PageSetup.MirrorMargins = printLayoutMode.Value == PrintLayoutMode.TwoSidedMirror;
					}
					PageSetupFormatter.ApplyPageSetup(doc, profileModel);
					stopwatch.Stop();
					return FormattingResult.Succeeded(profileModel.RulePack, 0, 0, stopwatch.ElapsedMilliseconds);
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
		finally
		{
			comObj = null;
		}
	}

	public FormattingResult ApplyTablesOnly(Document doc, TablePreset preset = null)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		EnsureEditablePrintView(doc);
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application comObj = null;
		try
		{
			comObj = doc.Application;
			using (new ScreenFreezeScope(comObj))
			{
				using (new SafeUndoScope(comObj, "OfficeFormat Chuẩn Hóa Bảng Biểu"))
				{
					TablePreset preset2 = preset ?? _ruleEngine.ActiveProfile.TablePreset;
					TableFormatter.FormatTables(doc, preset2);
					stopwatch.Stop();
					return FormattingResult.Succeeded(_ruleEngine.ActiveProfile.RulePack, 0, doc.Tables.Count, stopwatch.ElapsedMilliseconds);
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
		finally
		{
			comObj = null;
		}
	}

	public FormattingResult FixDocumentTables(Document doc, RulePackType? forcedRulePack = null)
	{
		if (forcedRulePack.HasValue)
		{
			_ruleEngine.SetActiveProfile(forcedRulePack.Value);
		}
		return ApplyTablesOnly(doc, _ruleEngine.ActiveProfile.TablePreset);
	}

	public FormattingResult FixDocumentSpacing(Document doc, RulePackType? forcedRulePack = null)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application comObj = null;
		try
		{
			comObj = doc.Application;
			using (new ScreenFreezeScope(comObj))
			{
				using (new SafeUndoScope(comObj, "OfficeFormat Chuẩn Hóa Giãn Dòng & Đoạn"))
				{
					RulePackType rulePackType = (forcedRulePack.HasValue ? forcedRulePack.Value : _ruleEngine.ActiveProfile.RulePack);
					_ruleEngine.SetActiveProfile(rulePackType);
					ProfileModel activeProfile = _ruleEngine.ActiveProfile;
					int paragraphs = FormatAllParagraphs(doc, activeProfile);
					stopwatch.Stop();
					return FormattingResult.Succeeded(rulePackType, paragraphs, 0, stopwatch.ElapsedMilliseconds);
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
		finally
		{
			comObj = null;
		}
	}

	public FormattingResult ApplyPageNumbersOnly(Document doc, PageNumberPreset preset = null)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application comObj = null;
		try
		{
			comObj = doc.Application;
			using (new ScreenFreezeScope(comObj))
			{
				using (new SafeUndoScope(comObj, "OfficeFormat Đánh Số Trang"))
				{
					PageNumberPreset preset2 = preset ?? _ruleEngine.ActiveProfile.PageNumber;
					HeaderFooterFormatter.ApplyPageNumbers(doc, preset2);
					stopwatch.Stop();
					return FormattingResult.Succeeded(_ruleEngine.ActiveProfile.RulePack, 0, 0, stopwatch.ElapsedMilliseconds);
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
		finally
		{
			comObj = null;
		}
	}

	public Table InsertAgencyHeader(Document doc, AgencyProfile agency, ProfileModel profile = null)
	{
		if (doc == null)
		{
			return null;
		}
		ProfileModel profile2 = profile ?? _ruleEngine.ActiveProfile;
		Application application = doc.Application;
		using (new ScreenFreezeScope(application))
		{
			using (new SafeUndoScope(application, "OfficeFormat Chèn Khối Tiêu Đề"))
			{
				return NationalHeaderBuilder.InsertOrUpdateHeaderTable(doc, profile2, agency);
			}
		}
	}

	public Table InsertSignerBlock(Document doc, SignerProfile signer, RecipientProfile recipient, ProfileModel profile = null)
	{
		if (doc == null)
		{
			return null;
		}
		ProfileModel profile2 = profile ?? _ruleEngine.ActiveProfile;
		Application application = doc.Application;
		using (new ScreenFreezeScope(application))
		{
			using (new SafeUndoScope(application, "OfficeFormat Chèn Khối Ký Tên"))
			{
				return SignerBlockBuilder.InsertOrUpdateSignerTable(doc, profile2, signer, recipient);
			}
		}
	}

	private static string ConvertToUnsigned(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return "";
		}
		string pattern = "\\p{IsCombiningDiacriticalMarks}+";
		string input = text.Normalize(NormalizationForm.FormD);
		return Regex.Replace(input, pattern, string.Empty).Replace('đ', 'd').Replace('Đ', 'D');
	}

	public static bool IsBodyStarter(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
			.Trim();
		if (text2.Length == 0)
		{
			return false;
		}
		string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
		bool flag = text2.Length > 5 && text2.ToUpperInvariant() == text2;
		if (!flag && (text3.StartsWith("QUA KIEM TRA") || text3.StartsWith("QUA ") || text3.StartsWith("CAN CU") || text3.StartsWith("XET ") || text3.StartsWith("THOI GIAN QUA") || text3.StartsWith("TRONG NHUNG NAM QUA") || text3.StartsWith("TRONG THOI GIAN QUA")))
		{
			return true;
		}
		if (!flag && (text3.StartsWith("THUC HIEN") || text3.StartsWith("NHAM ") || text3.StartsWith("DE ")) && (text2.EndsWith(":") || text3.EndsWith("NHU SAU:") || text3.EndsWith("NHU SAU :") || (text2.Length > 80 && text2.EndsWith("."))))
		{
			return true;
		}
		if (text3.StartsWith("KINH GUI") || text3.StartsWith("DIEU ") || text3.StartsWith("QUYET DINH:") || text3.StartsWith("BAN HANH:") || text3.StartsWith("QUY DINH:") || text3.StartsWith("THONG BAO:") || text3.StartsWith("KET LUAN:") || text3.StartsWith("I. ") || text3.StartsWith("II. ") || text3.StartsWith("III. ") || text3.StartsWith("IV. ") || text3.StartsWith("V. ") || text3.StartsWith("VI. ") || text3.StartsWith("VII. ") || text3.StartsWith("1. ") || text3.StartsWith("2. ") || text3.StartsWith("3. ") || text3.StartsWith("- ") || text3.StartsWith("+ ") || text3.StartsWith("NOI NHAN") || text3.StartsWith("TM.") || text3.StartsWith("KT.") || text3.StartsWith("PHAN ") || text3.StartsWith("CHUONG ") || text3.StartsWith("MUC ") || text3.StartsWith("TIEU MUC "))
		{
			return true;
		}
		if (text3.EndsWith("NHU SAU:") || text3.EndsWith("NHU SAU :") || text3.EndsWith(":"))
		{
			return true;
		}
		if (!char.IsLower(text2[0]) && text2.Length > 200)
		{
			return true;
		}
		return false;
	}

	public static bool IsRomanNumeralHeading(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = text.Trim();
		if (text2.Length < 3)
		{
			return false;
		}
		int num = text2.IndexOf('.');
		int num2 = text2.IndexOf("- ", StringComparison.Ordinal);
		if (num2 >= 1 && num2 <= 6 && (num < 0 || num2 < num))
		{
			num = num2;
		}
		if (num >= 1 && num <= 6)
		{
			switch (text2.Substring(0, num).Trim().ToUpperInvariant())
			{
			case "I":
			case "II":
			case "III":
			case "IV":
			case "V":
			case "VI":
			case "VII":
			case "VIII":
			case "IX":
			case "X":
				return true;
			}
		}
		return false;
	}

	public static bool IsTitledClause(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = text.Replace("\r", "").Replace("\a", "").Trim();
		if (text2.Length == 0 || text2.Length > 240)
		{
			return false;
		}
		char c = text2[text2.Length - 1];
		if (c == '.' || c == ';' || c == ',')
		{
			return false;
		}
		string input = Regex.Replace(text2, "^(?:\\d+(?:\\.\\d+)*\\.?|[IVXLCDM]+\\s*[.\\-]|[A-Za-zĐđ][.\\)])\\s*", "", RegexOptions.IgnoreCase);
		input = Regex.Replace(input, "“[^”]*”|\"[^\"]*\"", "");
		if (input.IndexOfAny(new char[4] { '.', ';', '!', '?' }) >= 0)
		{
			return false;
		}
		return true;
	}

	public static int GetYouthMultiLevelDepth(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return 0;
		}
		Match match = YouthMultiLevelNumber.Match(text);
		if (!match.Success)
		{
			return 0;
		}
		return match.Groups[1].Value.Split('.').Length;
	}

	private static void ApplyYouthMultiLevelStyle(Paragraph p, string cleanTxt)
	{
		int youthMultiLevelDepth = GetYouthMultiLevelDepth(cleanTxt);
		if (youthMultiLevelDepth < 2 || p == null)
		{
			return;
		}
		int bold = ((youthMultiLevelDepth == 2) ? (-1) : 0);
		int italic = ((youthMultiLevelDepth <= 3) ? (-1) : 0);
		try
		{
			string text = YouthMultiLevelNumber.Replace(cleanTxt, "").Trim();
			if (!IsLeadInSentence(cleanTxt) && IsTitledClause(text.TrimEnd(':')))
			{
				p.Range.Font.Bold = bold;
				p.Range.Font.Italic = italic;
				p.Format.KeepWithNext = -1;
				p.Format.KeepTogether = -1;
				return;
			}
			int colonIdx = cleanTxt.IndexOf(':');
			if (!IsLeadInSentence(cleanTxt) && colonIdx > 0 && colonIdx < 120 && p.Range.Fields.Count == 0)
			{
				p.Range.Font.Bold = 0;
				p.Range.Font.Italic = 0;
				Range comObjColon = p.Range.Duplicate;
				comObjColon.SetRange(p.Range.Start, p.Range.Start + colonIdx + 1);
				comObjColon.Font.Bold = bold;
				comObjColon.Font.Italic = italic;
				ComReleaser.Release(ref comObjColon);
				return;
			}
			string input = p.Range.Text ?? "";
			Match match = YouthMultiLevelNumber.Match(input);
			if (match.Success && p.Range.Fields.Count == 0)
			{
				p.Range.Font.Bold = 0;
				p.Range.Font.Italic = 0;
				Range comObj = p.Range.Duplicate;
				comObj.SetRange(p.Range.Start, p.Range.Start + match.Length);
				comObj.Font.Bold = bold;
				comObj.Font.Italic = italic;
				ComReleaser.Release(ref comObj);
			}
		}
		catch
		{
		}
	}

	public static bool IsClauseHeading(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = text.Trim();
		if (text2.Length < 3)
		{
			return false;
		}
		int num = text2.IndexOf('.');
		if (num >= 1 && num <= 4)
		{
			string s = text2.Substring(0, num).Trim();
			if (int.TryParse(s, out var result) && result >= 1 && result <= 99)
			{
				return true;
			}
		}
		return false;
	}

	public static bool IsLetterHeading(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = text.Trim();
		if (text2.Length < 3 || text2.Length > 80)
		{
			return false;
		}
		if (char.IsLetter(text2[0]) && text2[1] == '.' && (text2[2] == ' ' || text2[2] == '\t' || text2[2] == '\u00a0'))
		{
			char c = char.ToUpperInvariant(text2[0]);
			if ((c >= 'A' && c <= 'Z') || c == 'Đ')
			{
				return true;
			}
		}
		return false;
	}

	public static bool IsLeadInSentence(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		return LeadInEnding.IsMatch(text.Replace("\r", "").Replace("\a", "").Trim());
	}

	private static int CleanupInlineSpacing(Document doc)
	{
		int num = 0;
		Paragraphs comObj = doc.Paragraphs;
		int count = comObj.Count;
		for (int i = 1; i <= count; i++)
		{
			OperationProgress.Step();
			Paragraph comObj2 = null;
			try
			{
				comObj2 = comObj[i];
				Range comObj3 = comObj2.Range;
				string text = comObj3.Text ?? "";
				if (text.Length < 3 || text.IndexOf('\t') >= 0 || (text.IndexOf("  ", StringComparison.Ordinal) < 0 && text.IndexOf("\u00a0", StringComparison.Ordinal) < 0 && !SpaceBeforePunct.IsMatch(text) && !SpaceAfterOpenParen.IsMatch(text)) || comObj3.Fields.Count > 0)
				{
					continue;
				}
				bool flag = comObj3.Tables.Count > 0;
				bool[] array = new bool[text.Length];
				foreach (Match item in MultiSpaceRun.Matches(text))
				{
					if (item.Length < 4 || flag)
					{
						for (int j = item.Index + 1; j < item.Index + item.Length; j++)
						{
							array[j] = true;
						}
					}
				}
				foreach (Match item2 in SpaceBeforePunct.Matches(text))
				{
					if (item2.Length >= 4 && !flag)
					{
						continue;
					}
					int num2 = item2.Index + item2.Length;
					if (text[num2] != '.' || num2 + 1 >= text.Length || !char.IsDigit(text[num2 + 1]))
					{
						for (int k = item2.Index; k < num2; k++)
						{
							array[k] = true;
						}
					}
				}
				foreach (Match item3 in SpaceAfterOpenParen.Matches(text))
				{
					if (item3.Length < 4 || flag)
					{
						for (int l = item3.Index; l < item3.Index + item3.Length; l++)
						{
							array[l] = true;
						}
					}
				}
				int start = comObj3.Start;
				int num3 = text.Length - 1;
				while (num3 >= 0)
				{
					if (!array[num3])
					{
						num3--;
						continue;
					}
					int num4 = num3;
					while (num4 - 1 >= 0 && array[num4 - 1])
					{
						num4--;
					}
					object Start = start + num4;
					object End = start + num3 + 1;
					Range comObj4 = doc.Range(ref Start, ref End);
					string text2 = comObj4.Text ?? "";
					if (text2.Length == num3 - num4 + 1 && text2.Trim(' ', '\u00a0').Length == 0)
					{
						Range range = comObj4;
						object Unit = Missing.Value;
						object Count = Missing.Value;
						range.Delete(ref Unit, ref Count);
						num++;
					}
					ComReleaser.Release(ref comObj4);
					num3 = num4 - 1;
				}
				ComReleaser.Release(ref comObj3);
			}
			catch
			{
			}
			finally
			{
				ComReleaser.Release(ref comObj2);
			}
		}
		ComReleaser.Release(ref comObj);
		return num;
	}

	private static bool IsBlankOrBreakText(string t)
	{
		if (t == null)
		{
			return true;
		}
		foreach (char c in t)
		{
			if (c != '\r' && c != '\n' && c != '\a' && c != '\f' && c != '\v' && c != '\t' && c != ' ' && c != '\u00a0')
			{
				return false;
			}
		}
		return true;
	}

	private static int RemoveTrailingBlankPages(Document doc)
	{
		return WordToolEngine.DeleteBlankPages(doc);
	}
	private static void StripLeadingTabs(Paragraph p)
	{
		try
		{
			string text = p.Range.Text;
			if (!string.IsNullOrEmpty(text) && (text[0] == '\t' || text[0] == ' ' || text[0] == '\u00a0'))
			{
				int i;
				for (i = 0; i < text.Length && (text[i] == '\t' || text[i] == ' ' || text[i] == '\u00a0') && text[i] != '\r' && text[i] != '\n'; i++)
				{
				}
				if (i > 0 && i < text.Length)
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

	private static void NormalizeHeadingAndListMarkers(Paragraph p)
	{
		if (p == null)
		{
			return;
		}
		try
		{
			WdListType listType = p.Range.ListFormat.ListType;
			if (listType == WdListType.wdListBullet)
			{
				object NumberType = WdNumberType.wdNumberParagraph;
				p.Range.ListFormat.RemoveNumbers(ref NumberType);
				StripLeadingTabs(p);
				string bulletTxt = (p.Range.Text ?? "").TrimStart();
				if (bulletTxt.Length > 0 && bulletTxt[0] != '-' && bulletTxt[0] != '+')
				{
					p.Range.InsertBefore("- ");
				}
			}
			else if (listType != WdListType.wdListNoNumbering)
			{
				object NumberType2 = WdNumberType.wdNumberParagraph;
				p.Range.ListFormat.ConvertNumbersToText(ref NumberType2);
			}
		}
		catch
		{
		}
		StripLeadingTabs(p);
		try
		{
			if (p.Range.Fields.Count == 0)
			{
				string raw = p.Range.Text ?? "";
				Match mRun = HeadingOrListMarkerRun.Match(raw);
				if (mRun.Success)
				{
					string marker = mRun.Groups["marker"].Value;
					string fullMatch = mRun.Value;
					if (fullMatch != marker + " ")
					{
						Range rMarker = p.Range.Duplicate;
						rMarker.SetRange(p.Range.Start, p.Range.Start + fullMatch.Length);
						rMarker.Text = marker + " ";
						ComReleaser.Release(ref rMarker);
					}
				}
				else
				{
					Match mMultiNoDot = MultiLevelNoDotMarkerRun.Match(raw);
					if (mMultiNoDot.Success)
					{
						string marker2 = mMultiNoDot.Groups["marker"].Value;
						string fullMatch2 = mMultiNoDot.Value;
						if (fullMatch2 != marker2 + " ")
						{
							Range rMarker2 = p.Range.Duplicate;
							rMarker2.SetRange(p.Range.Start, p.Range.Start + fullMatch2.Length);
							rMarker2.Text = marker2 + " ";
							ComReleaser.Release(ref rMarker2);
						}
					}
				}
			}
		}
		catch
		{
		}
		try
		{
			string afterNorm = p.Range.Text ?? "";
			if (afterNorm.IndexOf('\t') < 0)
			{
				p.Format.TabStops.ClearAll();
			}
		}
		catch
		{
		}
	}

	private int FormatAllParagraphs(Document doc, ProfileModel profile, Action<int, int> onProgress = null)
	{
		Paragraphs comObj = doc.Paragraphs;
		int count = comObj.Count;
		int num = 0;
		bool flag = profile.RulePack == RulePackType.HD05_VPTW;
		_ = profile.RulePack;
		bool flag2 = RulePackTraits.IsYouthUnion(profile.RulePack);
		bool flag3 = false;
		int num2 = 0;
		bool flag4 = false;
		bool flag5 = false;
		bool flagDirectRecipient = false;
		string documentType = DetectPrimaryDocumentType(doc);
		bool flag6 = false;
		int num3 = SemanticComponentFormatter.RawHeaderEnd(doc);
		for (int i = 1; i <= count; i++)
		{
			OperationProgress.Step();
			if (onProgress != null && (i % 5 == 0 || i == count))
			{
				try
				{
					onProgress(i, count);
				}
				catch
				{
				}
			}
			Paragraph comObj2 = null;
			try
			{
				comObj2 = comObj[i];
				if (comObj2.Range.Tables.Count > 0)
				{
					continue;
				}
				string text = comObj2.Range.Text;
				if (string.IsNullOrEmpty(text) || text.Trim().Length == 0 || (num3 > 0 && comObj2.Range.Start < num3))
				{
					continue;
				}
				string text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
				if (!flag6 && i >= Math.Max(1, count - 20) && (SignatureRebuilder.HasStrongSignatureAnchor(text2) || RegexPatterns.IsStandaloneSignPosition(text2)))
				{
					flag6 = true;
				}
				if (flag6)
				{
					num++;
					continue;
				}
				NormalizeHeadingAndListMarkers(comObj2);
				text = comObj2.Range.Text ?? "";
				text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				text3 = ConvertToUnsigned(text2).ToUpperInvariant();
				try
				{
					string input = comObj2.Range.Text ?? "";
					Match match = MissingSpaceAfterMarker.Match(input);
					if (match.Success && comObj2.Range.Fields.Count == 0)
					{
						Range comObj3 = comObj2.Range.Duplicate;
						comObj3.SetRange(comObj2.Range.Start + match.Length, comObj2.Range.Start + match.Length);
						comObj3.InsertAfter(" ");
						ComReleaser.Release(ref comObj3);
						text = comObj2.Range.Text;
						text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
							.Trim();
						text3 = ConvertToUnsigned(text2).ToUpperInvariant();
					}
				}
				catch
				{
				}
				SemanticRole semanticRole = _ruleEngine.ClassifyParagraph(text, i, count, isInTable: false);
				if (semanticRole == SemanticRole.SignTitle && !RegexPatterns.IsStandaloneSignPosition(text2))
				{
					semanticRole = SemanticRole.BodyText;
				}
				if (semanticRole == SemanticRole.DocumentType && !IsLikelyDocumentTitle(text2, i, flag5))
				{
					semanticRole = SemanticRole.BodyText;
				}
				if (HeaderRebuilder.IsSeparatorLine(text2))
				{
					flag3 = false;
					comObj2.Range.ParagraphFormat.KeepWithNext = -1;
					comObj2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					comObj2.Range.ParagraphFormat.FirstLineIndent = 0f;
					comObj2.Range.ParagraphFormat.LeftIndent = 0f;
					comObj2.Range.ParagraphFormat.RightIndent = 0f;
					comObj2.Range.ParagraphFormat.SpaceBefore = 2f;
					comObj2.Range.ParagraphFormat.SpaceAfter = 6f;
					comObj2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					comObj2.Range.Font.Name = "Times New Roman";
					comObj2.Range.Font.Size = ((flag || flag2) ? 14f : 10f);
					comObj2.Range.Font.Bold = ((!flag2) ? (-1) : 0);
					comObj2.Range.Font.Italic = 0;
					comObj2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						comObj2.Range.ParagraphFormat.Borders.Enable = 0;
						foreach (Border border3 in comObj2.Borders)
						{
							try
							{
								border3.LineStyle = WdLineStyle.wdLineStyleNone;
							}
							catch
							{
							}
						}
					}
					catch
					{
					}
					num++;
					continue;
				}
				if (flag3)
				{
					bool flag7 = false;
					if (text3.StartsWith("CAN CU") || RegexPatterns.LegalGround.IsMatch(text2))
					{
						flag7 = true;
					}
					else if (IsRomanNumeralHeading(text2) || IsClauseHeading(text2) || IsLetterHeading(text2))
					{
						flag7 = true;
					}
					else if (text2.EndsWith(":") || text3.EndsWith("NHU SAU:") || text3.EndsWith("NHU SAU :"))
					{
						flag7 = true;
					}
					else if (IsAbstractBodyBoundary(comObj2, text2))
					{
						flag7 = true;
					}
					else if (num2 >= 4)
					{
						flag7 = true;
					}
					if (flag7)
					{
						flag3 = false;
					}
					else
					{
						semanticRole = SemanticRole.SubjectExcerpt;
						num2++;
					}
				}
				if (!flag3 && IsLikelyDocumentTitle(text2, i, flag5))
				{
					flag3 = true;
					num2 = 0;
					semanticRole = SemanticRole.DocumentType;
					string text4 = ConvertToUnsigned(text2).ToUpperInvariant();
					if (text4.Contains("QUYET DINH") || text4.Contains("NGHI QUYET") || text4.Contains("QUY CHE") || text4.Contains("QUY DINH"))
					{
						documentType = text2;
					}
				}
				if (!flag3)
				{
					if (IsRomanNumeralHeading(text2))
					{
						semanticRole = SemanticRole.SectionHeader;
						flag5 = true;
					}
					else if (IsLetterHeading(text2))
					{
						semanticRole = (char.IsLower(text2[0]) ? SemanticRole.PointText : SemanticRole.SectionHeader);
						flag5 = true;
					}
					else if (IsClauseHeading(text2))
					{
						semanticRole = SemanticRole.ClauseText;
						flag5 = true;
					}
					else if (RegexPatterns.ArticleTitle.IsMatch(text2))
					{
						semanticRole = SemanticRole.ArticleTitle;
						flag5 = true;
					}
					else if (RegexPatterns.PartHeader.IsMatch(text2))
					{
						semanticRole = SemanticRole.PartHeader;
						flag5 = true;
					}
					else if (RegexPatterns.ChapterHeader.IsMatch(text2))
					{
						semanticRole = SemanticRole.ChapterHeader;
						flag5 = true;
					}
					else if (RegexPatterns.SectionHeader.IsMatch(text2))
					{
						semanticRole = SemanticRole.SectionHeader;
						flag5 = true;
					}
					else if (RegexPatterns.PointLetter.IsMatch(text2))
					{
						semanticRole = SemanticRole.PointText;
						flag5 = true;
					}
					else if (text3 == "QUYET DINH:" || text3.StartsWith("QUYET DINH:") || text3 == "QUYET NGHI:" || text3.StartsWith("QUYET NGHI:") || text3 == "YEU CAU:" || text3.StartsWith("YEU CAU:") || text2.EndsWith("như sau:") || text2.EndsWith("như sau :"))
					{
						flag5 = true;
					}
				}
				if (RegexPatterns.DirectRecipient.IsMatch(text2))
				{
					flagDirectRecipient = true;
					semanticRole = SemanticRole.DirectRecipient;
				}
				else if (flagDirectRecipient)
				{
					if (text2.StartsWith("-") || text2.StartsWith("–") || text2.StartsWith("—"))
					{
						semanticRole = SemanticRole.DirectRecipient;
					}
					else if (!string.IsNullOrEmpty(text2))
					{
						flagDirectRecipient = false;
					}
				}
				if (RegexPatterns.RecipientHeader.IsMatch(text2))
				{
					flag4 = true;
					semanticRole = SemanticRole.RecipientHeader;
				}
				else if (flag4)
				{
					if (text2.StartsWith("-") || text2.StartsWith("+"))
					{
						semanticRole = SemanticRole.RecipientItem;
					}
					else
					{
						flag4 = false;
					}
				}
				else if (semanticRole == SemanticRole.RecipientItem)
				{
					semanticRole = SemanticRole.PointText;
				}
				if (!flag5 && !flag4 && (text3.StartsWith("CAN CU") || RegexPatterns.LegalGround.IsMatch(text2)))
				{
					semanticRole = SemanticRole.LegalGround;
					if (flag)
					{
						if (!text2.StartsWith("-"))
						{
							comObj2.Range.InsertBefore("- ");
							text2 = "- " + text2;
						}
					}
					else if (text2.StartsWith("- ") || text2.StartsWith("-"))
					{
						Range comObj4 = comObj2.Range.Duplicate;
						comObj4.End = comObj4.Start + ((!text2.StartsWith("- ")) ? 1 : 2);
						Range range = comObj4;
						object Unit = Missing.Value;
						object Count = Missing.Value;
						range.Delete(ref Unit, ref Count);
						ComReleaser.Release(ref comObj4);
					}
					try
					{
						string input2 = comObj2.Range.Text ?? "";
						Match match2 = Regex.Match(input2, "^(?:-\\s*)?Căn\\s+cứ\\s+(?:vào\\s+)?([a-zà-ỹ])", RegexOptions.IgnoreCase);
						if (match2.Success)
						{
							Group obj5 = match2.Groups[1];
							string value = obj5.Value;
							string text5 = value.ToUpper(CultureInfo.CurrentCulture);
							if (value != text5)
							{
								int index = obj5.Index;
								Range comObj5 = comObj2.Range.Duplicate;
								comObj5.Start = comObj2.Range.Start + index;
								comObj5.End = comObj5.Start + 1;
								comObj5.Text = text5;
								ComReleaser.Release(ref comObj5);
							}
						}
					}
					catch
					{
					}
				}
				else if (flag5 && semanticRole == SemanticRole.LegalGround)
				{
					semanticRole = SemanticRole.BodyText;
				}
				ComponentStyle componentStyle = DocumentStyleResolver.Resolve(profile, semanticRole, documentType);
				ParagraphFormatter.FormatParagraph(comObj2, componentStyle, profile.Typography);
				float bodyFirstLineIndentPt = profile.GetComponentStyle(SemanticRole.BodyText)?.FirstLineIndentPt ?? profile.Typography.FirstLineIndentPt;
				if (bodyFirstLineIndentPt <= 0f)
				{
					bodyFirstLineIndentPt = UnitConverter.MmToPt(10f);
				}
				if ((semanticRole == SemanticRole.BodyText || semanticRole == SemanticRole.ClauseText || semanticRole == SemanticRole.PointText) && comObj2.Range.Font.Italic != 0)
				{
					comObj2.Range.Font.Italic = 0;
				}
				if (semanticRole == SemanticRole.BodyText && comObj2.Range.Font.Bold != 0 && (text2.Length > 80 || IsBodyStarter(text2)))
				{
					comObj2.Range.Font.Bold = 0;
				}
				if ((semanticRole == SemanticRole.SectionHeader || semanticRole == SemanticRole.PartHeader) && (IsRomanNumeralHeading(text2) || IsLetterHeading(text2)))
				{
					comObj2.Range.Font.Bold = -1;
					comObj2.Range.Font.Italic = 0;
					comObj2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
					comObj2.Format.FirstLineIndent = bodyFirstLineIndentPt;
					comObj2.Format.LeftIndent = 0f;
					comObj2.Format.RightIndent = 0f;
					comObj2.Format.KeepWithNext = -1;
					comObj2.Format.KeepTogether = -1;
				}
				if ((semanticRole == SemanticRole.SectionHeader || semanticRole == SemanticRole.PartHeader) && IsRomanNumeralHeading(text2) && text2.Length <= 200 && IsTitledClause(Regex.Replace(text2, "^\\s*[IVXivx]{1,4}\\s*[.\\-]\\s*", "")))
				{
					try
					{
						string text6 = comObj2.Range.Text ?? "";
						int num4 = text6.Length;
						while (num4 > 0 && (text6[num4 - 1] == '\r' || text6[num4 - 1] == '\a' || text6[num4 - 1] == '\n'))
						{
							num4--;
						}
						if (num4 > 0 && comObj2.Range.Fields.Count == 0)
						{
							string text7 = text6.Substring(0, num4);
							string text8 = text7.ToUpper(new CultureInfo("vi-VN"));
							if (text8 != text7)
							{
								Range comObj6 = comObj2.Range.Duplicate;
								comObj6.SetRange(comObj2.Range.Start, comObj2.Range.Start + num4);
								comObj6.Text = text8;
								ComReleaser.Release(ref comObj6);
								comObj2.Range.Font.Bold = -1;
							}
						}
						comObj2.Range.Font.AllCaps = 0;
					}
					catch
					{
					}
				}
				if (semanticRole == SemanticRole.PointText && IsLetterHeading(text2) && IsTitledClause(text2.TrimEnd(':')))
				{
					comObj2.Range.Font.Bold = -1;
					comObj2.Format.KeepWithNext = -1;
					comObj2.Format.KeepTogether = -1;
				}
				if (semanticRole == SemanticRole.PointText && RegexPatterns.PointLetter.IsMatch(text2) && text2.Length <= 160 && !IsLeadInSentence(text2) && IsTitledClause(Regex.Replace(text2, "^\\s*[a-zđA-ZĐ]\\)\\s*", "").TrimEnd(':')))
				{
					try
					{
						string input3 = comObj2.Range.Text ?? "";
						Match match3 = Regex.Match(input3, "^\\s*[a-zđ]\\)\\s+(\\p{Ll})");
						if (match3.Success && comObj2.Range.Fields.Count == 0)
						{
							Range comObj7 = comObj2.Range.Duplicate;
							int num5 = comObj2.Range.Start + match3.Groups[1].Index;
							comObj7.SetRange(num5, num5 + 1);
							comObj7.Text = match3.Groups[1].Value.ToUpper(CultureInfo.CurrentCulture);
							ComReleaser.Release(ref comObj7);
						}
					}
					catch
					{
					}
					comObj2.Range.Font.Bold = -1;
					comObj2.Format.KeepWithNext = -1;
					comObj2.Format.KeepTogether = -1;
				}
				if (GetYouthMultiLevelDepth(text2) >= 2)
				{
					comObj2.Format.FirstLineIndent = bodyFirstLineIndentPt;
					comObj2.Format.LeftIndent = 0f;
					comObj2.Format.RightIndent = 0f;
					ApplyYouthMultiLevelStyle(comObj2, text2);
				}
				else if (semanticRole == SemanticRole.ClauseText)
				{
					comObj2.Format.FirstLineIndent = bodyFirstLineIndentPt;
					comObj2.Format.LeftIndent = 0f;
					comObj2.Format.RightIndent = 0f;
					string text9 = comObj2.Range.Text ?? "";
					if (text9.Length > 2 && char.IsDigit(text9[0]))
					{
						int num6 = text9.IndexOf('.');
						if (num6 > 0 && num6 <= 3)
						{
							int num7 = num6 + 1;
							int j;
							for (j = 0; num7 + j < text9.Length && (text9[num7 + j] == ' ' || text9[num7 + j] == '\t' || text9[num7 + j] == '\u00a0'); j++)
							{
							}
							if (j > 1 || (j == 1 && (text9[num7] == '\t' || text9[num7] == '\u00a0')))
							{
								try
								{
									Range comObj8 = comObj2.Range.Duplicate;
									comObj8.End = comObj8.Start + num7 + j;
									comObj8.Text = text9.Substring(0, num7) + " ";
									ComReleaser.Release(ref comObj8);
									text2 = text9.Substring(0, num7) + " " + text9.Substring(num7 + j).Replace("\r", "").Replace("\a", "")
										.Trim();
								}
								catch
								{
								}
							}
						}
					}
					int num8 = text2.IndexOf(':');
					if (IsLeadInSentence(text2))
					{
						try
						{
							comObj2.Range.Font.Bold = 0;
						}
						catch
						{
						}
					}
					else if (num8 > 0 && num8 < 120)
					{
						try
						{
							Range comObj9 = comObj2.Range.Duplicate;
							comObj9.End = comObj9.Start + num8 + 1;
							comObj9.Font.Bold = -1;
							comObj9.Font.Italic = 0;
							ComReleaser.Release(ref comObj9);
						}
						catch
						{
						}
					}
					else if (IsTitledClause(text2))
					{
						comObj2.Range.Font.Bold = -1;
						comObj2.Format.KeepWithNext = -1;
						comObj2.Format.KeepTogether = -1;
					}
				}
				if (semanticRole == SemanticRole.DirectRecipient || RegexPatterns.DirectRecipient.IsMatch(text2))
				{
					try
					{
						comObj2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
						comObj2.Format.FirstLineIndent = 0f;
						comObj2.Format.LeftIndent = 0f;
						comObj2.Format.RightIndent = 0f;
						comObj2.Range.Font.Size = GetComponentSizeOrDefault(profile, SemanticRole.DirectRecipient, 14f);
						if (text2.StartsWith("-") || text2.StartsWith("–") || text2.StartsWith("—"))
						{
							comObj2.Range.Font.Bold = 0;
							comObj2.Range.Font.Italic = 0;
						}
						else if (false) { }
						comObj2.Format.FirstLineIndent = 0f;
						comObj2.Format.LeftIndent = 0f;
						comObj2.Range.Font.Bold = 0;
						comObj2.Range.Font.Size = GetComponentSizeOrDefault(profile, SemanticRole.DirectRecipient, 14f);
						if (RulePackTraits.IsYouthUnion(profile.RulePack))
						{
							int num9 = text2.IndexOf(':');
							int num10 = ((num9 > 0) ? (num9 + 1) : Math.Min(text2.Length, 9));
							Range comObj10 = comObj2.Range.Duplicate;
							comObj10.End = comObj10.Start + num10;
							comObj10.Font.Bold = -1;
							comObj10.Font.Italic = -1;
							ComReleaser.Release(ref comObj10);
							if (text2.Length > num10)
							{
								Range comObj11 = comObj2.Range.Duplicate;
								comObj11.Start += num10;
								comObj11.Font.Bold = -1;
								comObj11.Font.Italic = 0;
								ComReleaser.Release(ref comObj11);
							}
						}
						else if (profile.RulePack == RulePackType.HD05_VPTW)
						{
							int num11 = text2.IndexOf(':');
							int num12 = ((num11 > 0) ? (num11 + 1) : Math.Min(text2.Length, 9));
							Range comObj12 = comObj2.Range.Duplicate;
							comObj12.End = comObj12.Start + num12;
							comObj12.Font.Bold = 0;
							comObj12.Font.Italic = -1;
							ComReleaser.Release(ref comObj12);
							if (text2.Length > num12)
							{
								Range comObj13 = comObj2.Range.Duplicate;
								comObj13.Start += num12;
								comObj13.Font.Bold = 0;
								comObj13.Font.Italic = 0;
								ComReleaser.Release(ref comObj13);
							}
						}
						else
						{
							comObj2.Range.Font.Italic = 0;
						}
					}
					catch
					{
					}
				}
				if (semanticRole == SemanticRole.DocumentType || semanticRole == SemanticRole.SubjectExcerpt)
				{
					try
					{
						comObj2.Range.ParagraphFormat.Borders.Enable = 0;
						foreach (Border border4 in comObj2.Borders)
						{
							try
							{
								border4.LineStyle = WdLineStyle.wdLineStyleNone;
							}
							catch
							{
							}
						}
						comObj2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						comObj2.Range.Underline = WdUnderline.wdUnderlineNone;
					}
					catch
					{
					}
				}
				if (componentStyle != null && componentStyle.UnderlineRule != UnderlineRuleType.None)
				{
					UnderlineFormatter.ApplyUnderlineRule(comObj2, componentStyle.UnderlineRule);
				}
				ParagraphFormatter.FixOrphanWord(comObj2);
				num++;
				goto IL_1163;
			}
			catch
			{
				goto IL_1163;
			}
			finally
			{
				ComReleaser.Release(ref comObj2);
			}
			IL_1163:
			if (i % 15 == 0 || i == count)
			{
				int current = 25 + (int)((float)i / (float)count * 55f);
				ReportProgress(current, 100, $"Đang chuẩn hóa đoạn {i}/{count}...");
			}
		}
		ComReleaser.Release(ref comObj);
		return num;
	}

	private static string DescribeOutcome(string component, ComponentOutcome outcome)
	{
		return outcome switch
		{
			ComponentOutcome.Rebuilt => component + ": đã tái tạo theo quy chuẩn.", 
			ComponentOutcome.Converted => component + ": đã chuyển khối thô thành bảng chuẩn.", 
			ComponentOutcome.InsertedTemplate => component + ": đã chèn khối mẫu (cần điền thông tin).", 
			ComponentOutcome.AppendixEnding => component + ": văn bản kèm theo, kết thúc bằng dấu ./.", 
			ComponentOutcome.KeptRawBlock => component + ": giữ nguyên khối hiện có.", 
			ComponentOutcome.Missing => component + ": không tìm thấy, đã giữ nguyên.", 
			_ => component + ": đã chuẩn hóa khối hiện có.", 
		};
	}

	private static string DetectPrimaryDocumentType(Document doc)
	{
		if (doc == null)
		{
			return "";
		}
		for (int i = 1; i <= Math.Min(doc.Paragraphs.Count, 50); i++)
		{
			Paragraph comObj = null;
			try
			{
				comObj = doc.Paragraphs[i];
				string text = (comObj.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				if (RegexPatterns.LegalGround.IsMatch(text))
				{
					break;
				}
				if (IsLikelyDocumentTitle(text, i, hasEnteredBodyStructure: false))
				{
					return text;
				}
				continue;
			}
			finally
			{
				ComReleaser.Release(ref comObj);
			}
		}
		return "";
	}

	private static float GetComponentSizeOrDefault(ProfileModel profile, SemanticRole role, float fallback)
	{
		try
		{
			if (profile != null)
			{
				ComponentStyle componentStyle = profile.GetComponentStyle(role);
				if (componentStyle != null && componentStyle.SizePt > 0f)
				{
					return componentStyle.SizePt;
				}
			}
		}
		catch
		{
		}
		return fallback;
	}

	private ProfileModel PrepareProfile(RulePackType? forcedRulePack, out RulePackType rulePack)
	{
		try
		{
			VFormat.Core.Personalization.UserPersonalization latest = new VFormat.Core.Serialization.JsonProfileRepository().LoadPersonalization();
			if (latest != null)
			{
				_ruleEngine.UpdatePersonalization(latest);
			}
		}
		catch
		{
		}
		rulePack = (forcedRulePack.HasValue ? forcedRulePack.Value : _ruleEngine.ActiveProfile.RulePack);
		_ruleEngine.SetActiveProfile(rulePack);
		return _ruleEngine.ActiveProfile;
	}

	private static List<Tuple<Range, string>> CaptureSymbolFontRuns(Document doc)
	{
		List<Tuple<Range, string>> list = new List<Tuple<Range, string>>();
		if (doc == null)
		{
			return list;
		}
		string[] symbolFontNames = SymbolFontNames;
		foreach (string text in symbolFontNames)
		{
			try
			{
				Range content = doc.Content;
				Find find = content.Find;
				find.ClearFormatting();
				find.Text = "";
				find.Replacement.Text = "";
				find.Format = true;
				find.MatchWildcards = false;
				find.Forward = true;
				find.Wrap = WdFindWrap.wdFindStop;
				find.Font.Name = text;
				int num = 0;
				int num2 = -1;
				while (true)
				{
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
					object Replace = Missing.Value;
					object MatchKashida = Missing.Value;
					object MatchDiacritics = Missing.Value;
					object MatchAlefHamza = Missing.Value;
					object MatchControl = Missing.Value;
					if (find.Execute(ref FindText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl) && num++ < 5000 && content.Start > num2 && content.End > content.Start)
					{
						num2 = content.Start;
						list.Add(Tuple.Create(content.Duplicate, text));
						object Direction = WdCollapseDirection.wdCollapseEnd;
						content.Collapse(ref Direction);
						continue;
					}
					break;
				}
			}
			catch
			{
			}
		}
		return list;
	}

	private static void RestoreSymbolFontRuns(Document doc, List<Tuple<Range, string>> runs)
	{
		if (doc == null || runs == null)
		{
			return;
		}
		foreach (Tuple<Range, string> run in runs)
		{
			try
			{
				Range item = run.Item1;
				if (item.End > item.Start)
				{
					item.Font.Name = run.Item2;
				}
			}
			catch
			{
			}
		}
	}

	private ComponentOutcome FormatHeaderComponent(Document doc, RulePackType rulePack, ProfileModel profile, bool requireAnchors, bool allowTemplateInsert)
	{
		if (HeaderRebuilder.IsPersonalDocument(doc))
		{
			HeaderRebuilder.FormatPersonalDocumentHeader(doc);
			return ComponentOutcome.Formatted;
		}
		Table table = HeaderRebuilder.FindExistingHeaderTable(doc);
		if (table != null)
		{
			if (HeaderRebuilder.NeedsRebuild(table, rulePack))
			{
				HeaderData headerData = ComponentDataExtractor.ExtractHeaderData(doc, table, rulePack);
				if (headerData != null && headerData.HasExtractedAgency)
				{
					HeaderRebuilder.RebuildHeaderCleanly(doc, headerData, rulePack, table, profile);
					return ComponentOutcome.Rebuilt;
				}
			}
			HeaderRebuilder.FormatExistingHeaderTable(table, rulePack, profile);
			return ComponentOutcome.Formatted;
		}
		if (HeaderRebuilder.ConvertRawHeaderToTable(doc, rulePack, profile))
		{
			return ComponentOutcome.Converted;
		}
		HeaderData headerData2 = ComponentDataExtractor.ExtractHeaderData(doc, null, rulePack);
		if (headerData2 != null && (headerData2.HasExtractedAgency || !string.IsNullOrEmpty(headerData2.DocNumber)) && (!requireAnchors || IsHeaderRebuildSafe(doc, null, rulePack)))
		{
			HeaderRebuilder.RebuildHeaderCleanly(doc, headerData2, rulePack, null, profile);
			return ComponentOutcome.Rebuilt;
		}
		if (allowTemplateInsert)
		{
			HeaderRebuilder.AutoInsertMissingHeader(doc, rulePack, profile);
			return ComponentOutcome.InsertedTemplate;
		}
		return ComponentOutcome.Missing;
	}

	private ComponentOutcome FormatSignatureComponent(Document doc, RulePackType rulePack, bool cautious, bool allowTemplateInsert)
	{
		if (SignatureRebuilder.IsDocumentAttachedToDecision(doc))
		{
			try
			{
				Range content = doc.Content;
				object Direction = WdCollapseDirection.wdCollapseEnd;
				content.Collapse(ref Direction);
				string text = (content.Paragraphs.Last.Range.Text ?? "").Trim();
				if (!text.EndsWith("./.") && !text.EndsWith("/."))
				{
					content.InsertAfter("\r\n./.");
				}
			}
			catch
			{
			}
			return ComponentOutcome.AppendixEnding;
		}
		SignatureRebuilder.CleanDuplicatePlaceholderSignatures(doc);
		Table table = SignatureRebuilder.FindExistingSignatureTable(doc);
		ComponentOutcome result;
		if (table != null)
		{
			result = ComponentOutcome.Formatted;
			bool flag = false;
			if (table.Rows.Count > 1)
			{
				SignatureData signatureData = ComponentDataExtractor.ExtractSignatureData(doc, table, rulePack);
				if (signatureData != null && (signatureData.Recipients.Count > 0 || !string.IsNullOrEmpty(signatureData.Title) || !string.IsNullOrEmpty(signatureData.SignerName)))
				{
					SignatureRebuilder.RebuildSignatureCleanly(doc, signatureData, rulePack, table, null, _ruleEngine.ActiveProfile);
					result = ComponentOutcome.Rebuilt;
					flag = true;
				}
			}
			if (!flag)
			{
				SignatureRebuilder.FormatExistingSignatureTable(table, rulePack, _ruleEngine.ActiveProfile);
			}
		}
		else if (SignatureRebuilder.ConvertRawSignatureToTable(doc, rulePack, _ruleEngine.ActiveProfile))
		{
			result = ComponentOutcome.Converted;
		}
		else if (SignatureRebuilder.IsDocEndingWithSignatureBlock(doc))
		{
			SemanticComponentFormatter.Apply(doc, _ruleEngine.ActiveProfile, null, reflowTables: false, signatureOnly: true);
			result = ComponentOutcome.KeptRawBlock;
		}
		else
		{
			SignatureData signatureData2 = (cautious ? null : ComponentDataExtractor.ExtractSignatureData(doc, null, rulePack));
			if (signatureData2 != null && (!string.IsNullOrEmpty(signatureData2.SignerName) || !string.IsNullOrEmpty(signatureData2.Title) || signatureData2.Recipients.Count > 0))
			{
				SignatureRebuilder.RebuildSignatureCleanly(doc, signatureData2, rulePack, null, null, _ruleEngine.ActiveProfile);
				result = ComponentOutcome.Rebuilt;
			}
			else if (allowTemplateInsert)
			{
				SignatureRebuilder.AutoInsertMissingSignature(doc, rulePack, _ruleEngine.ActiveProfile);
				result = ComponentOutcome.InsertedTemplate;
			}
			else
			{
				result = ComponentOutcome.Missing;
			}
		}
		SignatureRebuilder.CleanDuplicatePlaceholderSignatures(doc);
		return result;
	}

	private static void RemoveStrayLinesInTitleZone(Document doc, int titleParagraphIndex)
	{
		int start;
		int end;
		try
		{
			start = doc.Paragraphs[titleParagraphIndex].Range.Start;
			int index = Math.Min(doc.Paragraphs.Count, titleParagraphIndex + 7);
			end = doc.Paragraphs[index].Range.End;
		}
		catch
		{
			return;
		}
		try
		{
			for (int num = doc.Shapes.Count; num >= 1; num--)
			{
				try
				{
					Shapes shapes = doc.Shapes;
					object Index = num;
					dynamic val = shapes[ref Index];
					string text = "";
					try
					{
						text = (string)val.Name;
					}
					catch
					{
					}
					if (string.IsNullOrEmpty(text) || !text.StartsWith("VFormat_"))
					{
						int num2 = (int)val.Type;
						_ = (float)val.Height;
						if (num2 == 9)
						{
							Range range = null;
							try
							{
								range = val.Anchor;
							}
							catch
							{
							}
							if (range != null && range.Tables.Count <= 0 && range.Start >= start && range.Start <= end)
							{
								val.Delete();
							}
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
		try
		{
			object Start = start;
			object End = end;
			Range range2 = doc.Range(ref Start, ref End);
			for (int num3 = range2.InlineShapes.Count; num3 >= 1; num3--)
			{
				try
				{
					InlineShape inlineShape = range2.InlineShapes[num3];
					if (inlineShape.Range.Tables.Count <= 0 && inlineShape.Type == WdInlineShapeType.wdInlineShapeHorizontalLine)
					{
						inlineShape.Delete();
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

	public FormattingResult ManualFormatHeader(Document doc, RulePackType? forcedRulePack = null, bool allowTemplateInsert = false)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		FormattingResult formattingResult = CheckLegacyEncoding(doc, forcedRulePack);
		if (formattingResult != null)
		{
			return formattingResult;
		}
		EnsureEditablePrintView(doc);
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application application = doc.Application;
		try
		{
			using (new ScreenFreezeScope(application))
			{
				using (new SafeUndoScope(application, "OfficeFormat Chuẩn Hóa Tiêu Đề Đầu Trang"))
				{
					using (new SymbolFontScope(doc))
					{
						ProfileModel profile = PrepareProfile(forcedRulePack, out var rulePack);
						ComponentOutcome componentOutcome = FormatHeaderComponent(doc, rulePack, profile, requireAnchors: true, allowTemplateInsert);
						stopwatch.Stop();
						if (componentOutcome == ComponentOutcome.Missing)
						{
							return FormattingResult.NotApplied(rulePack, "Không tìm thấy đủ thành phần tiêu đề (Quốc hiệu/Tiêu đề Đảng, tên cơ quan, số ký hiệu) để chuẩn hóa an toàn. OfficeFormat đã giữ nguyên văn bản.", canInsertTemplate: true);
						}
						FormattingResult formattingResult2 = FormattingResult.Succeeded(rulePack, 0, 1, stopwatch.ElapsedMilliseconds);
						formattingResult2.HeaderStatus = DescribeOutcome("Tiêu đề", componentOutcome);
						formattingResult2.Message = formattingResult2.HeaderStatus;
						return formattingResult2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
	}

	public FormattingResult ManualFormatTitle(Document doc, RulePackType? forcedRulePack = null)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		FormattingResult formattingResult = CheckLegacyEncoding(doc, forcedRulePack);
		if (formattingResult != null)
		{
			return formattingResult;
		}
		EnsureEditablePrintView(doc);
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application application = doc.Application;
		try
		{
			using (new ScreenFreezeScope(application))
			{
				using (new SafeUndoScope(application, "OfficeFormat Chuẩn Hóa Tên Loại & Trích Yếu"))
				{
					using (new SymbolFontScope(doc))
					{
						PrepareProfile(forcedRulePack, out var rulePack);
						bool flag = StandardizeTitleAndAbstract(doc, rulePack);
						stopwatch.Stop();
						if (!flag)
						{
							return FormattingResult.NotApplied(rulePack, "Không tìm thấy Tên loại văn bản (QUYẾT ĐỊNH, BÁO CÁO, KẾ HOẠCH...) trong phần đầu văn bản. Nếu đây là Công văn, trích yếu \"V/v...\" được chuẩn hóa bằng nút Tiêu đề Đầu trang.", canInsertTemplate: false);
						}
						return FormattingResult.Succeeded(rulePack, 1, 0, stopwatch.ElapsedMilliseconds);
					}
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
	}

	public FormattingResult ManualFormatBody(Document doc, RulePackType? forcedRulePack = null)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		FormattingResult formattingResult = CheckLegacyEncoding(doc, forcedRulePack);
		if (formattingResult != null)
		{
			return formattingResult;
		}
		EnsureEditablePrintView(doc);
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application application = doc.Application;
		try
		{
			using (new ScreenFreezeScope(application))
			{
				using (new SafeUndoScope(application, "OfficeFormat Chuẩn Hóa Đề Mục & Thân Bài"))
				{
					using (new SymbolFontScope(doc))
					{
						RulePackType rulePackType = forcedRulePack ?? _ruleEngine.ActiveProfile.RulePack;
						try
						{
							JsonProfileRepository jsonProfileRepository = new JsonProfileRepository();
							UserPersonalization userPersonalization = jsonProfileRepository.LoadPersonalization();
							if (userPersonalization != null)
							{
								_ruleEngine.UpdatePersonalization(userPersonalization);
							}
						}
						catch
						{
						}
						_ruleEngine.SetActiveProfile(rulePackType);
						ProfileModel activeProfile = _ruleEngine.ActiveProfile;
						int paragraphs = FormatAllParagraphs(doc, activeProfile, (int cur, int tot) =>
						{
							int current = ((tot > 0) ? ((int)Math.Round((double)cur / (double)tot * 100.0)) : cur);
							ReportProgress(current, 100, $"Đang chuẩn hóa đoạn {cur}/{tot}...");
						});
						stopwatch.Stop();
						return FormattingResult.Succeeded(rulePackType, paragraphs, 0, stopwatch.ElapsedMilliseconds);
					}
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
	}

	public FormattingResult ManualFormatSignature(Document doc, RulePackType? forcedRulePack = null, bool allowTemplateInsert = false)
	{
		if (doc == null)
		{
			return FormattingResult.Failed(new ArgumentNullException("doc"));
		}
		FormattingResult formattingResult = CheckLegacyEncoding(doc, forcedRulePack);
		if (formattingResult != null)
		{
			return formattingResult;
		}
		EnsureEditablePrintView(doc);
		Stopwatch stopwatch = Stopwatch.StartNew();
		Application application = doc.Application;
		try
		{
			using (new ScreenFreezeScope(application))
			{
				using (new SafeUndoScope(application, "OfficeFormat Chuẩn Hóa Chữ Ký & Nơi Nhận"))
				{
					using (new SymbolFontScope(doc))
					{
						PrepareProfile(forcedRulePack, out var rulePack);
						ComponentOutcome componentOutcome = FormatSignatureComponent(doc, rulePack, cautious: false, allowTemplateInsert);
						stopwatch.Stop();
						if (componentOutcome == ComponentOutcome.Missing)
						{
							return FormattingResult.NotApplied(rulePack, "Không tìm thấy khối Nơi nhận / Chữ ký ở cuối văn bản. OfficeFormat đã giữ nguyên văn bản.", canInsertTemplate: true);
						}
						FormattingResult formattingResult2 = FormattingResult.Succeeded(rulePack, 0, 1, stopwatch.ElapsedMilliseconds);
						formattingResult2.SignatureStatus = DescribeOutcome("Chữ ký & Nơi nhận", componentOutcome);
						formattingResult2.Message = formattingResult2.SignatureStatus;
						return formattingResult2;
					}
				}
			}
		}
		catch (Exception ex)
		{
			return FormattingResult.Failed(ex);
		}
	}

	private static Range InsertSeparatorParagraph(Document doc, Paragraph paragraph, string separator)
	{
		int num = Math.Max(paragraph.Range.Start, paragraph.Range.End - 1);
		object Start = num;
		object End = num;
		Range comObj = doc.Range(ref Start, ref End);
		comObj.Text = "\r" + separator;
		ComReleaser.Release(ref comObj);
		object Start2 = num + 1;
		object End2 = num + 1 + separator.Length;
		return doc.Range(ref Start2, ref End2);
	}

	private static bool IsAbstractBodyBoundary(Paragraph paragraph, string text)
	{
		string text2 = (text ?? "").Trim();
		if (text2.Length == 0)
		{
			return false;
		}
		if (char.IsLower(text2[0]))
		{
			return false;
		}
		string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
		if (RegexPatterns.ArticleTitle.IsMatch(text2) || RegexPatterns.ChapterHeader.IsMatch(text2) || RegexPatterns.SectionHeader.IsMatch(text2) || Regex.IsMatch(text3, "^(DIEU\\s+\\d+|PHAN\\s+[IVXLCDM\\d]+)\\b"))
		{
			return true;
		}
		if (text3.StartsWith("CAN CU ") || text3.StartsWith("CAN CU:") || text3.StartsWith("KINH GUI") || text3.StartsWith("NOI NHAN") || text3.StartsWith("TM.") || text3.StartsWith("KT.") || text3.StartsWith("TL.") || text3.StartsWith("TUQ.") || SignatureRebuilder.IsSignatureOrRecipientBlock(text2))
		{
			return true;
		}
		bool flag = paragraph != null && paragraph.Range.ParagraphFormat.Alignment == WdParagraphAlignment.wdAlignParagraphCenter;
		bool flag2 = paragraph != null && paragraph.Range.Font.Bold == -1;
		if (!flag && !flag2)
		{
			return IsBodyStarter(text2);
		}
		return false;
	}

	private bool StandardizeTitleAndAbstract(Document doc, RulePackType rulePack, bool preserveLayout = false)
	{
		if (doc == null)
		{
			return false;
		}
		int count = doc.Paragraphs.Count;
		int num = 0;
		bool flag = RulePackTraits.UsesTextSeparatorBelowExcerpt(rulePack);
		string excerptSeparatorText = RulePackTraits.GetExcerptSeparatorText(rulePack);
		int bold = (RulePackTraits.ExcerptSeparatorBold(rulePack) ? (-1) : 0);
		ComponentStyle componentStyle = _ruleEngine.ActiveProfile.GetComponentStyle(SemanticRole.DocumentType);
		ComponentStyle componentStyle2 = _ruleEngine.ActiveProfile.GetComponentStyle(SemanticRole.SubjectExcerpt);
		float sizePt = componentStyle.SizePt;
		float sizePt2 = componentStyle2.SizePt;
		for (int i = 1; i <= count; i++)
		{
			OperationProgress.Step();
			Paragraph paragraph = doc.Paragraphs[i];
			if (paragraph.Range.Tables.Count != 0)
			{
				continue;
			}
			if (++num > 25)
			{
				break;
			}
			string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			if (!IsLikelyDocumentTitle(text, num, hasEnteredBodyStructure: false))
			{
				continue;
			}
			RemoveStrayLinesInTitleZone(doc, i);
			paragraph = doc.Paragraphs[i];
			paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph.Range.ParagraphFormat.FirstLineIndent = 0f;
			paragraph.Range.ParagraphFormat.LeftIndent = 0f;
			paragraph.Range.ParagraphFormat.RightIndent = 0f;
			paragraph.Range.ParagraphFormat.SpaceBefore = componentStyle.SpacingBeforePt.GetValueOrDefault();
			paragraph.Range.ParagraphFormat.SpaceAfter = componentStyle.SpacingAfterPt.GetValueOrDefault();
			paragraph.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			paragraph.Range.Font.Name = "Times New Roman";
			paragraph.Range.Font.Size = sizePt;
			paragraph.Range.Font.Bold = ((componentStyle == null) ? (-1) : (componentStyle.Bold ? (-1) : 0));
			paragraph.Range.Font.Italic = ((componentStyle != null) ? (componentStyle.Italic ? (-1) : 0) : 0);
			paragraph.Range.Underline = WdUnderline.wdUnderlineNone;
			paragraph.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			try
			{
				foreach (Range word in paragraph.Range.Words)
				{
					try
					{
						word.Underline = WdUnderline.wdUnderlineNone;
						word.Font.Underline = WdUnderline.wdUnderlineNone;
					}
					catch
					{
					}
				}
			}
			catch
			{
			}
			try
			{
				paragraph.Range.ParagraphFormat.Borders.Enable = 0;
				foreach (Border border4 in paragraph.Borders)
				{
					try
					{
						border4.LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
			}
			catch
			{
			}
			int num2 = Math.Min(doc.Paragraphs.Count, i + 6);
			int num3 = 0;
			int num4 = i;
			bool flag2 = false;
			int num5 = -1;
			for (int j = i + 1; j <= num2; j++)
			{
				Paragraph paragraph2 = doc.Paragraphs[j];
				if (paragraph2.Range.Tables.Count > 0)
				{
					break;
				}
				string text2 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (text2.Length == 0)
				{
					continue;
				}
				if (HeaderRebuilder.IsSeparatorLine(text2))
				{
					if (num3 == 0)
					{
						try
						{
							Range range2 = paragraph2.Range;
							object Unit = Missing.Value;
							object Count = Missing.Value;
							range2.Delete(ref Unit, ref Count);
							j--;
							num2--;
						}
						catch
						{
							goto IL_053d;
						}
						continue;
					}
					bool flag3 = false;
					if (j + 1 <= doc.Paragraphs.Count)
					{
						Paragraph paragraph3 = doc.Paragraphs[j + 1];
						string text3 = (paragraph3.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
							.Trim();
						string text4 = ConvertToUnsigned(text3).ToUpperInvariant();
						if (text3.Length > 0 && !HeaderRebuilder.IsSeparatorLine(text3) && !RegexPatterns.LegalGround.IsMatch(text3) && !text4.StartsWith("CAN CU") && !IsRomanNumeralHeading(text3) && !IsLetterHeading(text3) && !IsClauseHeading(text3) && !text3.EndsWith(":") && !text4.EndsWith("NHU SAU:") && !IsAbstractBodyBoundary(paragraph3, text3))
						{
							flag3 = true;
						}
					}
					if (flag3)
					{
						try
						{
							Range range3 = paragraph2.Range;
							object Unit2 = Missing.Value;
							object Count2 = Missing.Value;
							range3.Delete(ref Unit2, ref Count2);
							j--;
							num2--;
						}
						catch
						{
							goto IL_053d;
						}
						continue;
					}
					flag2 = true;
					num5 = j;
					break;
				}
				goto IL_053d;
				IL_053d:
				string text5 = ConvertToUnsigned(text2).ToUpperInvariant();
				if (!char.IsLower(text2[0]) && (RegexPatterns.LegalGround.IsMatch(text2) || text5.StartsWith("CAN CU") || IsRomanNumeralHeading(text2) || IsLetterHeading(text2) || IsClauseHeading(text2) || text2.EndsWith(":") || text5.EndsWith("NHU SAU:") || text5.EndsWith("NHU SAU :") || text5.StartsWith("NOI NHAN") || text5.StartsWith("KINH GUI") || text5.StartsWith("TM.") || text5.StartsWith("KT.") || SignatureRebuilder.IsSignatureOrRecipientBlock(text2) || IsAbstractBodyBoundary(paragraph2, text2) || num3 >= 4))
				{
					break;
				}
				paragraph2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				paragraph2.Range.ParagraphFormat.FirstLineIndent = 0f;
				paragraph2.Range.ParagraphFormat.LeftIndent = 0f;
				paragraph2.Range.ParagraphFormat.RightIndent = 0f;
				paragraph2.Range.ParagraphFormat.SpaceBefore = componentStyle2.SpacingBeforePt.GetValueOrDefault();
				paragraph2.Range.ParagraphFormat.SpaceAfter = componentStyle2.SpacingAfterPt.GetValueOrDefault();
				paragraph2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				paragraph2.Range.Font.Name = "Times New Roman";
				paragraph2.Range.Font.Size = sizePt2;
				paragraph2.Range.Underline = WdUnderline.wdUnderlineNone;
				paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				try
				{
					foreach (Range word2 in paragraph2.Range.Words)
					{
						try
						{
							word2.Underline = WdUnderline.wdUnderlineNone;
							word2.Font.Underline = WdUnderline.wdUnderlineNone;
						}
						catch
						{
						}
					}
				}
				catch
				{
				}
				try
				{
					paragraph2.Range.ParagraphFormat.Borders.Enable = 0;
					foreach (Border border5 in paragraph2.Borders)
					{
						try
						{
							border5.LineStyle = WdLineStyle.wdLineStyleNone;
						}
						catch
						{
						}
					}
				}
				catch
				{
				}
				if (text2.StartsWith("(") && (text5.Contains("BAN HANH") || text5.Contains("KEM THEO")))
				{
					paragraph2.Range.Font.Bold = 0;
					paragraph2.Range.Font.Italic = -1;
				}
				else
				{
					paragraph2.Range.Font.Bold = ((componentStyle2 == null) ? (-1) : (componentStyle2.Bold ? (-1) : 0));
					paragraph2.Range.Font.Italic = ((componentStyle2 != null) ? (componentStyle2.Italic ? (-1) : 0) : 0);
				}
				num3++;
				num4 = j;
			}
			if (flag)
			{
				if (flag2 && num5 > 0)
				{
					Paragraph paragraph4 = doc.Paragraphs[num5];
					paragraph4.Range.Text = excerptSeparatorText + "\r";
					paragraph4.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph4.Range.ParagraphFormat.FirstLineIndent = 0f;
					paragraph4.Range.ParagraphFormat.LeftIndent = 0f;
					paragraph4.Range.ParagraphFormat.RightIndent = 0f;
					paragraph4.Range.ParagraphFormat.SpaceBefore = 2f;
					paragraph4.Range.ParagraphFormat.SpaceAfter = 6f;
					paragraph4.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					paragraph4.Range.Font.Name = "Times New Roman";
					paragraph4.Range.Font.Size = 14f;
					paragraph4.Range.Font.Bold = bold;
					paragraph4.Range.Font.Italic = 0;
					paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph4.Range.ParagraphFormat.Borders.Enable = 0;
						foreach (Border border6 in paragraph4.Borders)
						{
							try
							{
								border6.LineStyle = WdLineStyle.wdLineStyleNone;
							}
							catch
							{
							}
						}
					}
					catch
					{
					}
					if (num5 + 1 <= doc.Paragraphs.Count)
					{
						Paragraph paragraph5 = doc.Paragraphs[num5 + 1];
						string text6 = (paragraph5.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
							.Trim();
						if (paragraph5.Range.Tables.Count == 0 && HeaderRebuilder.IsSeparatorLine(text6))
						{
							try
							{
								Range range5 = paragraph5.Range;
								object Unit = Missing.Value;
								object Count = Missing.Value;
								range5.Delete(ref Unit, ref Count);
							}
							catch
							{
							}
						}
					}
				}
				else if (num4 > i)
				{
					Paragraph paragraph6 = doc.Paragraphs[num4];
					Range range6 = InsertSeparatorParagraph(doc, paragraph6, excerptSeparatorText);
					range6.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					range6.ParagraphFormat.FirstLineIndent = 0f;
					range6.ParagraphFormat.LeftIndent = 0f;
					range6.ParagraphFormat.RightIndent = 0f;
					range6.ParagraphFormat.SpaceBefore = 2f;
					range6.ParagraphFormat.SpaceAfter = 6f;
					range6.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					range6.Font.Name = "Times New Roman";
					range6.Font.Size = 14f;
					range6.Font.Bold = bold;
					range6.Font.Italic = 0;
					range6.Font.Underline = WdUnderline.wdUnderlineNone;
				}
			}
			else if (num3 > 0 && num4 >= 1 && num4 <= doc.Paragraphs.Count)
			{
				try
				{
					Paragraph paragraph7 = doc.Paragraphs[num4];
					paragraph7.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					paragraph7.Range.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph7.Range.ParagraphFormat.Borders.Enable = 0;
					}
					catch
					{
					}
					try
					{
						paragraph7.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					string text7 = (paragraph7.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
						.Trim();
					int targetCharCount = Math.Max(7, Math.Min(16, (int)Math.Round((double)text7.Length * 0.35)));
					string text8 = UnderlineFormatter.CreateSolidSeparator(targetCharCount);
					int num6 = -1;
					for (int k = num4 + 1; k <= Math.Min(doc.Paragraphs.Count, num4 + 3) && doc.Paragraphs[k].Range.Tables.Count <= 0; k++)
					{
						string text9 = (doc.Paragraphs[k].Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
							.Trim();
						if (HeaderRebuilder.IsSeparatorLine(text9))
						{
							num6 = k;
							break;
						}
						if (text9.Length > 0)
						{
							break;
						}
					}
					if (num6 > 0 || (flag2 && num5 > num4))
					{
						int num7 = ((num6 > 0) ? num6 : num5);
						Paragraph paragraph8 = doc.Paragraphs[num7];
						paragraph8.Range.Text = text8 + "\r";
						paragraph8.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
						paragraph8.Range.ParagraphFormat.FirstLineIndent = 0f;
						paragraph8.Range.ParagraphFormat.LeftIndent = 0f;
						paragraph8.Range.ParagraphFormat.RightIndent = 0f;
						paragraph8.Range.ParagraphFormat.SpaceBefore = 2f;
						paragraph8.Range.ParagraphFormat.SpaceAfter = 6f;
						paragraph8.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
						paragraph8.Range.Font.Name = "Times New Roman";
						paragraph8.Range.Font.Size = 10f;
						paragraph8.Range.Font.Bold = -1;
						paragraph8.Range.Font.Italic = 0;
						paragraph8.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						try
						{
							paragraph8.Range.ParagraphFormat.Borders.Enable = 0;
						}
						catch
						{
						}
						while (num7 + 1 <= doc.Paragraphs.Count)
						{
							Paragraph paragraph9 = doc.Paragraphs[num7 + 1];
							string text10 = (paragraph9.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
								.Trim();
							if (paragraph9.Range.Tables.Count == 0 && HeaderRebuilder.IsSeparatorLine(text10))
							{
								try
								{
									Range range7 = paragraph9.Range;
									object Unit = Missing.Value;
									object Count = Missing.Value;
									range7.Delete(ref Unit, ref Count);
								}
								catch
								{
									break;
								}
								continue;
							}
							break;
						}
					}
					else
					{
						Range range8 = InsertSeparatorParagraph(doc, paragraph7, text8);
						range8.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
						range8.ParagraphFormat.FirstLineIndent = 0f;
						range8.ParagraphFormat.LeftIndent = 0f;
						range8.ParagraphFormat.RightIndent = 0f;
						range8.ParagraphFormat.SpaceBefore = 2f;
						range8.ParagraphFormat.SpaceAfter = 6f;
						range8.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
						range8.Font.Name = "Times New Roman";
						range8.Font.Size = 10f;
						range8.Font.Bold = -1;
						range8.Font.Italic = 0;
						range8.Font.Underline = WdUnderline.wdUnderlineNone;
						try
						{
							range8.ParagraphFormat.Borders.Enable = 0;
						}
						catch
						{
						}
						if (num4 + 2 <= doc.Paragraphs.Count)
						{
							Paragraph paragraph10 = doc.Paragraphs[num4 + 2];
							string text11 = (paragraph10.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
								.Trim();
							if (paragraph10.Range.Tables.Count == 0 && HeaderRebuilder.IsSeparatorLine(text11))
							{
								try
								{
									Range range9 = paragraph10.Range;
									object Unit = Missing.Value;
									object Count = Missing.Value;
									range9.Delete(ref Unit, ref Count);
								}
								catch
								{
								}
							}
						}
					}
				}
				catch
				{
				}
			}
			else if (num3 == 0 && flag2 && num5 > 0 && num5 <= doc.Paragraphs.Count)
			{
				try
				{
					Range range10 = doc.Paragraphs[num5].Range;
					object Unit = Missing.Value;
					object Count = Missing.Value;
					range10.Delete(ref Unit, ref Count);
				}
				catch
				{
				}
			}
			return true;
		}
		return false;
	}

	public static bool IsLikelyLegacyEncoded(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		int num = 0;
		int num2 = 0;
		foreach (char c in text)
		{
			if (char.IsLetter(c))
			{
				num2++;
			}
			if ("µ\u00b8¹¶·\u00a8©ª«¬®¦¤¢£¡Þþ−§".IndexOf(c) >= 0)
			{
				num++;
			}
		}
		if (num2 < 50)
		{
			return false;
		}
		if (num >= 15 && num * 100 >= num2)
		{
			return true;
		}
		int count = Regex.Matches(text, "[aeiouyAEIOUYôơưâăêÔƠƯÂĂÊ][øùûõïäöüëØÙÛÕÏ]").Count;
		if (count >= 15)
		{
			return count * 100 >= num2;
		}
		return false;
	}

	private FormattingResult CheckLegacyEncoding(Document doc, RulePackType? forcedRulePack)
	{
		string text = null;
		try
		{
			object Start = 0;
			object End = Math.Min(doc.Content.End, 20000);
			Range range = doc.Range(ref Start, ref End);
			text = range.Text;
		}
		catch
		{
		}
		if (!IsLikelyLegacyEncoded(text))
		{
			return null;
		}
		return FormattingResult.NotApplied(forcedRulePack ?? RulePackType.ND30_2020, "Văn bản đang dùng bảng mã cũ (TCVN3/.VnTime hoặc VNI). OfficeFormat đã giữ nguyên văn bản để tránh lỗi font. Hãy chuyển mã sang Unicode trước (nút Chuyển mã Unicode hoặc Unikey), sau đó chuẩn hóa lại.", canInsertTemplate: false);
	}

	private string GetDocumentSampleText(Document doc)
	{
		try
		{
			object Start = 0;
			object End = Math.Min(doc.Content.End, 2500);
			Range comObj = doc.Range(ref Start, ref End);
			string text = comObj.Text;
			ComReleaser.Release(ref comObj);
			return text ?? "";
		}
		catch
		{
			return "";
		}
	}
}
