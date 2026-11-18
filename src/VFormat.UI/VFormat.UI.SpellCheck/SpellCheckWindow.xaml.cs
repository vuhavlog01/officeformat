using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Threading;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Enums;
using VFormat.Core.SpellCheck;
using VFormat.Core.SpellCheck.Models;
using VFormat.WordEngine.SpellCheck;

namespace VFormat.UI.SpellCheck;

public partial class SpellCheckWindow : System.Windows.Window, IComponentConnector
{
	private readonly Microsoft.Office.Interop.Word.Application _app;

	private readonly Document _doc;

	private readonly bool _isSpacingMode;

	private WordSpellCheckController _controller;

	private List<SpellIssueViewModel> _allViewModels;

	private SpellCheckProfile _initialProfile;

	private readonly DispatcherTimer _scanTimer;

	private int _nextParagraph;

	private int _paragraphCount;

	private SpellCheckProfile _scanProfile;

	private SpellCheckService _sharedSpellService;

	private bool _selectingInitialProfile;

	private CancellationTokenSource _scanCancellation;

	private int _scanGeneration;

	public SpellCheckWindow()
		: this(null, null, SpellCheckProfile.Administrative, spacingCheckMode: false)
	{
	}

	public SpellCheckWindow(Microsoft.Office.Interop.Word.Application app, Document doc)
		: this(app, doc, SpellCheckProfile.Administrative, spacingCheckMode: false)
	{
	}

	public SpellCheckWindow(Microsoft.Office.Interop.Word.Application app, Document doc, RulePackType rulePack)
		: this(app, doc, MapRulePackToProfile(rulePack), spacingCheckMode: false)
	{
	}

	public SpellCheckWindow(Microsoft.Office.Interop.Word.Application app, Document doc, RulePackType rulePack, bool spacingCheckMode)
		: this(app, doc, MapRulePackToProfile(rulePack), spacingCheckMode)
	{
	}

	public SpellCheckWindow(Microsoft.Office.Interop.Word.Application app, Document doc, SpellCheckProfile initialProfile, bool spacingCheckMode = false)
	{
		InitializeComponent();
		_app = app;
		_doc = doc;
		_initialProfile = initialProfile;
		_isSpacingMode = spacingCheckMode;
		_controller = ((_app != null) ? new WordSpellCheckController((Microsoft.Office.Interop.Word.Application)_app) : null);
		_allViewModels = new List<SpellIssueViewModel>();
		if (_isSpacingMode)
		{
			ConfigureSpacingModeUi();
		}
		_scanTimer = new DispatcherTimer(DispatcherPriority.Background);
		_scanTimer.Interval = TimeSpan.FromMilliseconds(1.0);
		_scanTimer.Tick += ScanTimer_Tick;
		Loaded += SpellCheckWindow_Loaded;
		ContentRendered += SpellCheckWindow_ContentRendered;
		Closed += (object sender, EventArgs e) =>
		{
			CancelScan(showStatus: false);
		};
	}

	private void ConfigureSpacingModeUi()
	{
		Title = "Kiểm Tra Khoảng Trắng, Dấu Nháy & Dấu Gạch - OfficeFormat";
		if (txtHeaderTitle != null)
		{
			txtHeaderTitle.Text = "KIỂM TRA KHOẢNG TRẮNG, DẤU NHÁY (\") & DẤU GẠCH (–)";
		}
		if (txtHeaderSubtitle != null)
		{
			txtHeaderSubtitle.Text = "Rà soát khoảng trắng thừa, dấu cách quanh dấu câu/ngoặc, đổi dấu nháy \" sang “ và dấu gạch – sang - trực tiếp từng lỗi.";
		}
		if (pnlProfileSelector != null)
		{
			pnlProfileSelector.Visibility = Visibility.Collapsed;
		}
		if (txtColOriginalHeader != null)
		{
			txtColOriginalHeader.Text = "Ký tự / Khoảng trắng lỗi";
		}
		if (txtNoIssuesTitle != null)
		{
			txtNoIssuesTitle.Text = "Tuyệt vời! Không phát hiện lỗi khoảng trắng, dấu nháy hay dấu gạch nào.";
		}
		if (txtNoIssuesSubtitle != null)
		{
			txtNoIssuesSubtitle.Text = "Khoảng cách giữa các từ, dấu câu, dấu nháy kép và dấu gạch nối trong văn bản đều đã chuẩn xác.";
		}
		if (radFilterSpelling != null)
		{
			radFilterSpelling.Content = "Khoảng trắng: 0";
		}
		if (radFilterCap != null)
		{
			radFilterCap.Content = "Dấu câu / Ngoặc: 0";
		}
		if (radFilterAbbr != null)
		{
			radFilterAbbr.Content = "Dấu nháy (\"): 0";
		}
		if (radFilterVariant != null)
		{
			radFilterVariant.Content = "Dấu gạch (–): 0";
		}
		if (btnReplace != null)
		{
			btnReplace.Content = "✓ Sửa lỗi này";
		}
		if (btnReplaceAll != null)
		{
			btnReplaceAll.Content = "⚡ Sửa tất cả lỗi";
		}
		if (btnAddDictionary != null)
		{
			btnAddDictionary.Visibility = Visibility.Collapsed;
		}
	}

	public static SpellCheckProfile MapRulePackToProfile(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.TT68_2023_BQP => SpellCheckProfile.Military, 
			RulePackType.HD05_VPTW => SpellCheckProfile.Party, 
			_ => SpellCheckProfile.Administrative, 
		};
	}

	private void SpellCheckWindow_Loaded(object sender, RoutedEventArgs e)
	{
		_selectingInitialProfile = true;
		try
		{
			SelectProfileItem(_initialProfile);
		}
		finally
		{
			_selectingInitialProfile = false;
		}
	}

	private void SpellCheckWindow_ContentRendered(object sender, EventArgs e)
	{
		ContentRendered -= SpellCheckWindow_ContentRendered;
		RunScan();
	}

	private void SelectProfileItem(SpellCheckProfile profile)
	{
		if (cboProfile == null)
		{
			return;
		}
		string text = "Administrative";
		text = profile switch
		{
			SpellCheckProfile.Military => "Military", 
			SpellCheckProfile.Party => "Party", 
			SpellCheckProfile.LegalNormative => "LegalNormative", 
			SpellCheckProfile.EducationOrthography => "EducationOrthography", 
			_ => "Administrative", 
		};
		foreach (object item in (IEnumerable)cboProfile.Items)
		{
			if (item is ComboBoxItem comboBoxItem && (string)comboBoxItem.Tag == text)
			{
				cboProfile.SelectedItem = comboBoxItem;
				break;
			}
		}
	}

	private SpellCheckProfile GetCurrentProfile()
	{
		if (cboProfile != null && cboProfile.SelectedItem is ComboBoxItem)
		{
			ComboBoxItem comboBoxItem = (ComboBoxItem)cboProfile.SelectedItem;
			switch (comboBoxItem.Tag as string)
			{
			case "Military":
				return SpellCheckProfile.Military;
			case "Party":
				return SpellCheckProfile.Party;
			case "LegalNormative":
				return SpellCheckProfile.LegalNormative;
			case "EducationOrthography":
				return SpellCheckProfile.EducationOrthography;
			}
		}
		return SpellCheckProfile.Administrative;
	}

	private void RunScan()
	{
		CancelScan(showStatus: false);
		if (_doc == null)
		{
			lblFooterStatus.Text = "Không có tài liệu Word nào đang mở.";
			return;
		}
		_controller = new WordSpellCheckController((Microsoft.Office.Interop.Word.Application)_app, _sharedSpellService);
		_scanCancellation = new CancellationTokenSource();
		_allViewModels.Clear();
		lstIssues.ItemsSource = null;
		pnlNoIssues.Visibility = Visibility.Collapsed;
		pnlDetail.Visibility = Visibility.Collapsed;
		_scanProfile = GetCurrentProfile();
		_nextParagraph = 0;
		_paragraphCount = 0;
		barScanProgress.Value = 0.0;
		barScanProgress.Maximum = 1.0;
		barScanProgress.Visibility = Visibility.Visible;
		btnCancelScan.Visibility = Visibility.Visible;
		lblFooterStatus.Text = (_isSpacingMode ? "Đang chuẩn bị quét khoảng trắng, dấu nháy & dấu gạch..." : "Đang chuẩn bị quét tài liệu...");
		_scanTimer.Start();
	}

	private void ScanTimer_Tick(object sender, EventArgs e)
	{
		try
		{
			if (_nextParagraph == 0)
			{
				_paragraphCount = _controller.BeginScan((Document)_doc);
				barScanProgress.Maximum = Math.Max(1, _paragraphCount);
				_nextParagraph = 1;
			}
			if (_nextParagraph <= _paragraphCount)
			{
				SpellParagraphSnapshot snapshot = _controller.CaptureParagraph((Document)_doc, _nextParagraph);
				WordSpellCheckController controller = _controller;
				SpellCheckProfile profile = _scanProfile;
				bool isSpacingMode = _isSpacingMode;
				CancellationToken token = _scanCancellation.Token;
				int generation = _scanGeneration;
				_scanTimer.Stop();
				System.Threading.Tasks.Task.Factory.StartNew(() =>
				{
					token.ThrowIfCancellationRequested();
					return isSpacingMode
						? controller.AnalyzeSpacingSnapshot(snapshot, token)
						: controller.AnalyzeSnapshot(snapshot, profile, token);
				}, token, TaskCreationOptions.None, TaskScheduler.Default).ContinueWith((Task<List<WordSpellIssueContext>> task) =>
				{
					if (task.IsFaulted)
					{
						_ = task.Exception;
					}
					Dispatcher.BeginInvoke((Action)(() =>
					{
						if (generation == _scanGeneration && IsLoaded)
						{
							if (task.IsCanceled)
							{
								CancelScan(showStatus: true);
							}
							else if (task.IsFaulted)
							{
								FailScan((task.Exception != null && task.Exception.InnerException != null) ? task.Exception.InnerException.Message : "Lỗi phân tích không xác định");
							}
							else
							{
								controller.AddIssues(task.Result);
								if (_sharedSpellService == null && controller.SpellService != null)
								{
									_sharedSpellService = controller.SpellService;
								}
								barScanProgress.Value = _nextParagraph;
								lblFooterStatus.Text = $"Đang quét đoạn {_nextParagraph}/{_paragraphCount}...";
								_nextParagraph++;
								_scanTimer.Start();
							}
						}
					}), DispatcherPriority.Background);
				});
			}
			else
			{
				_scanTimer.Stop();
				barScanProgress.Visibility = Visibility.Collapsed;
				btnCancelScan.Visibility = Visibility.Collapsed;
				List<WordSpellIssueContext> activeIssues = _controller.ActiveIssues;
				for (int num = 0; num < activeIssues.Count; num++)
				{
					_allViewModels.Add(new SpellIssueViewModel
					{
						Context = activeIssues[num]
					});
				}
				UpdateCountersAndList();
			}
		}
		catch (Exception ex)
		{
			FailScan(ex.Message);
		}
	}

	private void FailScan(string message)
	{
		CancelScan(showStatus: false);
		_allViewModels.Clear();
		lblFooterStatus.Text = $"Không thể quét đoạn {_nextParagraph}: {message}. Kết quả chưa đầy đủ; hãy thử quét lại.";
	}

	private void CancelScan(bool showStatus)
	{
		_scanGeneration++;
		_scanTimer.Stop();
		if (_scanCancellation != null)
		{
			_scanCancellation.Cancel();
		}
		if (barScanProgress != null)
		{
			barScanProgress.Visibility = Visibility.Collapsed;
		}
		if (btnCancelScan != null)
		{
			btnCancelScan.Visibility = Visibility.Collapsed;
		}
		if (showStatus && lblFooterStatus != null)
		{
			lblFooterStatus.Text = "Đã hủy quét. Tài liệu chưa bị thay đổi.";
		}
	}

	private void BtnCancelScan_Click(object sender, RoutedEventArgs e)
	{
		CancelScan(showStatus: true);
	}

	private void UpdateCountersAndList()
	{
		int num = 0;
		int num2 = 0;
		int num3 = 0;
		int num4 = 0;
		int num5 = 0;
		for (int i = 0; i < _allViewModels.Count; i++)
		{
			SpellCheckIssue issue = _allViewModels[i].Issue;
			if (_allViewModels[i].IsForeignTermNote)
			{
				num5++;
				continue;
			}
			switch (issue.Type)
			{
			case SpellCheckIssueType.Spelling:
				num++;
				break;
			case SpellCheckIssueType.Capitalization:
				num2++;
				break;
			case SpellCheckIssueType.Abbreviation:
				num3++;
				break;
			case SpellCheckIssueType.Variant:
				num4++;
				break;
			case SpellCheckIssueType.Repetition:
				num4++;
				break;
			}
		}
		int count = _allViewModels.Count;
		radFilterAll.Content = $"Tất cả ({count})";
		if (_isSpacingMode)
		{
			radFilterSpelling.Content = $"Khoảng trắng: {num}";
			radFilterCap.Content = $"Dấu câu / Ngoặc: {num2}";
			radFilterAbbr.Content = $"Dấu nháy (\"): {num3}";
			radFilterVariant.Content = $"Dấu gạch (–): {num4}";
			ApplyFilter();
			lblFooterStatus.Text = $"Đã quét xong: Phát hiện {count} lỗi ({num} khoảng trắng, {num2} dấu câu/ngoặc, {num3} dấu nháy \", {num4} dấu gạch –).";
		}
		else
		{
			radFilterSpelling.Content = $"Chính tả: {num}";
			radFilterCap.Content = $"Viết hoa: {num2}";
			radFilterAbbr.Content = $"Viết tắt: {num3}";
			radFilterVariant.Content = $"Từ ngữ: {num4}";
			ApplyFilter();
			lblFooterStatus.Text = ((num5 > 0) ? $"Đã quét xong: {count} mục ({num} chính tả, {num2} viết hoa, {num3} viết tắt, {num4} từ ngữ; {num5} thuật ngữ nước ngoài chỉ để tham khảo)." : $"Đã quét xong: Phát hiện {count} vấn đề ({num} lỗi chính tả, {num2} viết hoa, {num3} viết tắt, {num4} từ ngữ).");
		}
	}

	private void ApplyFilter()
	{
		List<SpellIssueViewModel> list = new List<SpellIssueViewModel>();
		for (int i = 0; i < _allViewModels.Count; i++)
		{
			SpellIssueViewModel spellIssueViewModel = _allViewModels[i];
			bool? isChecked = radFilterSpelling.IsChecked;
			if (isChecked == true && isChecked.HasValue && (spellIssueViewModel.Issue.Type != SpellCheckIssueType.Spelling || spellIssueViewModel.IsForeignTermNote))
			{
				continue;
			}
			bool? isChecked2 = radFilterCap.IsChecked;
			if (isChecked2 == true && isChecked2.HasValue && spellIssueViewModel.Issue.Type != SpellCheckIssueType.Capitalization)
			{
				continue;
			}
			bool? isChecked3 = radFilterAbbr.IsChecked;
			if (isChecked3 != true || !isChecked3.HasValue || spellIssueViewModel.Issue.Type == SpellCheckIssueType.Abbreviation)
			{
				bool? isChecked4 = radFilterVariant.IsChecked;
				if (isChecked4 != true || !isChecked4.HasValue || spellIssueViewModel.Issue.Type == SpellCheckIssueType.Variant || spellIssueViewModel.Issue.Type == SpellCheckIssueType.Repetition)
				{
					list.Add(spellIssueViewModel);
				}
			}
		}
		lstIssues.ItemsSource = list;
		if (list.Count == 0)
		{
			pnlNoIssues.Visibility = Visibility.Visible;
			pnlDetail.Visibility = Visibility.Collapsed;
		}
		else
		{
			pnlNoIssues.Visibility = Visibility.Collapsed;
			pnlDetail.Visibility = Visibility.Visible;
			lstIssues.SelectedIndex = 0;
		}
	}

	private void LstIssues_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (!(lstIssues.SelectedItem is SpellIssueViewModel spellIssueViewModel))
		{
			pnlDetail.Visibility = Visibility.Collapsed;
			return;
		}
		pnlDetail.Visibility = Visibility.Visible;
		SpellCheckIssue issue = spellIssueViewModel.Issue;
		if (_controller != null && _doc != null)
		{
			_controller.HighlightIssue((Document)_doc, spellIssueViewModel.Context);
		}
		txtOriginalWord.Text = spellIssueViewModel.OriginalDisplay;
		lblIssueTypeBadge.Text = spellIssueViewModel.TagLabel;
		lblIssueTypeBadge.Foreground = spellIssueViewModel.TagColor;
		txtIssueDescription.Text = issue.Description;
		TextBlock textBlock = txtIssueSource;
		string text;
		if (string.IsNullOrEmpty(issue.Source))
		{
			text = "Nguồn: Từ điển chính thức OfficeFormat";
		}
		else
		{
			text = ((issue.Source.IndexOf(".dic", StringComparison.OrdinalIgnoreCase) >= 0) ? "Căn cứ: Bộ dữ liệu kiểm tra chính tả OfficeFormat" : string.Format("Căn cứ: {0}", (issue.Source == "VFORMAT_HEURISTIC") ? "Quy tắc kiểm tra chính tả của OfficeFormat" : issue.Source));
		}
		textBlock.Text = text;
		if (issue.HasUsageGuide)
		{
			pnlUsageGuide.Visibility = Visibility.Visible;
			if (!string.IsNullOrEmpty(issue.WhenToUse))
			{
				boxWhenToUse.Visibility = Visibility.Visible;
				txtWhenToUse.Text = issue.WhenToUse;
			}
			else
			{
				boxWhenToUse.Visibility = Visibility.Collapsed;
			}
			if (!string.IsNullOrEmpty(issue.WhenNotToUse))
			{
				boxWhenNotToUse.Visibility = Visibility.Visible;
				txtWhenNotToUse.Text = issue.WhenNotToUse;
			}
			else
			{
				boxWhenNotToUse.Visibility = Visibility.Collapsed;
			}
			if (!string.IsNullOrEmpty(issue.UsageGuide))
			{
				boxUsageGuide.Visibility = Visibility.Visible;
				txtUsageGuide.Text = issue.UsageGuide;
			}
			else
			{
				boxUsageGuide.Visibility = Visibility.Collapsed;
			}
			if (issue.Examples != null && issue.Examples.Count > 0)
			{
				boxExamples.Visibility = Visibility.Visible;
				txtExamples.Text = string.Join(Environment.NewLine, issue.Examples.ToArray());
			}
			else
			{
				boxExamples.Visibility = Visibility.Collapsed;
			}
		}
		else
		{
			pnlUsageGuide.Visibility = Visibility.Collapsed;
		}
		List<SuggestionViewModel> list = new List<SuggestionViewModel>();
		if (issue.Suggestions != null)
		{
			for (int i = 0; i < issue.Suggestions.Count; i++)
			{
				SpellSuggestion spellSuggestion = issue.Suggestions[i];
				list.Add(new SuggestionViewModel
				{
					Text = spellSuggestion.Text,
					IsSpacingSuggestion = spellIssueViewModel.IsSpacingIssue,
					Score = spellSuggestion.Score,
					Reason = spellSuggestion.Reason
				});
			}
		}
		lstSuggestions.ItemsSource = list;
		if (list.Count > 0)
		{
			lstSuggestions.SelectedIndex = 0;
			txtCustomReplacement.Text = list[0].Text;
		}
		else
		{
			txtCustomReplacement.Text = string.Empty;
		}
	}

	private void LstSuggestions_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lstSuggestions.SelectedItem is SuggestionViewModel suggestionViewModel && txtCustomReplacement != null)
		{
			txtCustomReplacement.Text = suggestionViewModel.Text;
		}
	}

	private void BtnReplace_Click(object sender, RoutedEventArgs e)
	{
		if (!(lstIssues.SelectedItem is SpellIssueViewModel spellIssueViewModel))
		{
			return;
		}
		if (_isSpacingMode)
		{
			string replacement = txtCustomReplacement.Text ?? string.Empty;
			if (lstSuggestions.SelectedItem is SuggestionViewModel selectedSug &&
				string.Equals(replacement, selectedSug.Text, StringComparison.Ordinal))
			{
				replacement = selectedSug.Text;
			}
			int selectedIndex = lstIssues.SelectedIndex;
			int paraIdx = spellIssueViewModel.Context.ParagraphIndex;
			string origDisplay = spellIssueViewModel.OriginalDisplay;
			string repDisplay = SpacingCheckScanner.FormatVisibleSpaces(replacement);
			if (_controller.ReplaceIssue((Document)_doc, spellIssueViewModel.Context, replacement))
			{
				_allViewModels.Remove(spellIssueViewModel);
				UpdateCountersAndList();
				if (lstIssues.Items.Count > 0)
				{
					int selectedIndex2 = Math.Min(selectedIndex, lstIssues.Items.Count - 1);
					lstIssues.SelectedIndex = selectedIndex2;
				}
				lblFooterStatus.Text = $"Đã sửa tại Đoạn {paraIdx}: '{origDisplay}' ➔ '{repDisplay}'. Còn lại {_allViewModels.Count} lỗi.";
			}
			else
			{
				lblFooterStatus.Text = "Văn bản đã thay đổi, đang tự động quét lại...";
				RunScan();
			}
			return;
		}
		string text = txtCustomReplacement.Text.Trim();
		if (string.IsNullOrEmpty(text) && lstSuggestions.SelectedItem is SuggestionViewModel)
		{
			text = ((SuggestionViewModel)lstSuggestions.SelectedItem).Text;
		}
		if (!string.IsNullOrEmpty(text) && ConfirmReplacementPreview(new List<WordSpellIssueContext> { spellIssueViewModel.Context }, text, matchCase: false))
		{
			if (_controller.ReplaceIssue((Document)_doc, spellIssueViewModel.Context, text))
			{
				RunScan();
			}
			else
			{
				lblFooterStatus.Text = "Không thay thế: văn bản đã đổi hoặc Word từ chối sửa. Hãy quét lại.";
			}
		}
	}

	private void BtnReplaceAll_Click(object sender, RoutedEventArgs e)
	{
		List<SpellReplacement> list = _controller.BuildReplacementPlan(out var skipped);
		if (list.Count == 0)
		{
			lblFooterStatus.Text = "Không có lỗi có gợi ý sửa riêng phù hợp. Các lỗi không có gợi ý hoặc trùng vị trí được giữ nguyên.";
			return;
		}
		if (_isSpacingMode)
		{
			int fixedCount = _controller.ReplaceAllIssues((Document)_doc, list);
			if (fixedCount > 0)
			{
				RunScan();
				lblFooterStatus.Text = $"Đã tự động sửa {fixedCount} lỗi khoảng trắng, dấu nháy & dấu gạch!";
			}
			else
			{
				lblFooterStatus.Text = "Văn bản đã thay đổi, đang quét lại...";
				RunScan();
			}
			return;
		}
		List<ReplacementPreviewItem> list2 = new List<ReplacementPreviewItem>();
		try
		{
			foreach (SpellReplacement item in list)
			{
				if (!_controller.TryGetReplacementPreview((Document)_doc, item.Context, item.Replacement, matchCase: false, out var before, out var after))
				{
					lblFooterStatus.Text = "Văn bản đã đổi từ lần quét trước. Hãy quét lại trước khi sửa.";
					return;
				}
				list2.Add(new ReplacementPreviewItem
				{
					Paragraph = item.Context.ParagraphIndex,
					Original = item.Original,
					Replacement = item.Replacement,
					Before = before,
					After = after
				});
			}
		}
		catch (Exception ex)
		{
			lblFooterStatus.Text = "Không thể xem trước: " + ex.Message;
			return;
		}
		ReplacementPreviewWindow replacementPreviewWindow = new ReplacementPreviewWindow(list2, skipped, all: true);
		replacementPreviewWindow.Owner = this;
		if (replacementPreviewWindow.ShowDialog() != true)
		{
			return;
		}
		int num = _controller.ReplaceAllIssues((Document)_doc, list);
		if (num > 0)
		{
			RunScan();
			if (num != list.Count)
			{
				MessageBox.Show(this, $"Đã thay {num}/{list.Count} vị trí. Hãy dùng Ctrl+Z để hoàn tác lượt sửa nếu cần.", "Thay thế chưa đầy đủ", MessageBoxButton.OK, MessageBoxImage.Exclamation);
			}
		}
		else
		{
			lblFooterStatus.Text = "Không thay thế: một vị trí đã đổi hoặc Word từ chối sửa. Hãy quét lại.";
		}
	}

	private bool ConfirmReplacementPreview(List<WordSpellIssueContext> contexts, string replacement, bool matchCase)
	{
		if (contexts == null || contexts.Count == 0)
		{
			return false;
		}
		List<ReplacementPreviewItem> list = new List<ReplacementPreviewItem>();
		try
		{
			for (int i = 0; i < contexts.Count; i++)
			{
				if (!_controller.TryGetReplacementPreview((Document)_doc, contexts[i], replacement, matchCase, out var before, out var after))
				{
					lblFooterStatus.Text = "Văn bản đã đổi từ lần quét trước. Hãy quét lại trước khi sửa.";
					return false;
				}
				list.Add(new ReplacementPreviewItem
				{
					Paragraph = contexts[i].ParagraphIndex,
					Original = contexts[i].Issue.Original,
					Replacement = replacement,
					Before = before,
					After = after
				});
			}
		}
		catch (Exception ex)
		{
			lblFooterStatus.Text = "Không thể xem trước: " + ex.Message;
			return false;
		}
		ReplacementPreviewWindow replacementPreviewWindow = new ReplacementPreviewWindow(list, 0, all: false);
		replacementPreviewWindow.Owner = this;
		return replacementPreviewWindow.ShowDialog() == true;
	}

	private void BtnIgnore_Click(object sender, RoutedEventArgs e)
	{
		if (lstIssues.SelectedItem is SpellIssueViewModel spellIssueViewModel)
		{
			int selectedIndex = lstIssues.SelectedIndex;
			_controller.IgnoreIssue(spellIssueViewModel.Context);
			_allViewModels.Remove(spellIssueViewModel);
			UpdateCountersAndList();
			if (lstIssues.Items.Count > 0)
			{
				int selectedIndex2 = Math.Min(selectedIndex, lstIssues.Items.Count - 1);
				lstIssues.SelectedIndex = selectedIndex2;
			}
		}
	}

	private void BtnIgnoreAll_Click(object sender, RoutedEventArgs e)
	{
		int num = _controller.IgnoreAllIssues();
		_allViewModels.Clear();
		UpdateCountersAndList();
		lblFooterStatus.Text = $"Đã bỏ qua tất cả {num} lỗi đang báo. Nội dung tài liệu được giữ nguyên.";
	}

	private void BtnAddDictionary_Click(object sender, RoutedEventArgs e)
	{
		if (lstIssues.SelectedItem is SpellIssueViewModel spellIssueViewModel)
		{
			string token = spellIssueViewModel.Issue.Original;
			try
			{
				_controller.AddToDictionary(token);
			}
			catch (Exception ex)
			{
				MessageBox.Show(this, "Chưa lưu được từ vào từ điển. Bạn có thể thử lại.\n" + ex.Message, "Không thể lưu từ điển", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				return;
			}
			_allViewModels.RemoveAll((SpellIssueViewModel x) => x.Issue.Original.Equals(token, StringComparison.OrdinalIgnoreCase));
			UpdateCountersAndList();
			MessageBox.Show($"Đã thêm từ '{token}' vào từ điển cơ quan.\nTừ này sẽ không bị báo lỗi trong các lần kiểm tra sau.", "Thêm Vào Từ Điển Thành Công", MessageBoxButton.OK, MessageBoxImage.Asterisk);
		}
	}

	private void BtnRescan_Click(object sender, RoutedEventArgs e)
	{
		RunScan();
	}

	private void CboProfile_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (IsLoaded && !_selectingInitialProfile)
		{
			RunScan();
		}
	}

	private void Filter_Checked(object sender, RoutedEventArgs e)
	{
		if (IsLoaded)
		{
			ApplyFilter();
		}
	}

	private void BtnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}
}
