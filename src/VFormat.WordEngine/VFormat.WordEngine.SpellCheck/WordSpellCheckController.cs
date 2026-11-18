using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Interop.Word;
using VFormat.Core.SpellCheck;
using VFormat.Core.SpellCheck.Models;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.SpellCheck;

public class WordSpellCheckController
{
	private readonly Application _app;

	private SpellCheckService _spellService;

	private readonly List<WordSpellIssueContext> _activeIssues;

	public List<WordSpellIssueContext> ActiveIssues => _activeIssues;

	public SpellCheckService SpellService => _spellService;

	public WordSpellCheckController(Application app, SpellCheckService spellService = null)
	{
		_app = app;
		_spellService = spellService;
		_activeIssues = new List<WordSpellIssueContext>();
	}

	public List<WordSpellIssueContext> ScanDocument(Document doc, SpellCheckProfile profile = SpellCheckProfile.Administrative)
	{
		int num = BeginScan(doc);
		for (int i = 1; i <= num; i++)
		{
			ScanParagraph(doc, i, profile);
		}
		return _activeIssues;
	}

	public List<WordSpellIssueContext> ScanDocumentSpacing(Document doc)
	{
		int num = BeginScan(doc);
		for (int i = 1; i <= num; i++)
		{
			AddIssues(AnalyzeSpacingSnapshot(CaptureParagraph(doc, i), CancellationToken.None));
		}
		return _activeIssues;
	}

	public int BeginScan(Document doc)
	{
		_activeIssues.Clear();
		if (doc == null)
		{
			return 0;
		}
		if (_spellService != null)
		{
			_spellService.ResetDocumentSession();
		}
		return doc.Paragraphs.Count;
	}

	public void ScanParagraph(Document doc, int paragraphIndex, SpellCheckProfile profile)
	{
		AddIssues(AnalyzeSnapshot(CaptureParagraph(doc, paragraphIndex), profile));
	}

	public SpellParagraphSnapshot CaptureParagraph(Document doc, int paragraphIndex)
	{
		Paragraph comObj = null;
		Range comObj2 = null;
		try
		{
			comObj = doc.Paragraphs[paragraphIndex];
			comObj2 = comObj.Range;
			string text = comObj2.Text;
			string text2 = (text ?? string.Empty).TrimEnd('\r', '\a', '\n');
			return new SpellParagraphSnapshot(text2, paragraphIndex, comObj2.Start, comObj2.End);
		}
		finally
		{
			ComReleaser.Release(ref comObj2);
			ComReleaser.Release(ref comObj);
		}
	}

	public List<WordSpellIssueContext> AnalyzeSnapshot(SpellParagraphSnapshot snapshot, SpellCheckProfile profile)
	{
		return AnalyzeSnapshot(snapshot, profile, CancellationToken.None);
	}

	public List<WordSpellIssueContext> AnalyzeSnapshot(SpellParagraphSnapshot snapshot, SpellCheckProfile profile, CancellationToken cancellationToken)
	{
		cancellationToken.ThrowIfCancellationRequested();
		List<WordSpellIssueContext> list = new List<WordSpellIssueContext>();
		if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.Text))
		{
			return list;
		}
		if (_spellService == null)
		{
			_spellService = new SpellCheckService();
		}
		List<SpellCheckIssue> list2 = _spellService.CheckParagraph(snapshot.Text, snapshot.ParagraphIndex, profile, cancellationToken);
		for (int i = 0; i < list2.Count; i++)
		{
			list.Add(new WordSpellIssueContext(list2[i], snapshot.ParagraphIndex, snapshot.RangeStart, snapshot.RangeEnd));
		}
		return list;
	}

	public List<WordSpellIssueContext> AnalyzeSpacingSnapshot(SpellParagraphSnapshot snapshot, CancellationToken cancellationToken)
	{
		cancellationToken.ThrowIfCancellationRequested();
		List<WordSpellIssueContext> list = new List<WordSpellIssueContext>();
		if (snapshot == null || string.IsNullOrEmpty(snapshot.Text))
		{
			return list;
		}
		List<SpellCheckIssue> list2 = SpacingCheckScanner.CheckParagraph(snapshot.Text, snapshot.ParagraphIndex, cancellationToken);
		for (int i = 0; i < list2.Count; i++)
		{
			list.Add(new WordSpellIssueContext(list2[i], snapshot.ParagraphIndex, snapshot.RangeStart, snapshot.RangeEnd));
		}
		return list;
	}

	public void AddIssues(List<WordSpellIssueContext> issues)
	{
		if (issues != null)
		{
			_activeIssues.AddRange(issues);
		}
	}

	public void HighlightIssue(Document doc, WordSpellIssueContext issueCtx)
	{
		if (doc == null || issueCtx == null)
		{
			return;
		}
		try
		{
			int num = issueCtx.ParaRangeStart + issueCtx.Issue.Start;
			int num2 = issueCtx.ParaRangeStart + issueCtx.Issue.End;
			if (num >= 0 && num2 > num)
			{
				object Start = num;
				object End = num2;
				Range range = doc.Range(ref Start, ref End);
				range.Select();
			}
		}
		catch
		{
		}
	}

	public bool TryGetReplacementPreview(Document doc, WordSpellIssueContext issueCtx, string replacementText, bool matchCase, out string before, out string after)
	{
		before = (after = string.Empty);
		if (doc == null || issueCtx == null || replacementText == null)
		{
			return false;
		}
		bool isSpacingIssue = issueCtx.Issue?.RuleId != null && issueCtx.Issue.RuleId.StartsWith("SPACE_", StringComparison.Ordinal);
		if (!isSpacingIssue && string.IsNullOrEmpty(replacementText))
		{
			return false;
		}
		int num = issueCtx.ParaRangeStart + issueCtx.Issue.Start;
		int num2 = issueCtx.ParaRangeStart + issueCtx.Issue.End;
		if (num < 0 || num2 <= num || num2 > issueCtx.ParaRangeEnd)
		{
			return false;
		}
		Range comObj = null;
		Range comObj2 = null;
		try
		{
			object Start = num;
			object End = num2;
			comObj = doc.Range(ref Start, ref End);
			if (!string.Equals(comObj.Text, issueCtx.Issue.Original, StringComparison.Ordinal))
			{
				return false;
			}
			object Start2 = issueCtx.ParaRangeStart;
			object End2 = issueCtx.ParaRangeEnd;
			comObj2 = doc.Range(ref Start2, ref End2);
			string text = comObj2.Text ?? string.Empty;
			int start = issueCtx.Issue.Start;
			int end = issueCtx.Issue.End;
			if (start < 0 || end > text.Length)
			{
				return false;
			}
			int num3 = Math.Max(0, start - 35);
			int num4 = Math.Min(text.Length, end + 35);
			string text2 = ((num3 > 0) ? "…" : string.Empty);
			string text3 = ((num4 < text.Length) ? "…" : string.Empty);
			string text4 = ((!isSpacingIssue && matchCase) ? MatchCase(issueCtx.Issue.Original, replacementText) : replacementText);
			before = CleanPreview(text2 + text.Substring(num3, num4 - num3) + text3);
			after = CleanPreview(text2 + text.Substring(num3, start - num3) + text4 + text.Substring(end, num4 - end) + text3);
			return true;
		}
		finally
		{
			ComReleaser.Release(ref comObj2);
			ComReleaser.Release(ref comObj);
		}
	}

	private static string CleanPreview(string text)
	{
		return text.Replace('\r', ' ').Replace('\n', ' ').Replace('\a', ' ')
			.Trim();
	}

	public bool ReplaceIssue(Document doc, WordSpellIssueContext issueCtx, string replacementText)
	{
		if (doc == null || issueCtx == null || replacementText == null)
		{
			return false;
		}
		bool isSpacingIssue = issueCtx.Issue?.RuleId != null && issueCtx.Issue.RuleId.StartsWith("SPACE_", StringComparison.Ordinal);
		if (!isSpacingIssue && string.IsNullOrEmpty(replacementText))
		{
			return false;
		}
		if (!TryGetReplacementPreview(doc, issueCtx, replacementText, matchCase: false, out var _, out var _))
		{
			return false;
		}
		using (new SafeUndoScope(_app, "OfficeFormat Sửa Lỗi"))
		{
			try
			{
				int num = issueCtx.ParaRangeStart + issueCtx.Issue.Start;
				int num2 = issueCtx.ParaRangeStart + issueCtx.Issue.End;
				if (num >= 0 && num2 > num)
				{
					object Start = num;
					object End = num2;
					Range range = doc.Range(ref Start, ref End);
					int num3 = num2 - num;
					int length = replacementText.Length;
					int num4 = length - num3;
					range.Text = replacementText;
					int paragraphIndex = issueCtx.ParagraphIndex;
					for (int i = 0; i < _activeIssues.Count; i++)
					{
						WordSpellIssueContext wordSpellIssueContext = _activeIssues[i];
						if (wordSpellIssueContext == issueCtx)
						{
							continue;
						}
						if (wordSpellIssueContext.ParagraphIndex == paragraphIndex)
						{
							wordSpellIssueContext.ParaRangeEnd += num4;
							if (wordSpellIssueContext.Issue.Start >= issueCtx.Issue.End)
							{
								wordSpellIssueContext.Issue.Start += num4;
								wordSpellIssueContext.Issue.End += num4;
							}
						}
						else if (wordSpellIssueContext.ParaRangeStart > issueCtx.ParaRangeStart)
						{
							wordSpellIssueContext.ParaRangeStart += num4;
							wordSpellIssueContext.ParaRangeEnd += num4;
						}
					}
					_activeIssues.Remove(issueCtx);
					return true;
				}
			}
			catch
			{
				return false;
			}
		}
		return false;
	}

	public int ReplaceAll(Document doc, string originalToken, string replacementText)
	{
		if (doc == null || string.IsNullOrEmpty(originalToken) || string.IsNullOrEmpty(replacementText))
		{
			return 0;
		}
		WordSpellIssueContext selected = _activeIssues.Find((WordSpellIssueContext x) => x.Issue.Original.Equals(originalToken, StringComparison.OrdinalIgnoreCase));
		return ReplaceAll(doc, selected, replacementText);
	}

	public List<WordSpellIssueContext> GetReplacementGroup(WordSpellIssueContext selected)
	{
		if (selected == null || !_activeIssues.Contains(selected))
		{
			return new List<WordSpellIssueContext>();
		}
		return _activeIssues.FindAll((WordSpellIssueContext x) => x.Issue.Original.Equals(selected.Issue.Original, StringComparison.OrdinalIgnoreCase) && x.Issue.Type == selected.Issue.Type && string.Equals(x.Issue.RuleId, selected.Issue.RuleId, StringComparison.Ordinal) && string.Equals(x.Issue.FirstSuggestion, selected.Issue.FirstSuggestion, StringComparison.OrdinalIgnoreCase));
	}

	public int ReplaceAll(Document doc, WordSpellIssueContext selected, string replacementText)
	{
		if (doc == null || selected == null || string.IsNullOrEmpty(replacementText))
		{
			return 0;
		}
		string original = selected.Issue.Original;
		List<WordSpellIssueContext> replacementGroup = GetReplacementGroup(selected);
		if (replacementGroup.Count == 0)
		{
			return 0;
		}
		for (int i = 0; i < replacementGroup.Count; i++)
		{
			if (!TryGetReplacementPreview(doc, replacementGroup[i], replacementText, matchCase: true, out var _, out var _))
			{
				return 0;
			}
		}
		replacementGroup.Sort((WordSpellIssueContext a, WordSpellIssueContext b) =>
		{
			int value = a.ParaRangeStart + a.Issue.Start;
			return (b.ParaRangeStart + b.Issue.Start).CompareTo(value);
		});
		int num = 0;
		using (new SafeUndoScope(_app, "OfficeFormat Sửa Tất Cả: " + original))
		{
			for (int num2 = 0; num2 < replacementGroup.Count; num2++)
			{
				WordSpellIssueContext wordSpellIssueContext = replacementGroup[num2];
				try
				{
					int num3 = wordSpellIssueContext.ParaRangeStart + wordSpellIssueContext.Issue.Start;
					int num4 = wordSpellIssueContext.ParaRangeStart + wordSpellIssueContext.Issue.End;
					if (num3 < 0 || num4 <= num3)
					{
						continue;
					}
					object Start = num3;
					object End = num4;
					Range range = doc.Range(ref Start, ref End);
					int num5 = num4 - num3;
					string text = (range.Text = MatchCase(wordSpellIssueContext.Issue.Original, replacementText));
					int num6 = text.Length - num5;
					for (int num7 = 0; num7 < _activeIssues.Count; num7++)
					{
						WordSpellIssueContext wordSpellIssueContext2 = _activeIssues[num7];
						if (wordSpellIssueContext2 == wordSpellIssueContext)
						{
							continue;
						}
						int num8 = wordSpellIssueContext2.ParaRangeStart + wordSpellIssueContext2.Issue.Start;
						if (num8 > num3)
						{
							if (wordSpellIssueContext2.ParagraphIndex == wordSpellIssueContext.ParagraphIndex)
							{
								wordSpellIssueContext2.Issue.Start += num6;
								wordSpellIssueContext2.Issue.End += num6;
							}
							else
							{
								wordSpellIssueContext2.ParaRangeStart += num6;
								wordSpellIssueContext2.ParaRangeEnd += num6;
							}
						}
					}
					_activeIssues.Remove(wordSpellIssueContext);
					num++;
				}
				catch
				{
				}
			}
			return num;
		}
	}

	private static string MatchCase(string original, string target)
	{
		if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(target))
		{
			return target;
		}
		bool flag = true;
		foreach (char c in original)
		{
			if (char.IsLetter(c) && !char.IsUpper(c) && c != 'Đ')
			{
				flag = false;
				break;
			}
		}
		if (flag && original.Length > 1)
		{
			return target.ToUpperInvariant();
		}
		if (char.IsUpper(original[0]))
		{
			return char.ToUpperInvariant(target[0]) + ((target.Length > 1) ? target.Substring(1) : string.Empty);
		}
		return target;
	}

	public void IgnoreIssue(WordSpellIssueContext issueCtx)
	{
		if (issueCtx != null)
		{
			issueCtx.Issue.IsIgnored = true;
			_activeIssues.Remove(issueCtx);
		}
	}

	public void IgnoreAll(string token)
	{
		if (!string.IsNullOrEmpty(token))
		{
			if (_spellService != null)
			{
				_spellService.IgnoreWord(token);
			}
			_activeIssues.RemoveAll((WordSpellIssueContext x) => x.Issue.Original.Equals(token, StringComparison.OrdinalIgnoreCase));
		}
	}

	public int IgnoreAllIssues()
	{
		int count = _activeIssues.Count;
		foreach (WordSpellIssueContext activeIssue in _activeIssues)
		{
			activeIssue.Issue.IsIgnored = true;
			if (_spellService != null)
			{
				_spellService.IgnoreWord(activeIssue.Issue.Original);
			}
		}
		_activeIssues.Clear();
		return count;
	}

	public List<SpellReplacement> BuildReplacementPlan(out int skipped)
	{
		List<SpellReplacement> list = new List<SpellReplacement>();
		foreach (WordSpellIssueContext activeIssue in _activeIssues)
		{
			string text = activeIssue.Issue.FirstSuggestion;
			bool isSpacingIssue = activeIssue.Issue.RuleId != null && activeIssue.Issue.RuleId.StartsWith("SPACE_", StringComparison.Ordinal);
			if (isSpacingIssue ? (text != null) : (!string.IsNullOrWhiteSpace(text)))
			{
				if (!isSpacingIssue && activeIssue.Issue.Type != SpellCheckIssueType.Capitalization)
				{
					text = MatchCase(activeIssue.Issue.Original, text);
				}
				if (!string.Equals(activeIssue.Issue.Original, text, StringComparison.Ordinal))
				{
					list.Add(new SpellReplacement(activeIssue, text));
				}
			}
		}
		list.Sort((SpellReplacement a, SpellReplacement b) => a.Start.CompareTo(b.Start));
		HashSet<SpellReplacement> conflicts = new HashSet<SpellReplacement>();
		for (int num = 0; num < list.Count; num++)
		{
			for (int num2 = num + 1; num2 < list.Count && list[num2].Start < list[num].End; num2++)
			{
				conflicts.Add(list[num]);
				conflicts.Add(list[num2]);
			}
		}
		list.RemoveAll((SpellReplacement x) => conflicts.Contains(x));
		skipped = _activeIssues.Count - list.Count;
		return list;
	}

	public int ReplaceAllIssues(Document doc, List<SpellReplacement> plan)
	{
		if (doc == null || plan == null || plan.Count == 0)
		{
			return 0;
		}
		try
		{
			foreach (SpellReplacement item in plan)
			{
				if (!_activeIssues.Contains(item.Context))
				{
					return 0;
				}
				Range comObj = null;
				try
				{
					object Start = item.Start;
					object End = item.End;
					comObj = doc.Range(ref Start, ref End);
					if (!string.Equals(comObj.Text, item.Original, StringComparison.Ordinal))
					{
						return 0;
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
			return 0;
		}
		List<SpellReplacement> list = new List<SpellReplacement>(plan);
		list.Sort((SpellReplacement a, SpellReplacement b) => b.Start.CompareTo(a.Start));
		int num = 0;
		using (new SafeUndoScope(_app, "OfficeFormat Sửa Tất Cả Lỗi"))
		{
			foreach (SpellReplacement item2 in list)
			{
				Range comObj2 = null;
				try
				{
					object Start2 = item2.Start;
					object End2 = item2.End;
					comObj2 = doc.Range(ref Start2, ref End2);
					comObj2.Text = item2.Replacement;
					_activeIssues.Remove(item2.Context);
					num++;
				}
				catch
				{
					break;
				}
				finally
				{
					ComReleaser.Release(ref comObj2);
				}
			}
		}
		return num;
	}

	public void AddToDictionary(string token)
	{
		if (!string.IsNullOrEmpty(token))
		{
			if (_spellService != null)
			{
				_spellService.AddCustomWord(token);
			}
			_activeIssues.RemoveAll((WordSpellIssueContext x) => x.Issue.Original.Equals(token, StringComparison.OrdinalIgnoreCase));
		}
	}
}
