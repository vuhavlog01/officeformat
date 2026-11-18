using System;
using System.Collections.Generic;

namespace VFormat.Core.SpellCheck.Models;

public class SpellCheckIssue
{
	public string Id { get; set; }

	public int Start { get; set; }

	public int End { get; set; }

	public int ParagraphIndex { get; set; }

	public string Original { get; set; }

	public SpellCheckIssueType Type { get; set; }

	public SpellCheckSeverity Severity { get; set; }

	public List<SpellSuggestion> Suggestions { get; set; }

	public string RuleId { get; set; }

	public string Source { get; set; }

	public string Description { get; set; }

	public string ContextSnippet { get; set; }

	public bool IsIgnored { get; set; }

	public string WhenToUse { get; set; }

	public string WhenNotToUse { get; set; }

	public string UsageGuide { get; set; }

	public List<string> Examples { get; set; }

	public bool HasUsageGuide
	{
		get
		{
			if (string.IsNullOrEmpty(WhenToUse) && string.IsNullOrEmpty(WhenNotToUse))
			{
				return !string.IsNullOrEmpty(UsageGuide);
			}
			return true;
		}
	}

	public string FirstSuggestion
	{
		get
		{
			if (Suggestions != null && Suggestions.Count > 0)
			{
				return Suggestions[0].Text;
			}
			return string.Empty;
		}
	}

	public SpellCheckIssue()
	{
		Id = Guid.NewGuid().ToString("N");
		Suggestions = new List<SpellSuggestion>();
		Examples = new List<string>();
	}

	public override string ToString()
	{
		return $"[{Type}] '{Original}' at {Start}-{End}: {Description}";
	}
}
