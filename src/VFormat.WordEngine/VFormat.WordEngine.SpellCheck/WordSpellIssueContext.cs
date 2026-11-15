using VFormat.Core.SpellCheck.Models;

namespace VFormat.WordEngine.SpellCheck;

public class WordSpellIssueContext
{
	public SpellCheckIssue Issue { get; set; }

	public int ParagraphIndex { get; set; }

	public int ParaRangeStart { get; set; }

	public int ParaRangeEnd { get; set; }

	public WordSpellIssueContext(SpellCheckIssue issue, int paraIndex, int paraStart, int paraEnd)
	{
		Issue = issue;
		ParagraphIndex = paraIndex;
		ParaRangeStart = paraStart;
		ParaRangeEnd = paraEnd;
	}
}
