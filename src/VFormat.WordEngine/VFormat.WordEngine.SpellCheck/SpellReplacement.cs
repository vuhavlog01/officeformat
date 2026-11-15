namespace VFormat.WordEngine.SpellCheck;

public sealed class SpellReplacement
{
	public WordSpellIssueContext Context { get; private set; }

	public string Replacement { get; private set; }

	public string Original { get; private set; }

	public int Start { get; private set; }

	public int End { get; private set; }

	public SpellReplacement(WordSpellIssueContext context, string replacement)
	{
		Context = context;
		Replacement = replacement;
		Original = context.Issue.Original;
		Start = context.ParaRangeStart + context.Issue.Start;
		End = context.ParaRangeStart + context.Issue.End;
	}
}
