namespace VFormat.WordEngine.SpellCheck;

public sealed class SpellParagraphSnapshot
{
	public string Text { get; private set; }

	public int ParagraphIndex { get; private set; }

	public int RangeStart { get; private set; }

	public int RangeEnd { get; private set; }

	public SpellParagraphSnapshot(string text, int paragraphIndex, int rangeStart, int rangeEnd)
	{
		Text = text;
		ParagraphIndex = paragraphIndex;
		RangeStart = rangeStart;
		RangeEnd = rangeEnd;
	}
}
