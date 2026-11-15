namespace VFormat.Core.SpellCheck.Processing;

public class ProtectedSpan
{
	public int Start { get; set; }

	public int End { get; set; }

	public string Type { get; set; }

	public string Text { get; set; }

	public ProtectedSpan(int start, int end, string type, string text)
	{
		Start = start;
		End = end;
		Type = type;
		Text = text;
	}

	public override string ToString()
	{
		return $"[{Type}] '{Text}' ({Start}-{End})";
	}
}
