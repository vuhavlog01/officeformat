using VFormat.Core.SpellCheck.Models;

namespace VFormat.Core.SpellCheck.Processing;

public class TextToken
{
	public string Text { get; set; }

	public int Start { get; set; }

	public int End { get; set; }

	public TokenType Type { get; set; }

	public TextToken(string text, int start, int end, TokenType type)
	{
		Text = text;
		Start = start;
		End = end;
		Type = type;
	}

	public override string ToString()
	{
		return $"[{Type}] '{Text}' ({Start}-{End})";
	}
}
