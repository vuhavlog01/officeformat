namespace VFormat.Core.SpellCheck.Models;

public class SpellSuggestion
{
	public string Text { get; set; }

	public double Score { get; set; }

	public string Reason { get; set; }

	public SpellSuggestion()
	{
	}

	public SpellSuggestion(string text, double score, string reason)
	{
		Text = text;
		Score = score;
		Reason = reason;
	}

	public override string ToString()
	{
		return $"{Text} ({Score:0.##})";
	}
}
