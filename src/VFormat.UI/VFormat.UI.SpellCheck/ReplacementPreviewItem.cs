namespace VFormat.UI.SpellCheck;

public sealed class ReplacementPreviewItem
{
	public int Paragraph { get; set; }

	public string Original { get; set; }

	public string Replacement { get; set; }

	public string Before { get; set; }

	public string After { get; set; }
}
