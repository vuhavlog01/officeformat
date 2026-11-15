using VFormat.Core.Enums;

namespace VFormat.UI.TaskPane;

public class RulePackOption
{
	public RulePackType Type { get; set; }

	public string Title { get; set; }

	public string Subtitle { get; set; }

	public override string ToString()
	{
		return Title;
	}
}
