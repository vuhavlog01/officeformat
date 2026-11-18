using VFormat.Core.SpellCheck;

namespace VFormat.UI.SpellCheck;

public class SuggestionViewModel
{
	public string Text { get; set; }

	public bool IsSpacingSuggestion { get; set; }

	public string DisplayText
	{
		get
		{
			if (!IsSpacingSuggestion)
			{
				return Text;
			}
			if (string.IsNullOrEmpty(Text))
			{
				return "(Xóa khoảng trắng thừa)";
			}
			if (Text == " ")
			{
				return "␣ (1 dấu cách chuẩn)";
			}
			return SpacingCheckScanner.FormatVisibleSpaces(Text);
		}
	}

	public double Score { get; set; }

	public string Reason { get; set; }

	public string ScoreDisplay => $"{(int)(Score * 100.0)}% tin cậy";
}
