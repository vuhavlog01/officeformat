using System;

namespace VFormat.Core.Rules;

public class LegalFontSizeRange
{
	public int Min { get; set; }

	public int Max { get; set; }

	public int Default { get; set; }

	public string Description { get; set; }

	public LegalFontSizeRange()
	{
		Min = 13;
		Max = 14;
		Default = 14;
		Description = string.Empty;
	}

	public LegalFontSizeRange(int min, int max, int defaultValue, string description = "")
	{
		Min = min;
		Max = max;
		Default = defaultValue;
		Description = description ?? string.Empty;
	}

	public float Clamp(float value)
	{
		int num = (int)Math.Round(value);
		if (num < Min)
		{
			return Min;
		}
		if (num > Max)
		{
			return Max;
		}
		return num;
	}
}
