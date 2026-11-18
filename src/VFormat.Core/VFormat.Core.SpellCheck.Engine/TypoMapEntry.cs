using System.Collections.Generic;

namespace VFormat.Core.SpellCheck.Engine;

public class TypoMapEntry
{
	public List<string> suggestions { get; set; }

	public double confidence { get; set; }

	public string category { get; set; }

	public bool autofix { get; set; }
}
