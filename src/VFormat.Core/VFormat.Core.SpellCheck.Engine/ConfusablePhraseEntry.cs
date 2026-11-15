using System.Collections.Generic;

namespace VFormat.Core.SpellCheck.Engine;

public class ConfusablePhraseEntry
{
	public string replacement { get; set; }

	public string description { get; set; }

	public string category { get; set; }

	public string severity { get; set; }

	public string source { get; set; }

	public string when_to_use { get; set; }

	public string when_not_to_use { get; set; }

	public string usage_guide { get; set; }

	public List<string> examples { get; set; }

	public List<string> not_after { get; set; }

	public List<string> not_before { get; set; }

	public List<string> only_before { get; set; }
}
