using System.Collections.Generic;
using VFormat.Core.Enums;

namespace VFormat.Core.Models;

public sealed class RulePackDetectionResult
{
	public RulePackType RulePack { get; set; }

	public int Confidence { get; set; }

	public bool IsConfident { get; set; }

	public IList<string> Evidence { get; private set; }

	public IList<string> Warnings { get; private set; }

	public string Summary
	{
		get
		{
			string arg = ((Evidence.Count > 0) ? string.Join("; ", Evidence) : "chưa có dấu hiệu thể thức đủ mạnh");
			return $"Độ tin cậy {Confidence}% — {arg}";
		}
	}

	public RulePackDetectionResult()
	{
		RulePack = RulePackType.ND30_2020;
		Evidence = new List<string>();
		Warnings = new List<string>();
	}
}
