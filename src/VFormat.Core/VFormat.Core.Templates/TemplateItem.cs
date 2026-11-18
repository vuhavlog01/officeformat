using System;
using System.Collections.Generic;

namespace VFormat.Core.Templates;

[Serializable]
public class TemplateItem
{
	public string Id { get; set; }

	public string Name { get; set; }

	public string Title { get; set; }

	public string ShortTitle { get; set; }

	public string File { get; set; }

	public string Category { get; set; }

	public string Group { get; set; }

	public string TemplateNumber { get; set; }

	public string Source { get; set; }

	public string Description { get; set; }

	public List<string> Keywords { get; set; }

	public string Profile { get; set; }

	public string DocumentType { get; set; }

	public bool OfficialSample { get; set; }

	public string Type { get; set; }

	public bool Editable { get; set; }

	public string Version { get; set; }

	public string Ruleset { get; set; }

	public bool IsUserTemplate { get; set; }

	public string DisplayTitle
	{
		get
		{
			if (!string.IsNullOrEmpty(ShortTitle))
			{
				return ShortTitle;
			}
			if (!string.IsNullOrEmpty(Title))
			{
				return Title;
			}
			return Name;
		}
	}

	public string DisplayBadge
	{
		get
		{
			if (!string.IsNullOrEmpty(TemplateNumber))
			{
				if (TemplateNumber.StartsWith("Mẫu", StringComparison.OrdinalIgnoreCase))
				{
					return TemplateNumber;
				}
				if (char.IsDigit(TemplateNumber[0]))
				{
					return "Mẫu " + TemplateNumber;
				}
				return TemplateNumber;
			}
			return "Mẫu chuẩn";
		}
	}

	public string DisplaySource
	{
		get
		{
			if (string.IsNullOrEmpty(Source))
			{
				return "";
			}
			if (Source.Contains("152/2025"))
			{
				return "NĐ 152/2025/NĐ-CP";
			}
			if (Source.Contains("30/2020"))
			{
				return "NĐ 30/2020/NĐ-CP";
			}
			if (Source.Contains("05-HD/VPTW") || Source.Contains("05/VPTW"))
			{
				return "HD 05-HD/VPTW";
			}
			if (Source.Contains("68/2024"))
			{
				return "TT 68/2024/TT-BQP";
			}
			return Source;
		}
	}

	public TemplateItem()
	{
		Keywords = new List<string>();
	}
}
