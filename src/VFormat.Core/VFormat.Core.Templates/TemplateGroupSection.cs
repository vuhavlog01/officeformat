using System.Collections.Generic;

namespace VFormat.Core.Templates;

public class TemplateGroupSection
{
	public string GroupName { get; set; }

	public List<TemplateItem> Templates { get; set; }

	public int Count
	{
		get
		{
			if (Templates == null)
			{
				return 0;
			}
			return Templates.Count;
		}
	}

	public string BadgeText => Count + " mẫu";

	public TemplateGroupSection()
	{
		Templates = new List<TemplateItem>();
	}
}
