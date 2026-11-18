using System.Collections.Generic;

namespace VFormat.Core.Templates;

public class TemplateCategoryItem
{
	public string Id { get; set; }

	public string Name { get; set; }

	public string Icon { get; set; }

	public string Description { get; set; }

	public int TemplateCount { get; set; }

	public List<string> Groups { get; set; }

	public string BadgeText => TemplateCount + " biểu mẫu";

	public TemplateCategoryItem()
	{
		Groups = new List<string>();
	}
}
