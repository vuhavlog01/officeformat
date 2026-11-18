using System.Collections.Generic;
using VFormat.Core.Enums;

namespace VFormat.Core.Models;

public class ProfileModel
{
	public string ProfileId { get; set; }

	public RulePackType RulePack { get; set; }

	public string Name { get; set; }

	public string Source { get; set; }

	public string Description { get; set; }

	public TypographyPreset Typography { get; set; }

	public LayoutPreset PageSetup { get; set; }

	public PageNumberPreset PageNumber { get; set; }

	public TablePreset TablePreset { get; set; }

	public bool HeaderDateInTopCell { get; set; }

	public Dictionary<SemanticRole, ComponentStyle> ComponentStyles { get; set; }

	public ProfileModel()
	{
		ProfileId = "PROFILE_DEFAULT";
		RulePack = RulePackType.ND30_2020;
		Name = "Mặc định";
		Source = string.Empty;
		Description = string.Empty;
		Typography = new TypographyPreset();
		PageSetup = new LayoutPreset();
		PageNumber = new PageNumberPreset();
		TablePreset = new TablePreset();
		ComponentStyles = new Dictionary<SemanticRole, ComponentStyle>();
	}

	public ProfileModel(string profileId, RulePackType rulePack, string name, string source, string description)
	{
		ProfileId = profileId;
		RulePack = rulePack;
		Name = name;
		Source = source;
		Description = description;
		Typography = new TypographyPreset();
		PageSetup = new LayoutPreset();
		PageNumber = new PageNumberPreset();
		TablePreset = new TablePreset();
		ComponentStyles = new Dictionary<SemanticRole, ComponentStyle>();
	}

	public ComponentStyle GetComponentStyle(SemanticRole role)
	{
		if (ComponentStyles.TryGetValue(role, out var value))
		{
			return value;
		}
		return new ComponentStyle(role, Typography.BodySizePt, bold: false, italic: false, Typography.Alignment);
	}

	public void SetComponentStyle(ComponentStyle style)
	{
		if (style != null)
		{
			ComponentStyles[style.Role] = style;
		}
	}

	public ProfileModel Clone()
	{
		ProfileModel profileModel = new ProfileModel();
		profileModel.ProfileId = ProfileId;
		profileModel.RulePack = RulePack;
		profileModel.Name = Name;
		profileModel.Source = Source;
		profileModel.Description = Description;
		profileModel.Typography = Typography.Clone();
		profileModel.PageSetup = PageSetup.Clone();
		profileModel.PageNumber = PageNumber.Clone();
		profileModel.TablePreset = TablePreset.Clone();
		profileModel.HeaderDateInTopCell = HeaderDateInTopCell;
		ProfileModel profileModel2 = profileModel;
		foreach (KeyValuePair<SemanticRole, ComponentStyle> componentStyle in ComponentStyles)
		{
			profileModel2.ComponentStyles[componentStyle.Key] = componentStyle.Value.Clone();
		}
		return profileModel2;
	}
}
