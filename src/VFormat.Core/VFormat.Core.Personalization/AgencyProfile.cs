using System;
using VFormat.Core.Enums;

namespace VFormat.Core.Personalization;

public class AgencyProfile
{
	public string Id { get; set; }

	public string Alias { get; set; }

	public string ParentAgency { get; set; }

	public string AgencyName { get; set; }

	public string Location { get; set; }

	public string CodePrefix { get; set; }

	public RulePackType RulePack { get; set; }

	public AgencyProfile()
	{
		Id = Guid.NewGuid().ToString("N");
		Alias = "Mẫu cơ quan mới";
		ParentAgency = string.Empty;
		AgencyName = string.Empty;
		Location = string.Empty;
		CodePrefix = string.Empty;
		RulePack = RulePackType.ND30_2020;
	}

	public AgencyProfile(string id, string alias, string parentAgency, string agencyName, string location, string codePrefix, RulePackType rulePack = RulePackType.ND30_2020)
	{
		Id = id;
		Alias = alias;
		ParentAgency = parentAgency;
		AgencyName = agencyName;
		Location = location;
		CodePrefix = codePrefix;
		RulePack = rulePack;
	}

	public AgencyProfile Clone()
	{
		AgencyProfile agencyProfile = new AgencyProfile();
		agencyProfile.Id = Id;
		agencyProfile.Alias = Alias;
		agencyProfile.ParentAgency = ParentAgency;
		agencyProfile.AgencyName = AgencyName;
		agencyProfile.Location = Location;
		agencyProfile.CodePrefix = CodePrefix;
		agencyProfile.RulePack = RulePack;
		return agencyProfile;
	}
}
