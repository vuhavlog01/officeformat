using System;
using VFormat.Core.Enums;

namespace VFormat.Core.Personalization;

public class SignerProfile
{
	public string Id { get; set; }

	public string Alias { get; set; }

	public string AuthorityTitle { get; set; }

	public string Position { get; set; }

	public string FullName { get; set; }

	public string MilitaryRank { get; set; }

	public string SignerType { get; set; }

	public RulePackType RulePack { get; set; }

	public string RecipientList { get; set; }

	public string Storage { get; set; }

	public string DisplaySignerName
	{
		get
		{
			if (!string.IsNullOrEmpty(MilitaryRank))
			{
				return $"{MilitaryRank.Trim()} {FullName.Trim()}";
			}
			return FullName;
		}
	}

	public SignerProfile()
	{
		Id = Guid.NewGuid().ToString("N");
		Alias = "Mẫu người ký mới";
		AuthorityTitle = string.Empty;
		Position = string.Empty;
		FullName = string.Empty;
		MilitaryRank = string.Empty;
		SignerType = string.Empty;
		RulePack = RulePackType.ND30_2020;
		RecipientList = string.Empty;
		Storage = string.Empty;
	}

	public SignerProfile(string id, string alias, string authorityTitle, string position, string fullName, string signerType, string militaryRank = "", RulePackType rulePack = RulePackType.ND30_2020, string recipientList = "", string storage = "")
	{
		Id = id;
		Alias = alias;
		AuthorityTitle = authorityTitle;
		Position = position;
		FullName = fullName;
		SignerType = signerType;
		MilitaryRank = militaryRank;
		RulePack = rulePack;
		RecipientList = recipientList ?? string.Empty;
		Storage = storage ?? string.Empty;
	}

	public SignerProfile Clone()
	{
		SignerProfile signerProfile = new SignerProfile();
		signerProfile.Id = Id;
		signerProfile.Alias = Alias;
		signerProfile.AuthorityTitle = AuthorityTitle;
		signerProfile.Position = Position;
		signerProfile.FullName = FullName;
		signerProfile.MilitaryRank = MilitaryRank;
		signerProfile.SignerType = SignerType;
		signerProfile.RulePack = RulePack;
		signerProfile.RecipientList = RecipientList;
		signerProfile.Storage = Storage;
		return signerProfile;
	}
}
