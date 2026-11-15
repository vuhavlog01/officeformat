using System.Collections.Generic;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Personalization;

namespace VFormat.Core.Interfaces;

public interface IProfileRepository
{
	List<ProfileModel> GetAllProfiles();

	ProfileModel GetProfile(string profileId);

	ProfileModel GetDefaultProfile(RulePackType rulePack);

	void SaveProfile(ProfileModel profile);

	UserPersonalization LoadPersonalization();

	void SavePersonalization(UserPersonalization personalization);

	void ResetToDefaults();
}
