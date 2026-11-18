using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Personalization;

namespace VFormat.Core.Interfaces;

public interface IRuleEngine
{
	ProfileModel ActiveProfile { get; }

	UserPersonalization Personalization { get; }

	void SetActiveProfile(string profileId);

	void SetActiveProfile(RulePackType rulePack);

	RulePackType DetectRulePack(string sampleText);

	RulePackDetectionResult DetectRulePackDetailed(string sampleText);

	SemanticRole ClassifyParagraph(string text, int paragraphIndex, int totalParagraphs, bool isInTable);

	ComponentStyle GetComponentStyle(SemanticRole role);

	void UpdatePersonalization(UserPersonalization personalization);
}
