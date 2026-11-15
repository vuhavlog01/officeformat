using VFormat.Core.Enums;

namespace VFormat.Core.Interfaces;

public interface IFormatRule
{
	string RuleId { get; }

	RulePackType RulePack { get; }

	string Name { get; }

	string Description { get; }

	string SourceReference { get; }

	RuleSeverity Severity { get; }

	bool CanAutoFix { get; }
}
