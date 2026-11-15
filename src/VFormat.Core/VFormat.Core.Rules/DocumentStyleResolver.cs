using System;
using System.Text;
using System.Text.RegularExpressions;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.Core.Rules;

public static class DocumentStyleResolver
{
	public static ComponentStyle Resolve(ProfileModel profile, SemanticRole role, string documentType)
	{
		if (profile == null)
		{
			throw new ArgumentNullException("profile");
		}
		ComponentStyle componentStyle = profile.GetComponentStyle(role);
		if (role != SemanticRole.LegalGround)
		{
			return componentStyle;
		}
		ComponentStyle componentStyle2 = componentStyle.Clone();
		componentStyle2.Bold = false;
		componentStyle2.Italic = profile.RulePack == RulePackType.ND30_2020 || (profile.RulePack == RulePackType.TT68_2023_BQP && UsesItalicGrounds(documentType));
		componentStyle2.SizePt = LegalFontSizeStandards.EnsureLegalInteger(profile.RulePack, role, componentStyle.SizePt);
		return componentStyle2;
	}

	public static bool UsesItalicGrounds(string documentType)
	{
		string input = Regex.Replace((documentType ?? "").Normalize(NormalizationForm.FormD), "\\p{IsCombiningDiacriticalMarks}+", "").Replace('đ', 'd').Replace('Đ', 'D');
		input = Regex.Replace(input, "\\s+", " ").Trim().ToUpperInvariant();
		return Regex.IsMatch(input, "^(NGHI QUYET|QUYET DINH|QUY CHE|QUY DINH)(\\s*\\(CA BIET\\))?\\s*:?$");
	}
}
