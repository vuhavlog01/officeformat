using VFormat.Core.Enums;

namespace VFormat.Core.Rules;

public static class RulePackTraits
{
	public const string PartyHeaderText = "ĐẢNG CỘNG SẢN VIỆT NAM";

	public const string YouthUnionHeaderText = "ĐOÀN TNCS HỒ CHÍ MINH";

	public const int StandardIndexAdmin = 0;

	public const int StandardIndexParty = 1;

	public const int StandardIndexMilitary = 2;

	public const int StandardIndexCustom = 3;

	public const int StandardIndexYouth = 4;

	public static bool UsesOrganizationHeader(RulePackType rulePack)
	{
		if (rulePack != RulePackType.HD05_VPTW)
		{
			return rulePack == RulePackType.HD29_TWDTN;
		}
		return true;
	}

	public static bool IsParty(RulePackType rulePack)
	{
		return rulePack == RulePackType.HD05_VPTW;
	}

	public static bool IsYouthUnion(RulePackType rulePack)
	{
		return rulePack == RulePackType.HD29_TWDTN;
	}

	public static bool IsMilitary(RulePackType rulePack)
	{
		return rulePack == RulePackType.TT68_2023_BQP;
	}

	public static bool UsesNationalHeader(RulePackType rulePack)
	{
		return !UsesOrganizationHeader(rulePack);
	}

	public static string GetOrganizationHeaderText(RulePackType rulePack)
	{
		if (!IsYouthUnion(rulePack))
		{
			return "ĐẢNG CỘNG SẢN VIỆT NAM";
		}
		return "ĐOÀN TNCS HỒ CHÍ MINH";
	}

	public static string GetAgencySeparator(RulePackType rulePack)
	{
		if (IsYouthUnion(rulePack))
		{
			return "***";
		}
		if (IsParty(rulePack))
		{
			return "*";
		}
		return string.Empty;
	}

	public static bool IsAsteriskSeparator(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		string text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
			.Replace("\v", "")
			.Replace(" ", "")
			.Replace("\t", "")
			.Trim();
		if (text2.Length == 0 || text2.Length > 5)
		{
			return false;
		}
		string text3 = text2;
		foreach (char c in text3)
		{
			if (c != '*' && c != '∗' && c != '⁎' && c != '＊')
			{
				return false;
			}
		}
		return true;
	}

	public static string GetNumberPrefix(RulePackType rulePack)
	{
		if (!IsParty(rulePack))
		{
			return "Số: ";
		}
		return "Số ";
	}

	public static bool UsesDashNumberSymbol(RulePackType rulePack)
	{
		return UsesOrganizationHeader(rulePack);
	}

	public static bool UsesTextSeparatorBelowExcerpt(RulePackType rulePack)
	{
		return UsesOrganizationHeader(rulePack);
	}

	public static string GetExcerptSeparatorText(RulePackType rulePack)
	{
		if (!IsYouthUnion(rulePack))
		{
			return "-----";
		}
		return "---------------";
	}

	public static bool ExcerptSeparatorBold(RulePackType rulePack)
	{
		return !IsYouthUnion(rulePack);
	}

	public static bool UsesSlashAuthority(RulePackType rulePack)
	{
		return IsParty(rulePack);
	}

	public static string GetShortName(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => "Văn bản Đảng (HD 05)", 
			RulePackType.TT68_2023_BQP => "Bộ Quốc phòng (TT 68)", 
			RulePackType.HD29_TWDTN => "Văn bản Đoàn (HD 29)", 
			_ => "Hành chính Nhà nước (Nghị định 30)", 
		};
	}

	public static string GetSourceTitle(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => "Hướng dẫn 05-HD/VPTW (Văn bản Đảng)", 
			RulePackType.TT68_2023_BQP => "Thông tư 68/2023/TT-BQP (Quân đội)", 
			RulePackType.HD29_TWDTN => "Hướng dẫn 29-HD/TWĐTN-VP (Văn bản Đoàn TNCS Hồ Chí Minh)", 
			_ => "Nghị định 30/2020/NĐ-CP (Hành chính Nhà nước)", 
		};
	}

	public static int ToStandardIndex(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => 1, 
			RulePackType.TT68_2023_BQP => 2, 
			RulePackType.HD29_TWDTN => 4, 
			_ => 0, 
		};
	}

	public static RulePackType FromStandardIndex(int standard)
	{
		return standard switch
		{
			1 => RulePackType.HD05_VPTW, 
			2 => RulePackType.TT68_2023_BQP, 
			4 => RulePackType.HD29_TWDTN, 
			_ => RulePackType.ND30_2020, 
		};
	}
}
