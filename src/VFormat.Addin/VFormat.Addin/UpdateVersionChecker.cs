using System;
using System.Collections.Generic;

namespace VFormat.Addin;

internal static class UpdateVersionChecker
{
	public static bool TryReadLatestVersion(Dictionary<string, object> versionData, out Version latestVersion, out string latestVersionText)
	{
		latestVersion = null;
		latestVersionText = string.Empty;
		if (versionData == null)
		{
			return false;
		}
		if (!versionData.TryGetValue("version", out var value) || value == null)
		{
			return false;
		}
		latestVersionText = value.ToString().Trim();
		if (latestVersionText.StartsWith("v", StringComparison.OrdinalIgnoreCase))
		{
			latestVersionText = latestVersionText.Substring(1);
		}
		int num = latestVersionText.IndexOfAny(new char[2] { '-', '+' });
		string input = ((num >= 0) ? latestVersionText.Substring(0, num) : latestVersionText);
		if (!Version.TryParse(input, out latestVersion))
		{
			latestVersionText = string.Empty;
			return false;
		}
		return true;
	}
}
