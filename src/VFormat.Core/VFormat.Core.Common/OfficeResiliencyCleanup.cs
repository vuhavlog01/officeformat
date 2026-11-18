using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace VFormat.Core.Common;

public static class OfficeResiliencyCleanup
{
	private static readonly Regex OwnedIdentifier = new Regex("(?<![\\p{L}\\p{N}_.-])VFormat\\.(?:Addin\\.dll|Addin\\.TaskPaneControl|Addin)(?![\\p{L}\\p{N}_.-])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

	private static readonly Regex OwnedFriendlyName = new Regex("(?:^|\\x00)(?:OfficeFormat|OfficeFormat) \\d+\\.\\d+\\.\\d+(?:\\.\\d+)?(?:\\x00|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

	private static bool ContainsOwnedIdentifier(string text)
	{
		if (!OwnedIdentifier.IsMatch(text ?? ""))
		{
			return OwnedFriendlyName.IsMatch(text ?? "");
		}
		return true;
	}

	public static bool IdentifiesVFormat(string name, object data)
	{
		if (ContainsOwnedIdentifier(name))
		{
			return true;
		}
		if (data is byte[] array)
		{
			if (!ContainsOwnedIdentifier(Encoding.Unicode.GetString(array)) && !ContainsOwnedIdentifier(Encoding.UTF8.GetString(array)))
			{
				if (array.Length > 1)
				{
					return ContainsOwnedIdentifier(Encoding.Unicode.GetString(array, 1, array.Length - 1));
				}
				return false;
			}
			return true;
		}
		if (data is string text)
		{
			return ContainsOwnedIdentifier(text);
		}
		if (data is string[] array2)
		{
			string[] array3 = array2;
			foreach (string text2 in array3)
			{
				if (ContainsOwnedIdentifier(text2))
				{
					return true;
				}
			}
		}
		return false;
	}

	public static void RemoveOwnValues(RegistryKey root, string path)
	{
		using RegistryKey registryKey = root.OpenSubKey(path, writable: true);
		if (registryKey == null)
		{
			return;
		}
		string[] valueNames = registryKey.GetValueNames();
		foreach (string name in valueNames)
		{
			if (IdentifiesVFormat(name, registryKey.GetValue(name)))
			{
				registryKey.DeleteValue(name, throwOnMissingValue: false);
			}
		}
	}
}
