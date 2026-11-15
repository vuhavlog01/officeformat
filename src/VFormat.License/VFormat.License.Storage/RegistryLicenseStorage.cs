using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using VFormat.License.Hwid;

namespace VFormat.License.Storage;

public static class RegistryLicenseStorage
{
	public const string SecretSalt = "vformat_sec_k98f23490fdsmf902348mksdf9823m490f";

	private static string _testRegistryRoot;

	private static string VbaRegistryKey
	{
		get
		{
			if (_testRegistryRoot != null)
			{
				return _testRegistryRoot + "\\Legacy";
			}
			return "Software\\VB and VBA Program Settings\\VFormat\\License";
		}
	}

	private static string ModernRegistryKey
	{
		get
		{
			if (_testRegistryRoot != null)
			{
				return _testRegistryRoot + "\\Modern";
			}
			return "Software\\VFormat\\License";
		}
	}

	internal static void UseTestStorage(string root)
	{
		if (root == null || !root.StartsWith("Software\\VFormat\\Tests\\Release228\\", StringComparison.Ordinal))
		{
			throw new ArgumentException("Expected an isolated release regression key.");
		}
		_testRegistryRoot = root;
	}

	public static string GetLicenseKey()
	{
		string text = ReadString(ModernRegistryKey, "LicenseKey");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "LicenseKey");
	}

	public static string GetSavedLicenseKey()
	{
		return GetLicenseKey();
	}

	public static void SaveLicenseKey(string licenseKey)
	{
		WriteString(ModernRegistryKey, "LicenseKey", licenseKey);
		WriteString(VbaRegistryKey, "LicenseKey", licenseKey);
	}

	public static string GetOfflineCode()
	{
		string text = ReadString(ModernRegistryKey, "OfflineCode");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "OfflineCode");
	}

	public static bool IsOfflineLicense()
	{
		string text = ReadString(ModernRegistryKey, "IsOffline");
		if (string.IsNullOrEmpty(text))
		{
			text = ReadString(VbaRegistryKey, "IsOffline");
		}
		return text == "1";
	}

	public static bool IsBlocked()
	{
		string text = ReadString(ModernRegistryKey, "Blocked");
		if (string.IsNullOrEmpty(text))
		{
			text = ReadString(VbaRegistryKey, "Blocked");
		}
		return text == "1";
	}

	public static string GetCachedStatus()
	{
		string text = ReadString(ModernRegistryKey, "Status");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "Status");
	}

	public static string GetCachedPlanType()
	{
		string text = ReadString(ModernRegistryKey, "Type");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "Type");
	}

	public static string GetCachedCustomerName()
	{
		string text = ReadString(ModernRegistryKey, "CustomerName");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "CustomerName");
	}

	public static string GetCachedExpiresAt()
	{
		string text = ReadString(ModernRegistryKey, "ExpiresAt");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "ExpiresAt");
	}

	public static int GetCachedDaysLeft()
	{
		string text = ReadString(ModernRegistryKey, "DaysLeft");
		if (string.IsNullOrEmpty(text))
		{
			text = ReadString(VbaRegistryKey, "DaysLeft");
		}
		if (!int.TryParse(text, out var result))
		{
			return 0;
		}
		return result;
	}

	public static string GetCachedSeal()
	{
		string text = ReadString(ModernRegistryKey, "CacheSeal");
		if (!string.IsNullOrEmpty(text))
		{
			return text;
		}
		return ReadString(VbaRegistryKey, "CacheSeal");
	}

	public static DateTime? GetLastVerifiedTime()
	{
		string text = ReadString(ModernRegistryKey, "LastVerifiedAt");
		if (string.IsNullOrEmpty(text))
		{
			text = ReadString(VbaRegistryKey, "LastVerifiedAt");
		}
		if (DateTime.TryParse(text, out var result))
		{
			return result;
		}
		return null;
	}

	public static void SaveLastVerifiedTime(DateTime time)
	{
		string value = time.ToString("yyyy-MM-dd HH:mm:ss");
		WriteString(ModernRegistryKey, "LastVerifiedAt", value);
		WriteString(VbaRegistryKey, "LastVerifiedAt", value);
	}

	public static string GenerateCacheSeal(string licKey, string planType, string expDate)
	{
		try
		{
			string text = Regex.Replace((licKey ?? "").ToUpperInvariant(), "[^A-Z0-9]", "");
			string text2 = Regex.Replace(HwidGenerator.GetMachineHwid().ToUpperInvariant(), "[^A-Z0-9\\-]", "");
			string text3 = (planType ?? "yearly").Trim().ToLowerInvariant();
			string text4 = (expDate ?? "").Trim();
			if (text4.Contains(" "))
			{
				text4 = text4.Split(' ')[0];
			}
			if (text.Length < 8 || text2.Length < 4)
			{
				return string.Empty;
			}
			string s = $"VFLICSEAL|{text2}|{text}|{text3}|{text4}";
			using HMACSHA256 hMACSHA = new HMACSHA256(Encoding.UTF8.GetBytes("vformat_sec_k98f23490fdsmf902348mksdf9823m490f"));
			byte[] array = hMACSHA.ComputeHash(Encoding.UTF8.GetBytes(s));
			string text5 = BitConverter.ToString(array).Replace("-", "").ToUpperInvariant();
			return (text5.Length >= 24) ? text5.Substring(0, 24) : text5;
		}
		catch
		{
			return string.Empty;
		}
	}

	public static bool VerifyCacheSeal(string licKey, string planType, string expDate)
	{
		string text = ReadString(ModernRegistryKey, "CacheSeal");
		if (string.IsNullOrEmpty(text))
		{
			text = ReadString(VbaRegistryKey, "CacheSeal");
		}
		if (string.IsNullOrEmpty(text) || text.Length < 16)
		{
			return false;
		}
		string b = GenerateCacheSeal(licKey, planType, expDate);
		if (string.Equals(text, b, StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(expDate) && expDate.Contains(" "))
		{
			string b2 = GenerateCacheSeal(licKey, planType, expDate.Split(' ')[0]);
			if (string.Equals(text, b2, StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
		}
		if (!string.IsNullOrEmpty(planType) && !planType.Equals("yearly", StringComparison.OrdinalIgnoreCase))
		{
			string b3 = GenerateCacheSeal(licKey, "yearly", expDate);
			if (string.Equals(text, b3, StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
		}
		return false;
	}

	public static void SaveLicenseCache(string licKey, string status, string planType, string custName, string expDate, int daysLeft)
	{
		string text = (licKey ?? "").Trim().ToUpperInvariant();
		SaveLicenseKey(text);
		WriteString(ModernRegistryKey, "Status", status);
		WriteString(ModernRegistryKey, "Type", planType ?? "yearly");
		WriteString(ModernRegistryKey, "CustomerName", custName ?? "");
		WriteString(ModernRegistryKey, "ExpiresAt", expDate ?? "");
		WriteString(ModernRegistryKey, "DaysLeft", daysLeft.ToString());
		WriteString(ModernRegistryKey, "IsOffline", "0");
		WriteString(ModernRegistryKey, "Blocked", "0");
		WriteString(VbaRegistryKey, "Status", status);
		WriteString(VbaRegistryKey, "Type", planType ?? "yearly");
		WriteString(VbaRegistryKey, "CustomerName", custName ?? "");
		WriteString(VbaRegistryKey, "ExpiresAt", expDate ?? "");
		WriteString(VbaRegistryKey, "DaysLeft", daysLeft.ToString());
		WriteString(VbaRegistryKey, "IsOffline", "0");
		WriteString(VbaRegistryKey, "Blocked", "0");
		SaveLastVerifiedTime(DateTime.Now);
		if (status == "ACTIVE" || status == "TRIAL_ACTIVE")
		{
			string value = GenerateCacheSeal(text, planType, expDate);
			WriteString(ModernRegistryKey, "CacheSeal", value);
			WriteString(VbaRegistryKey, "CacheSeal", value);
		}
		else
		{
			WriteString(ModernRegistryKey, "CacheSeal", "");
			WriteString(VbaRegistryKey, "CacheSeal", "");
		}
	}

	public static void SaveOfflineLicense(string licenseKey, string hwid, string offlineCode, string expDate, string planType)
	{
		string text = (licenseKey ?? "").Trim().ToUpperInvariant();
		SaveLicenseKey(text);
		WriteString(ModernRegistryKey, "Status", "ACTIVE");
		WriteString(ModernRegistryKey, "Type", planType ?? "offline");
		WriteString(ModernRegistryKey, "CustomerName", "Bản quyền Nội bộ (Offline)");
		WriteString(ModernRegistryKey, "ExpiresAt", expDate ?? "");
		WriteString(ModernRegistryKey, "IsOffline", "1");
		WriteString(ModernRegistryKey, "OfflineCode", offlineCode);
		WriteString(ModernRegistryKey, "OfflineHWID", hwid);
		WriteString(ModernRegistryKey, "Blocked", "0");
		WriteString(VbaRegistryKey, "Status", "ACTIVE");
		WriteString(VbaRegistryKey, "Type", planType ?? "offline");
		WriteString(VbaRegistryKey, "CustomerName", "Bản quyền Nội bộ (Offline)");
		WriteString(VbaRegistryKey, "ExpiresAt", expDate ?? "");
		WriteString(VbaRegistryKey, "IsOffline", "1");
		WriteString(VbaRegistryKey, "OfflineCode", offlineCode);
		WriteString(VbaRegistryKey, "OfflineHWID", hwid);
		WriteString(VbaRegistryKey, "Blocked", "0");
		string value = GenerateCacheSeal(text, planType, expDate);
		WriteString(ModernRegistryKey, "CacheSeal", value);
		WriteString(VbaRegistryKey, "CacheSeal", value);
		SaveLastVerifiedTime(DateTime.Now);
	}

	public static void PurgeLicenseCache(string newStatus = "NO_LICENSE", bool clearKey = false)
	{
		if (clearKey)
		{
			SaveLicenseKey("");
		}
		WriteString(ModernRegistryKey, "Status", newStatus);
		WriteString(ModernRegistryKey, "CacheSeal", "");
		WriteString(ModernRegistryKey, "CustomerName", "");
		WriteString(ModernRegistryKey, "ExpiresAt", "");
		WriteString(ModernRegistryKey, "DaysLeft", "0");
		WriteString(ModernRegistryKey, "IsOffline", "0");
		WriteString(ModernRegistryKey, "OfflineCode", "");
		WriteString(VbaRegistryKey, "Status", newStatus);
		WriteString(VbaRegistryKey, "CacheSeal", "");
		WriteString(VbaRegistryKey, "CustomerName", "");
		WriteString(VbaRegistryKey, "ExpiresAt", "");
		WriteString(VbaRegistryKey, "DaysLeft", "0");
		WriteString(VbaRegistryKey, "IsOffline", "0");
		WriteString(VbaRegistryKey, "OfflineCode", "");
		if (newStatus.Contains("LOCK") || newStatus.Contains("BLOCK") || newStatus.Contains("REVOK") || newStatus.Contains("TRANSF"))
		{
			WriteString(ModernRegistryKey, "Blocked", "1");
			WriteString(VbaRegistryKey, "Blocked", "1");
		}
		else
		{
			WriteString(ModernRegistryKey, "Blocked", "0");
			WriteString(VbaRegistryKey, "Blocked", "0");
		}
	}

	public static void SavePendingTransferRequest(string requestId)
	{
		WriteString(ModernRegistryKey, "TransferRequestId", requestId ?? "");
		WriteString(ModernRegistryKey, "TransferRequestAt", DateTime.UtcNow.ToString("o"));
	}

	public static string GetPendingTransferRequestId()
	{
		return ReadString(ModernRegistryKey, "TransferRequestId");
	}

	public static DateTime? GetPendingTransferRequestTimeUtc()
	{
		if (DateTime.TryParse(ReadString(ModernRegistryKey, "TransferRequestAt"), null, DateTimeStyles.RoundtripKind, out var result))
		{
			return result.ToUniversalTime();
		}
		return null;
	}

	public static void ClearPendingTransferRequest()
	{
		DeleteValue(ModernRegistryKey, "TransferRequestId");
		DeleteValue(ModernRegistryKey, "TransferRequestAt");
	}

	public static void MarkDeviceTransferred()
	{
		PurgeLicenseCache("DEVICE_TRANSFERRED");
		ClearPendingTransferRequest();
	}

	public static bool IsMarkedTransferred()
	{
		if (IsBlocked())
		{
			return string.Equals(GetCachedStatus(), "DEVICE_TRANSFERRED", StringComparison.OrdinalIgnoreCase);
		}
		return false;
	}

	public static void ClearLicense()
	{
		PurgeLicenseCache("NO_LICENSE", clearKey: true);
	}

	private static string ReadString(string subKey, string valueName)
	{
		try
		{
			using RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(subKey, writable: false);
			if (registryKey != null)
			{
				object value = registryKey.GetValue(valueName);
				if (value != null)
				{
					return value.ToString().Trim();
				}
			}
		}
		catch
		{
		}
		return string.Empty;
	}

	private static void WriteString(string subKey, string valueName, string value)
	{
		try
		{
			using RegistryKey registryKey = Registry.CurrentUser.CreateSubKey(subKey);
			registryKey?.SetValue(valueName, value ?? "", RegistryValueKind.String);
		}
		catch
		{
		}
	}

	private static void DeleteValue(string subKey, string valueName)
	{
		try
		{
			using RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(subKey, writable: true);
			registryKey?.DeleteValue(valueName, throwOnMissingValue: false);
		}
		catch
		{
		}
	}
}
