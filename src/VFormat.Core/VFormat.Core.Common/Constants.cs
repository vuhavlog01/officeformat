using System;
using System.IO;

namespace VFormat.Core.Common;

public static class Constants
{
	public const string CurrentVersion = "1.0.0";

	public const string AppDisplayName = "OfficeFormat 1.0.0";

	public const string AppDisplayNameUpper = "OFFICEFORMAT 1.0.0";

	public const string VersionBadge = "v1.0.0";

	public const float A4WidthMm = 210f;

	public const float A4HeightMm = 297f;

	public const float DefaultMarginTopMm = 20f;

	public const float DefaultMarginBottomMm = 20f;

	public const float DefaultMarginLeftMm = 30f;

	public const float DefaultMarginRightMm = 15f;

	public const float DefaultDuplexMarginLeftMm = 15f;

	public const float DefaultDuplexMarginRightMm = 30f;

	public const float DefaultHeaderDistanceMm = 10f;

	public const float DefaultFooterDistanceMm = 10f;

	public const string StandardFontName = "Times New Roman";

	public const string StandardEncoding = "Unicode TCVN 6909:2001";

	public const string StandardFontColorHex = "#000000";

	public const float DefaultBodyFontSizePt = 14f;

	public const float OptionalBodyFontSizePt = 13f;

	public const float DefaultLineSpacingMultiple = 1.15f;

	public const float DefaultLineSpacingExactlyPt = 20f;

	public const float DefaultSpacingBeforePt = 0f;

	public const float DefaultSpacingAfterPt = 6f;

	public const float PartySpacingAfterPt = 0f;

	public const float DefaultFirstLineIndentMm = 10f;

	public const float DefaultPageNumberSizePt = 14f;

	public const float MilitaryPageNumberSizePt = 14f;

	public const float AdminPageNumberSizePt = 14f;

	public const string AddinProgId = "VFormat.Addin";

	public const string AddinRegistryPath = "Software\\Microsoft\\Office\\Word\\Addins\\VFormat.Addin";

	public const string AppFolder = "VFormat";

	public const string ConfigFileName = "vformat_presets.json";

	public const string LicenseFileName = "vformat_license.lic";

	public const string UpdaterExeName = "VFormat.Updater.exe";

	public static string AppDataDirectory
	{
		get
		{
			string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
			string text = Path.Combine(folderPath, "VFormat");
			if (!Directory.Exists(text))
			{
				try
				{
					Directory.CreateDirectory(text);
				}
				catch
				{
				}
			}
			return text;
		}
	}

	public static string ProgramDirectory
	{
		get
		{
			string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
			return Path.Combine(folderPath, "Programs", "VFormat");
		}
	}

	public static string UpdateTempDirectory
	{
		get
		{
			string tempPath = Path.GetTempPath();
			string text = Path.Combine(tempPath, "VFormatUpdate");
			if (!Directory.Exists(text))
			{
				try
				{
					Directory.CreateDirectory(text);
				}
				catch
				{
				}
			}
			return text;
		}
	}

	public static string ConfigFilePath => Path.Combine(AppDataDirectory, "vformat_presets.json");
}
