using System;
using System.IO;

namespace VFormat.Addin;

internal static class UpdaterLauncher
{
	public static string PrepareTemporaryCopy(string updaterPath)
	{
		if (string.IsNullOrEmpty(updaterPath) || !File.Exists(updaterPath))
		{
			throw new FileNotFoundException("Không tìm thấy VFormat.Updater.exe.", updaterPath);
		}
		string text = Path.Combine(Path.GetTempPath(), "VFormatUpdater");
		Directory.CreateDirectory(text);
		try
		{
			string[] directories = Directory.GetDirectories(text);
			foreach (string path in directories)
			{
				try
				{
					if (Directory.GetCreationTimeUtc(path) < DateTime.UtcNow.AddDays(-1.0))
					{
						Directory.Delete(path, recursive: true);
					}
				}
				catch
				{
				}
			}
		}
		catch
		{
		}
		string text2 = Path.Combine(text, Guid.NewGuid().ToString("N"));
		Directory.CreateDirectory(text2);
		string text3 = Path.Combine(text2, "VFormat.Updater.exe");
		File.Copy(updaterPath, text3, overwrite: true);
		string text4 = updaterPath + ".config";
		if (File.Exists(text4))
		{
			File.Copy(text4, text3 + ".config", overwrite: true);
		}
		return text3;
	}
}
