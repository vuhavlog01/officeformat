using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace VFormat.Core.Common;

public static class ComRegistrationInfo
{
	public const string MachineInstallKey = "Software\\TAFINEX\\VFormat";

	private const int WtsUserName = 5;

	private const int WtsDomainName = 7;

	public static string ToRegAsmCodeBase(string path)
	{
		if (string.IsNullOrEmpty(path))
		{
			throw new ArgumentNullException("path");
		}
		string fullPath = Path.GetFullPath(path);
		if (fullPath.StartsWith("\\\\", StringComparison.Ordinal))
		{
			throw new ArgumentException("Không đăng ký OfficeFormat từ đường dẫn mạng: " + fullPath);
		}
		if (fullPath.IndexOf('#') >= 0 || fullPath.IndexOf('%') >= 0)
		{
			throw new ArgumentException("Thư mục cài đặt chứa ký tự # hoặc %, không dùng được cho đăng ký COM: " + fullPath);
		}
		return "file:///" + fullPath.Replace('\\', '/');
	}

	public static bool TryGetPathFromCodeBase(string codeBase, out string path)
	{
		path = null;
		if (string.IsNullOrEmpty(codeBase) || !codeBase.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
		{
			return false;
		}
		path = codeBase.Substring(8).Replace('/', '\\');
		return true;
	}

	public static bool IsAdministratorToken()
	{
		try
		{
			return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
		}
		catch
		{
			return false;
		}
	}

	public static bool IsUacDisabled()
	{
		try
		{
			using RegistryKey registryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System");
			object obj = registryKey?.GetValue("EnableLUA");
			return obj is int && (int)obj == 0;
		}
		catch
		{
			return false;
		}
	}

	public static bool ComIgnoresPerUserClasses()
	{
		if (IsAdministratorToken())
		{
			return IsUacDisabled();
		}
		return false;
	}

	public static bool IsUnderProgramFiles(string directory)
	{
		if (string.IsNullOrEmpty(directory))
		{
			return false;
		}
		string text;
		try
		{
			text = Path.GetFullPath(directory).TrimEnd('\\', '/') + "\\";
		}
		catch
		{
			return false;
		}
		string[] array = new string[3]
		{
			Environment.GetEnvironmentVariable("ProgramW6432"),
			Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
			Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
		};
		foreach (string text2 in array)
		{
			if (!string.IsNullOrEmpty(text2))
			{
				string text3 = Path.GetFullPath(text2).TrimEnd('\\', '/') + "\\";
				if (text.StartsWith(text3, StringComparison.OrdinalIgnoreCase) && text.Length > text3.Length)
				{
					return true;
				}
			}
		}
		return false;
	}

	public static RegistryKey OpenMachineRoot()
	{
		return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);
	}

	public static RegistryKey OpenUserRoot()
	{
		return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);
	}

	[DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
	private static extern bool WTSQuerySessionInformation(IntPtr server, int sessionId, int infoClass, out IntPtr buffer, out int bytesReturned);

	[DllImport("wtsapi32.dll")]
	private static extern void WTSFreeMemory(IntPtr memory);

	private static string QuerySessionString(int sessionId, int infoClass)
	{
		if (!WTSQuerySessionInformation(IntPtr.Zero, sessionId, infoClass, out var buffer, out var _) || buffer == IntPtr.Zero)
		{
			return null;
		}
		try
		{
			return Marshal.PtrToStringUni(buffer);
		}
		finally
		{
			WTSFreeMemory(buffer);
		}
	}

	public static string GetSessionUserName()
	{
		try
		{
			int sessionId;
			using (Process process = Process.GetCurrentProcess())
			{
				sessionId = process.SessionId;
			}
			string text = QuerySessionString(sessionId, 5);
			if (string.IsNullOrEmpty(text))
			{
				return null;
			}
			string text2 = QuerySessionString(sessionId, 7);
			return string.IsNullOrEmpty(text2) ? text : (text2 + "\\" + text);
		}
		catch
		{
			return null;
		}
	}

	public static bool IsRunningAsDifferentUserThanSession()
	{
		try
		{
			string sessionUserName = GetSessionUserName();
			string name = WindowsIdentity.GetCurrent().Name;
			if (string.IsNullOrEmpty(sessionUserName) || string.IsNullOrEmpty(name))
			{
				return false;
			}
			if (string.Equals(sessionUserName, name, StringComparison.OrdinalIgnoreCase))
			{
				return false;
			}
			try
			{
				SecurityIdentifier securityIdentifier = (SecurityIdentifier)new NTAccount(sessionUserName).Translate(typeof(SecurityIdentifier));
				return !securityIdentifier.Equals(WindowsIdentity.GetCurrent().User);
			}
			catch
			{
				return true;
			}
		}
		catch
		{
			return false;
		}
	}

	public static string GetMachineInstallDirectory()
	{
		try
		{
			using RegistryKey registryKey = OpenMachineRoot();
			using RegistryKey registryKey2 = registryKey.OpenSubKey("Software\\TAFINEX\\VFormat");
			string text = ((registryKey2 == null) ? null : (registryKey2.GetValue("InstallDir") as string));
			if (!string.IsNullOrEmpty(text) && File.Exists(Path.Combine(text, "VFormat.Addin.dll")))
			{
				return text;
			}
		}
		catch
		{
		}
		return null;
	}
}
