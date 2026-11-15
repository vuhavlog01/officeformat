#define TRACE
using System;
using System.Diagnostics;
using System.Reflection;
using System.Security.Principal;
using Microsoft.Win32;
using VFormat.Core.Common;

namespace VFormat.Addin.Registration;

public static class AddinRegistrar
{
	public const string AddinProgId = "VFormat.Addin";

	public const string AddinClsid = "{7F28B490-A833-4F85-B2A8-3E2A398246E1}";

	public const string AddinClass = "VFormat.Addin.Addin";

	public const string ControlProgId = "VFormat.Addin.TaskPaneControl";

	public const string ControlClsid = "{6B68585D-7BF3-4C0E-9214-B9FD286E9CD5}";

	public const string ControlClass = "VFormat.Addin.TaskPaneControl";

	public const string FriendlyName = "OfficeFormat 1.0.0";

	public const string Description = "OfficeFormat 1.0.0 - Chuẩn hóa thể thức văn bản hành chính, Đảng bộ và Quân đội";

	private const string AssemblyVersionString = "2.0.1.0";

	private const string RuntimeVersion = "v4.0.30319";

	private static readonly string[] OfficeVersions = new string[5] { "16.0", "15.0", "14.0", "12.0", "11.0" };

	public static bool IsAdministrator()
	{
		try
		{
			WindowsIdentity current = WindowsIdentity.GetCurrent();
			WindowsPrincipal windowsPrincipal = new WindowsPrincipal(current);
			return windowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);
		}
		catch
		{
			return false;
		}
	}

	public static bool RegisterDualHive(string assemblyLocation = null)
	{
		try
		{
			if (string.IsNullOrEmpty(assemblyLocation))
			{
				assemblyLocation = Assembly.GetExecutingAssembly().Location;
			}
			string fullName = Assembly.GetExecutingAssembly().FullName;
			string codebaseUri = ComRegistrationInfo.ToRegAsmCodeBase(assemblyLocation);
			RegisterToRoot(Registry.CurrentUser, fullName, codebaseUri);
			return true;
		}
		catch (Exception ex)
		{
			Trace.WriteLine("RegisterDualHive Error: " + ex.Message);
			return false;
		}
	}

	public static bool RegisterPerUser(string assemblyLocation = null)
	{
		return RegisterDualHive(assemblyLocation);
	}

	private static void RegisterToRoot(RegistryKey root, string assemblyFullName, string codebaseUri)
	{
		RegisterComClass(root, "{7F28B490-A833-4F85-B2A8-3E2A398246E1}", "VFormat.Addin", "VFormat.Addin.Addin", assemblyFullName, codebaseUri);
		RegisterComClass(root, "{6B68585D-7BF3-4C0E-9214-B9FD286E9CD5}", "VFormat.Addin.TaskPaneControl", "VFormat.Addin.TaskPaneControl", assemblyFullName, codebaseUri);
		if (Environment.Is64BitOperatingSystem)
		{
			RegisterComClassWow64(root, "{7F28B490-A833-4F85-B2A8-3E2A398246E1}", "VFormat.Addin", "VFormat.Addin.Addin", assemblyFullName, codebaseUri);
			RegisterComClassWow64(root, "{6B68585D-7BF3-4C0E-9214-B9FD286E9CD5}", "VFormat.Addin.TaskPaneControl", "VFormat.Addin.TaskPaneControl", assemblyFullName, codebaseUri);
		}
		SetAddinRegistration(root, "Software\\Microsoft\\Office\\Word\\Addins\\VFormat.Addin");
		if (Environment.Is64BitOperatingSystem)
		{
			SetAddinRegistration(root, "Software\\Wow6432Node\\Microsoft\\Office\\Word\\Addins\\VFormat.Addin");
		}
	}

	private static void SetAddinRegistration(RegistryKey root, string subkey)
	{
		using RegistryKey registryKey = root.CreateSubKey(subkey);
		if (registryKey != null)
		{
			registryKey.SetValue("FriendlyName", "OfficeFormat 1.0.0", RegistryValueKind.String);
			registryKey.SetValue("Description", "OfficeFormat 1.0.0 - Chuẩn hóa thể thức văn bản hành chính, Đảng bộ và Quân đội", RegistryValueKind.String);
			registryKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
			registryKey.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
		}
	}

	public static bool UnregisterPerUser()
	{
		try
		{
			RegistryKey[] array = new RegistryKey[1] { Registry.CurrentUser };
			RegistryKey[] array2 = array;
			foreach (RegistryKey registryKey in array2)
			{
				try
				{
					registryKey.DeleteSubKeyTree("Software\\Microsoft\\Office\\Word\\Addins\\VFormat.Addin", throwOnMissingSubKey: false);
				}
				catch
				{
				}
				if (Environment.Is64BitOperatingSystem)
				{
					try
					{
						registryKey.DeleteSubKeyTree("Software\\Wow6432Node\\Microsoft\\Office\\Word\\Addins\\VFormat.Addin", throwOnMissingSubKey: false);
					}
					catch
					{
					}
				}
				UnregisterComClass(registryKey, "{7F28B490-A833-4F85-B2A8-3E2A398246E1}", "VFormat.Addin");
				UnregisterComClass(registryKey, "{6B68585D-7BF3-4C0E-9214-B9FD286E9CD5}", "VFormat.Addin.TaskPaneControl");
				if (Environment.Is64BitOperatingSystem)
				{
					UnregisterComClassWow64(registryKey, "{7F28B490-A833-4F85-B2A8-3E2A398246E1}", "VFormat.Addin");
					UnregisterComClassWow64(registryKey, "{6B68585D-7BF3-4C0E-9214-B9FD286E9CD5}", "VFormat.Addin.TaskPaneControl");
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			Trace.WriteLine("UnregisterPerUser Error: " + ex.Message);
			return false;
		}
	}

	private static void RegisterComClass(RegistryKey root, string clsid, string progId, string className, string assemblyFullName, string codebaseUri)
	{
		using (RegistryKey registryKey = root.CreateSubKey("Software\\Classes\\" + progId))
		{
			if (registryKey != null)
			{
				registryKey.SetValue("", className);
				using RegistryKey registryKey2 = registryKey.CreateSubKey("CLSID");
				registryKey2?.SetValue("", clsid);
			}
		}
		using RegistryKey registryKey3 = root.CreateSubKey("Software\\Classes\\CLSID\\" + clsid);
		if (registryKey3 == null)
		{
			return;
		}
		registryKey3.SetValue("", className);
		using (RegistryKey registryKey4 = registryKey3.CreateSubKey("InprocServer32"))
		{
			if (registryKey4 != null)
			{
				registryKey4.SetValue("", "mscoree.dll");
				registryKey4.SetValue("ThreadingModel", "Both");
				registryKey4.SetValue("Class", className);
				registryKey4.SetValue("Assembly", assemblyFullName);
				registryKey4.SetValue("RuntimeVersion", "v4.0.30319");
				registryKey4.SetValue("CodeBase", codebaseUri);
				using RegistryKey registryKey5 = registryKey4.CreateSubKey("2.0.1.0");
				if (registryKey5 != null)
				{
					registryKey5.SetValue("Class", className);
					registryKey5.SetValue("Assembly", assemblyFullName);
					registryKey5.SetValue("RuntimeVersion", "v4.0.30319");
					registryKey5.SetValue("CodeBase", codebaseUri);
				}
			}
		}
		using (RegistryKey registryKey6 = registryKey3.CreateSubKey("ProgId"))
		{
			registryKey6?.SetValue("", progId);
		}
		using (registryKey3.CreateSubKey("Implemented Categories\\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"))
		{
		}
	}

	private static void UnregisterComClass(RegistryKey root, string clsid, string progId)
	{
		try
		{
			root.DeleteSubKeyTree("Software\\Classes\\" + progId, throwOnMissingSubKey: false);
		}
		catch
		{
		}
		try
		{
			root.DeleteSubKeyTree("Software\\Classes\\CLSID\\" + clsid, throwOnMissingSubKey: false);
		}
		catch
		{
		}
	}

	private static void RegisterComClassWow64(RegistryKey root, string clsid, string progId, string className, string assemblyFullName, string codebaseUri)
	{
		using (RegistryKey registryKey = root.CreateSubKey("Software\\Classes\\Wow6432Node\\" + progId))
		{
			if (registryKey != null)
			{
				registryKey.SetValue("", className);
				using RegistryKey registryKey2 = registryKey.CreateSubKey("CLSID");
				registryKey2?.SetValue("", clsid);
			}
		}
		using RegistryKey registryKey3 = root.CreateSubKey("Software\\Classes\\Wow6432Node\\CLSID\\" + clsid);
		if (registryKey3 == null)
		{
			return;
		}
		registryKey3.SetValue("", className);
		using (RegistryKey registryKey4 = registryKey3.CreateSubKey("InprocServer32"))
		{
			if (registryKey4 != null)
			{
				registryKey4.SetValue("", "mscoree.dll");
				registryKey4.SetValue("ThreadingModel", "Both");
				registryKey4.SetValue("Class", className);
				registryKey4.SetValue("Assembly", assemblyFullName);
				registryKey4.SetValue("RuntimeVersion", "v4.0.30319");
				registryKey4.SetValue("CodeBase", codebaseUri);
				using RegistryKey registryKey5 = registryKey4.CreateSubKey("2.0.1.0");
				if (registryKey5 != null)
				{
					registryKey5.SetValue("Class", className);
					registryKey5.SetValue("Assembly", assemblyFullName);
					registryKey5.SetValue("RuntimeVersion", "v4.0.30319");
					registryKey5.SetValue("CodeBase", codebaseUri);
				}
			}
		}
		using (RegistryKey registryKey6 = registryKey3.CreateSubKey("ProgId"))
		{
			registryKey6?.SetValue("", progId);
		}
		using (registryKey3.CreateSubKey("Implemented Categories\\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"))
		{
		}
	}

	private static void UnregisterComClassWow64(RegistryKey root, string clsid, string progId)
	{
		try
		{
			root.DeleteSubKeyTree("Software\\Classes\\Wow6432Node\\" + progId, throwOnMissingSubKey: false);
		}
		catch
		{
		}
		try
		{
			root.DeleteSubKeyTree("Software\\Classes\\Wow6432Node\\CLSID\\" + clsid, throwOnMissingSubKey: false);
		}
		catch
		{
		}
	}
}
