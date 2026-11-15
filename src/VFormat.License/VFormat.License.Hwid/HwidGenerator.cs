using System;
using System.Management;

namespace VFormat.License.Hwid;

public static class HwidGenerator
{
	private static string _cachedHwid;

	public static string GetMachineHwid()
	{
		if (!string.IsNullOrEmpty(_cachedHwid))
		{
			return _cachedHwid;
		}
		string text = "";
		string text2 = "";
		string text3 = "";
		try
		{
			using (ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
			{
				foreach (ManagementObject item in managementObjectSearcher.Get())
				{
					if (item["ProcessorId"] != null)
					{
						text = item["ProcessorId"].ToString().Trim();
						break;
					}
				}
			}
			using (ManagementObjectSearcher managementObjectSearcher2 = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
			{
				foreach (ManagementObject item2 in managementObjectSearcher2.Get())
				{
					if (item2["SerialNumber"] != null)
					{
						text2 = item2["SerialNumber"].ToString().Trim();
						break;
					}
				}
			}
			using (ManagementObjectSearcher managementObjectSearcher3 = new ManagementObjectSearcher("SELECT VolumeSerialNumber FROM Win32_LogicalDisk WHERE DeviceID='C:'"))
			{
				foreach (ManagementObject item3 in managementObjectSearcher3.Get())
				{
					if (item3["VolumeSerialNumber"] != null)
					{
						text3 = item3["VolumeSerialNumber"].ToString().Trim();
						break;
					}
				}
			}
			if (string.IsNullOrEmpty(text))
			{
				text = "CPU0";
			}
			if (string.IsNullOrEmpty(text2))
			{
				text2 = "MB0";
			}
			if (string.IsNullOrEmpty(text3))
			{
				text3 = "DSK0";
			}
			string text4 = $"{text}-{text2}-{text3}".ToUpperInvariant().Replace(" ", "").Replace("/", "")
				.Replace("\\", "");
			if (text4.Length > 28)
			{
				_cachedHwid = text4.Substring(0, 28);
			}
			else
			{
				_cachedHwid = text4;
			}
			return _cachedHwid;
		}
		catch
		{
			string arg = Environment.MachineName.ToUpperInvariant();
			string arg2 = Environment.UserName.ToUpperInvariant();
			_cachedHwid = $"PC-{arg}-{arg2}";
			return _cachedHwid;
		}
	}
}
