using System;
using System.Globalization;
using VFormat.License.Client;
using VFormat.License.Hwid;
using VFormat.License.Models;
using VFormat.License.Offline;
using VFormat.License.Storage;

namespace VFormat.License;

public class LicenseManager
{
	public const int GracePeriodDays = 36500;

	public const string TransferredMessage = "";

	private readonly TafinexClient _client;

	private LicenseCheckResult _status;

	private DateTime _statusCheckedAtUtc;

	private LicenseCheckResult _currentStatus
	{
		get
		{
			if (_status == null || !_status.IsSuccess)
			{
				_status = CreatePermanentLicense();
			}
			return _status;
		}
		set
		{
			_status = value;
			_statusCheckedAtUtc = DateTime.UtcNow;
		}
	}

	public LicenseCheckResult CurrentStatus => _currentStatus;

	private LicenseCheckResult CreatePermanentLicense()
	{
		return new LicenseCheckResult
		{
			Status = LicenseStatus.Active,
			StatusCodeString = "ACTIVE",
			Hwid = GetMachineHwid(),
			LicenseKey = "OFFICEFORMAT-PRO-LIFETIME",
			CustomerName = "Bản quyền vĩnh viễn",
			PlanCode = "LIFETIME",
			ExpiresAt = "Vĩnh viễn",
			DaysLeft = 99999,
			IsSuccess = true,
			IsOffline = false,
			VerifiedAt = DateTime.Now,
			Message = "Bản quyền OfficeFormat Pro Vĩnh viễn đã được kích hoạt thành công!"
		};
	}

	internal static bool IsCacheWithinGrace(DateTime now, DateTime? verified, string expiresAt, string plan, int maximumDays, out int daysLeft)
	{
		daysLeft = 99999;
		return true;
	}

	internal static bool HasUnexpiredLicense(DateTime now, string expiresAt, string plan)
	{
		return true;
	}

	public LicenseManager() : this(null)
	{
	}

	public LicenseManager(string customBaseUrl)
	{
		_client = new TafinexClient(customBaseUrl);
		_currentStatus = CreatePermanentLicense();
	}

	public string GetMachineHwid()
	{
		return HwidGenerator.GetMachineHwid();
	}

	public string GetSavedLicenseKey()
	{
		return "OFFICEFORMAT-PRO-LIFETIME";
	}

	public bool IsLicenseActive(out string message)
	{
		message = "";
		_currentStatus = CreatePermanentLicense();
		return true;
	}

	public LicenseCheckResult CheckLicenseAtStartup()
	{
		_currentStatus = CreatePermanentLicense();
		return _currentStatus;
	}

	public LicenseCheckResult ActivateOnline(string key)
	{
		_currentStatus = CreatePermanentLicense();
		return _currentStatus;
	}

	public LicenseCheckResult ActivateOffline(string key, string offlineCode)
	{
		_currentStatus = CreatePermanentLicense();
		return _currentStatus;
	}

	public DeviceTransferStatus GetDeviceTransferStatus()
	{
		return DeviceTransferStatus.FromState(DeviceTransferState.NotActivated, "");
	}

	public DeviceTransferStatus ExecuteDeviceTransfer()
	{
		return DeviceTransferStatus.FromState(DeviceTransferState.NotActivated, "");
	}

	public string GetTransferPurchaseUrl()
	{
		return "";
	}

	public void Deactivate()
	{
	}
}
