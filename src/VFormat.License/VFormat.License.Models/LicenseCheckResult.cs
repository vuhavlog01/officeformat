using System;

namespace VFormat.License.Models;

public class LicenseCheckResult
{
	public LicenseStatus Status { get; set; }

	public string StatusCodeString { get; set; }

	public string LicenseKey { get; set; }

	public string Hwid { get; set; }

	public string CustomerName { get; set; }

	public string PlanCode { get; set; }

	public int DaysLeft { get; set; }

	public string ExpiresAt { get; set; }

	public string Message { get; set; }

	public bool IsSuccess { get; set; }

	public bool IsOffline { get; set; }

	public DateTime VerifiedAt { get; set; }

	public bool IsLicensed
	{
		get
		{
			if (Status != LicenseStatus.Active && Status != LicenseStatus.OfflineActive)
			{
				return Status == LicenseStatus.GracePeriod;
			}
			return true;
		}
	}

	public LicenseCheckResult()
	{
		Status = LicenseStatus.NoLicense;
		StatusCodeString = "NO_LICENSE";
		LicenseKey = string.Empty;
		Hwid = string.Empty;
		CustomerName = string.Empty;
		PlanCode = string.Empty;
		DaysLeft = 0;
		ExpiresAt = string.Empty;
		Message = string.Empty;
		IsSuccess = false;
		IsOffline = false;
		VerifiedAt = DateTime.Now;
	}
}
