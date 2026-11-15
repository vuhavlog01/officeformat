namespace VFormat.License.Models;

public class DeviceTransferStatus
{
	public DeviceTransferState State { get; set; }

	public string ServerCode { get; set; }

	public bool CreditsKnown { get; set; }

	public int CreditsRemaining { get; set; }

	public bool CanBuyCredits { get; set; }

	public bool AwaitingNewDevice { get; set; }

	public string RequestState { get; set; }

	public string Message { get; set; }

	public bool IsNetworkError { get; set; }

	public DeviceTransferStatus()
	{
		State = DeviceTransferState.Unknown;
		ServerCode = string.Empty;
		RequestState = string.Empty;
		Message = string.Empty;
	}

	public static DeviceTransferStatus FromState(DeviceTransferState state, string message)
	{
		DeviceTransferStatus deviceTransferStatus = new DeviceTransferStatus();
		deviceTransferStatus.State = state;
		deviceTransferStatus.Message = message ?? string.Empty;
		return deviceTransferStatus;
	}

	public static DeviceTransferStatus FromServerCode(string code)
	{
		switch ((code ?? string.Empty).ToUpperInvariant())
		{
		case "ELIGIBLE":
		{
			DeviceTransferStatus deviceTransferStatus2 = new DeviceTransferStatus();
			deviceTransferStatus2.State = DeviceTransferState.Eligible;
			deviceTransferStatus2.ServerCode = code;
			return deviceTransferStatus2;
		}
		case "NO_CREDITS":
		{
			DeviceTransferStatus deviceTransferStatus3 = new DeviceTransferStatus();
			deviceTransferStatus3.State = DeviceTransferState.NoCredits;
			deviceTransferStatus3.ServerCode = code;
			return deviceTransferStatus3;
		}
		case "OFFLINE_LICENSE":
		{
			DeviceTransferStatus deviceTransferStatus9 = new DeviceTransferStatus();
			deviceTransferStatus9.State = DeviceTransferState.OfflineLicense;
			deviceTransferStatus9.ServerCode = code;
			return deviceTransferStatus9;
		}
		case "TIMED_PLAN":
		{
			DeviceTransferStatus deviceTransferStatus8 = new DeviceTransferStatus();
			deviceTransferStatus8.State = DeviceTransferState.TimedPlan;
			deviceTransferStatus8.ServerCode = code;
			return deviceTransferStatus8;
		}
		case "DEVICE_NOT_BOUND":
		{
			DeviceTransferStatus deviceTransferStatus7 = new DeviceTransferStatus();
			deviceTransferStatus7.State = DeviceTransferState.NotActivated;
			deviceTransferStatus7.ServerCode = code;
			return deviceTransferStatus7;
		}
		case "DEVICE_TRANSFERRED":
		{
			DeviceTransferStatus deviceTransferStatus6 = new DeviceTransferStatus();
			deviceTransferStatus6.State = DeviceTransferState.Transferred;
			deviceTransferStatus6.ServerCode = code;
			return deviceTransferStatus6;
		}
		case "TRANSFER_DONE":
		{
			DeviceTransferStatus deviceTransferStatus5 = new DeviceTransferStatus();
			deviceTransferStatus5.State = DeviceTransferState.TransferDone;
			deviceTransferStatus5.ServerCode = code;
			return deviceTransferStatus5;
		}
		case "KEY_LOCKED":
		case "KEY_INVALID":
		case "DEVICE_BLOCKED":
		case "PRODUCT_NOT_SUPPORTED":
		{
			DeviceTransferStatus deviceTransferStatus4 = new DeviceTransferStatus();
			deviceTransferStatus4.State = DeviceTransferState.Blocked;
			deviceTransferStatus4.ServerCode = code;
			return deviceTransferStatus4;
		}
		default:
		{
			DeviceTransferStatus deviceTransferStatus = new DeviceTransferStatus();
			deviceTransferStatus.State = DeviceTransferState.Unknown;
			deviceTransferStatus.ServerCode = code ?? string.Empty;
			return deviceTransferStatus;
		}
		}
	}
}
