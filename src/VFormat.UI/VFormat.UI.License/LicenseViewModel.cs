using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using Microsoft.Win32;
using VFormat.License;
using VFormat.License.Hwid;
using VFormat.License.Models;
using VFormat.License.Offline;
using VFormat.UI.Common;

namespace VFormat.UI.License;

public class LicenseViewModel : ViewModelBase
{
	public const string BuyLicensePromptText = "Chưa có mã bản quyền? Mua bản quyền OfficeFormat để kích hoạt và sử dụng các tính năng.";

	public const string TransferConfirmTitle = "Xác nhận đổi máy";

	public const string TransferConfirmMessage = "Thao tác này sẽ ngắt bản quyền OfficeFormat trên máy tính hiện tại và sử dụng 01 lượt đổi máy.\n\nSau khi hoàn tất, bạn có thể nhập mã bản quyền trên máy tính mới để tiếp tục sử dụng.\n\nBạn có muốn tiếp tục không?";

	public const string TransferSuccessMessage = "Đổi máy thành công. Bản quyền OfficeFormat trên máy tính này đã được ngắt. Bạn có thể nhập mã bản quyền trên máy tính mới để tiếp tục sử dụng.";

	public const string NoCreditsMessage = "Bạn đã sử dụng hết lượt đổi máy. Bạn có thể mua thêm với giá 89.000 đồng/lượt để chuyển bản quyền sang máy tính mới.";

	public const string NetworkTransferMessage = "Chưa thể kiểm tra số lượt đổi máy. Vui lòng kết nối Internet và thử lại.";

	private readonly LicenseManager _licenseManager;

	private string _hwid;

	private string _licenseKey;

	private string _offlineCode;

	private string _licenseStatus;

	private string _expirationInfo;

	private string _statusMessage;

	private bool _isSuccess;

	private bool _isProcessing;

	private readonly Dispatcher _dispatcher;

	private bool _showBuyLicensePrompt;

	private string _transferStatusText;

	private DeviceTransferState _transferState;

	private bool _transferChecked;

	private bool _showBuyTransferButton;

	public Action CloseAction { get; set; }

	public Window OwnerWindow { get; set; }

	public bool ShowBuyLicensePrompt
	{
		get
		{
			return _showBuyLicensePrompt;
		}
		set
		{
			SetProperty(ref _showBuyLicensePrompt, value, "ShowBuyLicensePrompt");
		}
	}

	public string TransferStatusText
	{
		get
		{
			return _transferStatusText;
		}
		set
		{
			SetProperty(ref _transferStatusText, value, "TransferStatusText");
		}
	}

	public bool ShowBuyTransferButton
	{
		get
		{
			return _showBuyTransferButton;
		}
		set
		{
			SetProperty(ref _showBuyTransferButton, value, "ShowBuyTransferButton");
		}
	}

	public ICommand BuyLicenseCommand { get; private set; }

	public ICommand DeviceTransferCommand { get; private set; }

	public ICommand RefreshCommand { get; private set; }

	public ICommand BuyTransferCreditCommand { get; private set; }

	public string Hwid
	{
		get
		{
			return _hwid;
		}
		set
		{
			SetProperty(ref _hwid, value, "Hwid");
		}
	}

	public string LicenseKey
	{
		get
		{
			return _licenseKey;
		}
		set
		{
			SetProperty(ref _licenseKey, value, "LicenseKey");
		}
	}

	public string OfflineCode
	{
		get
		{
			return _offlineCode;
		}
		set
		{
			SetProperty(ref _offlineCode, value, "OfflineCode");
		}
	}

	public string LicenseStatus
	{
		get
		{
			return _licenseStatus;
		}
		set
		{
			SetProperty(ref _licenseStatus, value, "LicenseStatus");
		}
	}

	public string ExpirationInfo
	{
		get
		{
			return _expirationInfo;
		}
		set
		{
			SetProperty(ref _expirationInfo, value, "ExpirationInfo");
		}
	}

	public string StatusMessage
	{
		get
		{
			return _statusMessage;
		}
		set
		{
			SetProperty(ref _statusMessage, value, "StatusMessage");
		}
	}

	public bool IsSuccess
	{
		get
		{
			return _isSuccess;
		}
		set
		{
			SetProperty(ref _isSuccess, value, "IsSuccess");
		}
	}

	public bool IsProcessing
	{
		get
		{
			return _isProcessing;
		}
		set
		{
			SetProperty(ref _isProcessing, value, "IsProcessing");
		}
	}

	public ICommand CopyHwidCommand { get; private set; }

	public ICommand ActivateOnlineCommand { get; private set; }

	public ICommand ActivateOfflineCommand { get; private set; }

	public ICommand ImportActivationFileCommand { get; private set; }

	public ICommand CloseCommand { get; private set; }

	private IntPtr OwnerHandle
	{
		get
		{
			try
			{
				return (OwnerWindow != null) ? new WindowInteropHelper(OwnerWindow).Handle : IntPtr.Zero;
			}
			catch
			{
				return IntPtr.Zero;
			}
		}
	}

	public LicenseViewModel()
	{
		_dispatcher = Dispatcher.CurrentDispatcher;
		_licenseManager = new LicenseManager();
		Hwid = HwidGenerator.GetMachineHwid();
		TransferStatusText = "Đang kiểm tra số lượt đổi máy...";
		LicenseKey = _licenseManager.GetSavedLicenseKey();
		RefreshStatus();
		CopyHwidCommand = new RelayCommand(ExecuteCopyHwid);
		Action execute = ExecuteActivateOnline;
		Func<bool> canExecute = () => !IsProcessing && !string.IsNullOrEmpty(LicenseKey);
		ActivateOnlineCommand = new RelayCommand(execute, canExecute);
		ActivateOfflineCommand = new RelayCommand(ExecuteActivateOffline, () => !IsProcessing && !string.IsNullOrEmpty(OfflineCode) && !string.IsNullOrEmpty(LicenseKey));
		ImportActivationFileCommand = new RelayCommand(ExecuteImportActivationFile, () => !IsProcessing);
		CloseCommand = new RelayCommand(ExecuteClose);
		BuyLicenseCommand = new RelayCommand(ExecuteBuyLicense);
		DeviceTransferCommand = new RelayCommand(ExecuteDeviceTransfer, () => !IsProcessing && _transferChecked && (_transferState == DeviceTransferState.Eligible || _transferState == DeviceTransferState.NoCredits || _transferState == DeviceTransferState.Unknown));
		RefreshCommand = new RelayCommand(ExecuteRefresh, () => !IsProcessing);
		BuyTransferCreditCommand = new RelayCommand(ExecuteBuyTransferCredit, () => !IsProcessing);
	}

	private void ExecuteClose()
	{
		if (CloseAction != null)
		{
			CloseAction();
		}
	}

	public void RefreshStatus()
	{
		LicenseCheckResult licenseCheckResult = _licenseManager.CheckLicenseAtStartup();
		if (licenseCheckResult.IsSuccess && (licenseCheckResult.Status == VFormat.License.Models.LicenseStatus.Active || licenseCheckResult.Status == VFormat.License.Models.LicenseStatus.OfflineActive || licenseCheckResult.Status == VFormat.License.Models.LicenseStatus.GracePeriod))
		{
			IsSuccess = true;
			LicenseStatus = "ĐÃ KÍCH HOẠT BẢN QUYỀN THÀNH CÔNG";
			if (licenseCheckResult.DaysLeft > 3650)
			{
				ExpirationInfo = "Bản quyền Vĩnh viễn (Trọn đời)";
			}
			else
			{
				ExpirationInfo = $"Gói bản quyền: {licenseCheckResult.PlanCode} (Còn {licenseCheckResult.DaysLeft} ngày)";
			}
		}
		else
		{
			IsSuccess = false;
			LicenseStatus = "CHƯA KÍCH HOẠT BẢN QUYỀN";
			ExpirationInfo = licenseCheckResult.Message ?? "Vui lòng nhập Mã bản quyền để kích hoạt sử dụng không giới hạn.";
		}
		ApplyBuyPrompt(licenseCheckResult);
	}

	private void ApplyBuyPrompt(LicenseCheckResult state)
	{
		if (state == null || state.IsLicensed)
		{
			ShowBuyLicensePrompt = false;
			return;
		}
		string text = state.StatusCodeString ?? "";
		if (state.Status == VFormat.License.Models.LicenseStatus.Expired)
		{
			goto IL_0061;
		}
		switch (text)
		{
		case "NO_LICENSE":
		case "KEY_INVALID":
		case "DEVICE_TRANSFERRED":
			goto IL_0061;
		}
		int showBuyLicensePrompt = ((text == "DEVICE_RESET") ? 1 : 0);
		goto IL_0062;
		IL_0061:
		showBuyLicensePrompt = 1;
		goto IL_0062;
		IL_0062:
		ShowBuyLicensePrompt = (byte)showBuyLicensePrompt != 0;
	}

	private void ExecuteBuyLicense()
	{
		ExternalLinkHelper.OpenBuyLicensePage(OwnerWindow);
	}

	private void RunInBackground<T>(string busyText, Func<T> work, Action<T> done)
	{
		if (IsProcessing)
		{
			return;
		}
		IsProcessing = true;
		if (!string.IsNullOrEmpty(busyText))
		{
			StatusMessage = busyText;
		}
		Task.Factory.StartNew(() =>
		{
			try
			{
				return Tuple.Create<T, Exception>(work(), null);
			}
			catch (Exception item)
			{
				return Tuple.Create<T, Exception>(default, item);
			}
		}, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).ContinueWith((Task<Tuple<T, Exception>> t) =>
		{
			_dispatcher.BeginInvoke((Action)(() =>
			{
				try
				{
					Tuple<T, Exception> result = t.Result;
					if (result.Item2 != null)
					{
						StatusMessage = "Đã xảy ra lỗi khi kết nối máy chủ. Vui lòng thử lại.";
					}
					else
					{
						done(result.Item1);
					}
				}
				catch (Exception ex)
				{
					StatusMessage = "Không thể cập nhật trạng thái: " + ex.Message;
				}
				finally
				{
					IsProcessing = false;
					CommandManager.InvalidateRequerySuggested();
				}
			}));
		}, TaskScheduler.Default);
	}

	public void LoadTransferStatus()
	{
		RunInBackground(null, () => _licenseManager.GetDeviceTransferStatus(), ApplyTransferStatus);
	}

	private void ExecuteRefresh()
	{
		TransferStatusText = "Đang kiểm tra số lượt đổi máy...";
		RunInBackground("Đang làm mới thông tin bản quyền...", () =>
		{
			DeviceTransferStatus deviceTransferStatus = _licenseManager.GetDeviceTransferStatus();
			return Tuple.Create(_licenseManager.CheckLicenseAtStartup(), deviceTransferStatus);
		}, (Tuple<LicenseCheckResult, DeviceTransferStatus> r) =>
		{
			ApplyLicenseCheck(r.Item1);
			ApplyTransferStatus(r.Item2);
			StatusMessage = "Đã làm mới thông tin bản quyền.";
		});
	}

	private void ApplyLicenseCheck(LicenseCheckResult state)
	{
		if (state != null && state.IsSuccess && state.IsLicensed)
		{
			IsSuccess = true;
			LicenseStatus = "ĐÃ KÍCH HOẠT BẢN QUYỀN THÀNH CÔNG";
			ExpirationInfo = ((state.DaysLeft > 3650) ? "Bản quyền Vĩnh viễn (Trọn đời)" : $"Gói bản quyền: {state.PlanCode} (Còn {state.DaysLeft} ngày)");
		}
		else
		{
			IsSuccess = false;
			LicenseStatus = "CHƯA KÍCH HOẠT BẢN QUYỀN";
			ExpirationInfo = state?.Message ?? "Vui lòng nhập Mã bản quyền để kích hoạt sử dụng không giới hạn.";
		}
		ApplyBuyPrompt(state);
	}

	private void ApplyTransferStatus(DeviceTransferStatus st)
	{
		if (st == null)
		{
			DeviceTransferStatus deviceTransferStatus = new DeviceTransferStatus();
			deviceTransferStatus.IsNetworkError = true;
			st = deviceTransferStatus;
		}
		_transferChecked = true;
		_transferState = ((st.State == DeviceTransferState.TransferDone) ? DeviceTransferState.Transferred : st.State);
		TransferStatusText = TransferText(st);
		ShowBuyTransferButton = st.CreditsKnown && st.CanBuyCredits && (st.State == DeviceTransferState.Eligible || st.State == DeviceTransferState.NoCredits);
		CommandManager.InvalidateRequerySuggested();
	}

	public static string TransferText(DeviceTransferStatus st)
	{
		switch (st.State)
		{
		case DeviceTransferState.Eligible:
			return $"Số lượt đổi máy còn lại: {st.CreditsRemaining} lượt.";
		case DeviceTransferState.NoCredits:
			return "Bạn đã sử dụng hết lượt đổi máy.";
		case DeviceTransferState.OfflineLicense:
			return "Bản quyền kích hoạt offline không hỗ trợ đổi máy.";
		case DeviceTransferState.TimedPlan:
			return "Tính năng đổi máy chỉ áp dụng cho bản vĩnh viễn kích hoạt online.";
		case DeviceTransferState.NotActivated:
			return "Kích hoạt bản quyền OfficeFormat trên máy này (Kích Hoạt Online) để xem thông tin đổi máy.";
		case DeviceTransferState.Transferred:
		case DeviceTransferState.TransferDone:
			return "Bản quyền OfficeFormat trên máy tính này đã được chuyển sang máy khác. Bạn có thể nhập mã bản quyền trên máy tính mới để tiếp tục sử dụng.";
		case DeviceTransferState.Blocked:
			if (string.IsNullOrEmpty(st.Message))
			{
				return "Bản quyền này không thể đổi máy. Vui lòng liên hệ bộ phận hỗ trợ.";
			}
			return st.Message;
		default:
			return "Chưa thể kiểm tra số lượt đổi máy. Vui lòng kết nối Internet và thử lại.";
		}
	}

	private void ExecuteDeviceTransfer()
	{
		TransferStatusText = "Đang kiểm tra với máy chủ...";
		RunInBackground("Đang kiểm tra điều kiện đổi máy...", () => _licenseManager.GetDeviceTransferStatus(), (DeviceTransferStatus st) =>
		{
			ApplyTransferStatus(st);
			StatusMessage = "";
			switch (st.State)
			{
			case DeviceTransferState.Eligible:
				AskAndTransfer();
				break;
			case DeviceTransferState.NoCredits:
				if (ModernMessageDialog.ShowConfirm("Bạn đã sử dụng hết lượt đổi máy. Bạn có thể mua thêm với giá 89.000 đồng/lượt để chuyển bản quyền sang máy tính mới.", "Đổi máy", "Mua lượt đổi máy OfficeFormat", "Đóng", OwnerHandle))
				{
					ExecuteBuyTransferCredit();
				}
				break;
			case DeviceTransferState.Unknown:
				ModernMessageDialog.ShowWarning("Chưa thể kiểm tra số lượt đổi máy. Vui lòng kết nối Internet và thử lại.", "Đổi máy", OwnerHandle);
				break;
			default:
				ModernMessageDialog.ShowInfo(TransferText(st), "Đổi máy", OwnerHandle);
				break;
			}
		});
	}

	private void AskAndTransfer()
	{
		if (!ModernMessageDialog.ShowConfirm("Thao tác này sẽ ngắt bản quyền OfficeFormat trên máy tính hiện tại và sử dụng 01 lượt đổi máy.\n\nSau khi hoàn tất, bạn có thể nhập mã bản quyền trên máy tính mới để tiếp tục sử dụng.\n\nBạn có muốn tiếp tục không?", "Xác nhận đổi máy", "Xác nhận đổi máy", "Hủy", OwnerHandle))
		{
			StatusMessage = "Đã hủy. Lượt đổi máy không bị trừ.";
			return;
		}
		RunInBackground("Đang đổi máy, vui lòng không tắt Word...", () => _licenseManager.ExecuteDeviceTransfer(), (DeviceTransferStatus st) =>
		{
			if (st.State == DeviceTransferState.TransferDone)
			{
				ApplyTransferStatus(st);
				ApplyLicenseCheck(_licenseManager.CurrentStatus);
				StatusMessage = "Đổi máy thành công.";
				ModernMessageDialog.ShowSuccess("Đổi máy thành công. Bản quyền OfficeFormat trên máy tính này đã được ngắt. Bạn có thể nhập mã bản quyền trên máy tính mới để tiếp tục sử dụng.", "Đổi máy", OwnerHandle);
			}
			else if (st.IsNetworkError)
			{
				TransferStatusText = "Chưa nhận được kết quả đổi máy từ máy chủ.";
				_transferChecked = true;
				_transferState = DeviceTransferState.Unknown;
				ModernMessageDialog.ShowWarning("Chưa nhận được kết quả đổi máy từ máy chủ. Vui lòng kết nối Internet rồi bấm \"Làm mới\" để kiểm tra lại. Nếu máy chủ đã thực hiện, bản quyền trên máy này sẽ được ngắt; nếu chưa, lượt đổi máy không bị trừ.", "Đổi máy", OwnerHandle);
			}
			else
			{
				ApplyTransferStatus(st);
				if (st.State == DeviceTransferState.NoCredits)
				{
					if (ModernMessageDialog.ShowConfirm("Bạn đã sử dụng hết lượt đổi máy. Bạn có thể mua thêm với giá 89.000 đồng/lượt để chuyển bản quyền sang máy tính mới.", "Đổi máy", "Mua lượt đổi máy OfficeFormat", "Đóng", OwnerHandle))
					{
						ExecuteBuyTransferCredit();
					}
				}
				else
				{
					ModernMessageDialog.ShowWarning("Chưa đổi được máy. " + TransferText(st), "Đổi máy", OwnerHandle);
				}
			}
		});
	}

	private void ExecuteBuyTransferCredit()
	{
		RunInBackground("Đang mở trang mua lượt đổi máy...", () => _licenseManager.GetTransferPurchaseUrl(), (string url) =>
		{
			StatusMessage = "";
			ExternalLinkHelper.TryOpen(url, "Không thể mở trang mua lượt đổi máy. Bạn có thể sao chép liên kết bên dưới và mở bằng trình duyệt.", OwnerWindow);
		});
	}

	private void ExecuteCopyHwid()
	{
		try
		{
			Clipboard.SetText(Hwid);
			StatusMessage = "Đã sao chép Mã máy (HWID) vào bộ nhớ tạm!";
		}
		catch (Exception ex)
		{
			StatusMessage = "Không thể sao chép: " + ex.Message;
		}
	}

	private void ExecuteActivateOnline()
	{
		IsProcessing = true;
		StatusMessage = "Đang kết nối máy chủ để kích hoạt...";
		try
		{
			string licenseKey = ((LicenseKey != null) ? LicenseKey.Trim() : "");
			LicenseCheckResult licenseCheckResult = _licenseManager.ActivateOnline(licenseKey);
			if (licenseCheckResult.IsSuccess)
			{
				IsSuccess = true;
				StatusMessage = licenseCheckResult.Message;
				RefreshStatus();
			}
			else
			{
				IsSuccess = false;
				StatusMessage = "Kích hoạt thất bại: " + licenseCheckResult.Message;
			}
		}
		catch (Exception ex)
		{
			IsSuccess = false;
			StatusMessage = "Lỗi kết nối: " + ex.Message;
		}
		finally
		{
			IsProcessing = false;
		}
		if (IsSuccess)
		{
			LoadTransferStatus();
		}
	}

	private void ExecuteImportActivationFile()
	{
		OpenFileDialog openFileDialog = new OpenFileDialog();
		openFileDialog.Title = "Nạp tệp kích hoạt OfficeFormat";
		openFileDialog.Filter = "Tệp kích hoạt OfficeFormat (*.vflic)|*.vflic";
		openFileDialog.CheckFileExists = true;
		openFileDialog.Multiselect = false;
		OpenFileDialog openFileDialog2 = openFileDialog;
		if (openFileDialog2.ShowDialog() != true)
		{
			return;
		}
		IsProcessing = true;
		StatusMessage = "Đang kiểm tra tệp kích hoạt trên máy này...";
		try
		{
			OfflineActivationFile offlineActivationFile = OfflineActivationFile.Load(openFileDialog2.FileName, Hwid);
			LicenseCheckResult licenseCheckResult = _licenseManager.ActivateOffline(offlineActivationFile.LicenseKey, offlineActivationFile.OfflineCode);
			if (licenseCheckResult.IsSuccess)
			{
				LicenseKey = offlineActivationFile.LicenseKey;
				OfflineCode = offlineActivationFile.OfflineCode;
				RefreshStatus();
				StatusMessage = "Đã nạp tệp kích hoạt thành công. " + licenseCheckResult.Message;
			}
			else
			{
				StatusMessage = "Không thể kích hoạt: " + licenseCheckResult.Message;
			}
		}
		catch (Exception ex)
		{
			StatusMessage = "Không thể nạp tệp kích hoạt: " + ex.Message;
		}
		finally
		{
			IsProcessing = false;
		}
	}

	private void ExecuteActivateOffline()
	{
		IsProcessing = true;
		StatusMessage = "Đang kiểm tra mã kích hoạt ngoại tuyến...";
		try
		{
			string text = ((LicenseKey != null) ? LicenseKey.Trim() : "");
			string text2 = ((OfflineCode != null) ? OfflineCode.Trim() : "");
			if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2))
			{
				IsSuccess = false;
				StatusMessage = "Vui lòng nhập cả Mã bản quyền (Key) và Mã kích hoạt ngoại tuyến!";
				return;
			}
			LicenseCheckResult licenseCheckResult = _licenseManager.ActivateOffline(text, text2);
			if (licenseCheckResult.IsSuccess)
			{
				IsSuccess = true;
				StatusMessage = licenseCheckResult.Message;
				RefreshStatus();
			}
			else
			{
				IsSuccess = false;
				StatusMessage = "Kích hoạt thất bại: " + licenseCheckResult.Message;
			}
		}
		catch (Exception ex)
		{
			IsSuccess = false;
			StatusMessage = "Lỗi ngoại tuyến: " + ex.Message;
		}
		finally
		{
			IsProcessing = false;
		}
	}
}
