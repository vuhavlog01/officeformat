#define TRACE
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using VFormat.Core.Enums;
using VFormat.License;
using VFormat.License.Models;
using VFormat.UI.Common;
using VFormat.WordEngine;

namespace VFormat.UI.TaskPane;

public class TaskPaneViewModel : ViewModelBase
{
	private readonly LicenseManager _licenseManager;

	private int _licenseCheckRunning;

	private readonly Dispatcher _uiDispatcher = Dispatcher.CurrentDispatcher;

	private RulePackType _selectedRulePack = RulePackType.ND30_2020;

	private PrintLayoutMode _selectedPrintLayout;

	private bool _isProcessing;

	private int _progressValue;

	private string _statusMessage = "Sẵn sàng chuẩn hóa văn bản.";

	private string _licenseStatusBadge = "Đang kiểm tra...";

	private bool _isLicenseValid;

	private bool _optPageSetup = true;

	private bool _optTypography = true;

	private bool _optTables = true;

	private bool _optPageNumbers = true;

	public Func<object> GetActiveDocument { get; set; }

	public Func<object> GetActiveSelection { get; set; }

	public Action OpenPersonalizationDialogAction { get; set; }

	public Action OpenLicenseDialogAction { get; set; }

	public Action<PrintLayoutMode> OnPrintLayoutChangedAction { get; set; }

	public Action<RulePackType> OnRulePackChangedAction { get; set; }

	public Action<StandardizeOptions> StandardizeAllAction { get; set; }

	public Action StandardizeSelectionAction { get; set; }

	public Action InsertHeaderAction { get; set; }

	public Action InsertSignerAction { get; set; }

	public List<RulePackOption> RulePackOptions { get; private set; }

	public List<PrintLayoutOption> PrintLayoutOptions { get; private set; }

	public RulePackType SelectedRulePack
	{
		get
		{
			return _selectedRulePack;
		}
		set
		{
			if (SetProperty(ref _selectedRulePack, value, "SelectedRulePack") && OnRulePackChangedAction != null)
			{
				OnRulePackChangedAction(value);
			}
		}
	}

	public PrintLayoutMode SelectedPrintLayout
	{
		get
		{
			return _selectedPrintLayout;
		}
		set
		{
			if (SetProperty(ref _selectedPrintLayout, value, "SelectedPrintLayout") && OnPrintLayoutChangedAction != null)
			{
				OnPrintLayoutChangedAction(value);
			}
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
			if (SetProperty(ref _isProcessing, value, "IsProcessing"))
			{
				OnPropertyChanged("CanExecuteActions");
			}
		}
	}

	public bool CanExecuteActions => !_isProcessing;

	public int ProgressValue
	{
		get
		{
			return _progressValue;
		}
		set
		{
			SetProperty(ref _progressValue, value, "ProgressValue");
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

	public string LicenseStatusBadge
	{
		get
		{
			return _licenseStatusBadge;
		}
		set
		{
			SetProperty(ref _licenseStatusBadge, value, "LicenseStatusBadge");
		}
	}

	public bool IsLicenseValid
	{
		get
		{
			return _isLicenseValid;
		}
		set
		{
			SetProperty(ref _isLicenseValid, value, "IsLicenseValid");
		}
	}

	public bool OptPageSetup
	{
		get
		{
			return _optPageSetup;
		}
		set
		{
			SetProperty(ref _optPageSetup, value, "OptPageSetup");
		}
	}

	public bool OptTypography
	{
		get
		{
			return _optTypography;
		}
		set
		{
			SetProperty(ref _optTypography, value, "OptTypography");
		}
	}

	public bool OptTables
	{
		get
		{
			return _optTables;
		}
		set
		{
			SetProperty(ref _optTables, value, "OptTables");
		}
	}

	public bool OptPageNumbers
	{
		get
		{
			return _optPageNumbers;
		}
		set
		{
			SetProperty(ref _optPageNumbers, value, "OptPageNumbers");
		}
	}

	public ICommand StandardizeAllCommand { get; private set; }

	public ICommand StandardizeSelectionCommand { get; private set; }

	public ICommand InsertHeaderCommand { get; private set; }

	public ICommand InsertSignerCommand { get; private set; }

	public ICommand OpenPersonalizationCommand { get; private set; }

	public ICommand OpenLicenseCommand { get; private set; }

	public ICommand OpenSpellCheckCommand { get; private set; }

	public ICommand OpenTemplateLibraryCommand { get; private set; }

	public Action OpenSpellCheckDialogAction { get; set; }

	public Action OpenTemplateLibraryDialogAction { get; set; }

	public TaskPaneViewModel()
	{
		_licenseManager = new LicenseManager();
		InitRulePackOptions();
		InitPrintLayoutOptions();
		InitCommands();
	}

	public StandardizeOptions BuildStandardizeOptions()
	{
		StandardizeOptions standardizeOptions = new StandardizeOptions();
		standardizeOptions.PageSetup = OptPageSetup;
		standardizeOptions.Typography = OptTypography;
		standardizeOptions.Tables = OptTables;
		standardizeOptions.PageNumbers = OptPageNumbers;
		return standardizeOptions;
	}

	private void InitRulePackOptions()
	{
		RulePackOptions = new List<RulePackOption>
		{
			new RulePackOption
			{
				Type = RulePackType.ND30_2020,
				Title = "Nghị định 30/2020/NĐ-CP (Hành chính)",
				Subtitle = "Font Times New Roman 14pt, giãn dòng 1.15, lề 20-20-30-15mm"
			},
			new RulePackOption
			{
				Type = RulePackType.HD05_VPTW,
				Title = "Hướng dẫn 05-HD/VPTW (Đảng bộ)",
				Subtitle = "Font 14pt, giãn dòng Exactly 20pt, trích yếu 5 gạch"
			},
			new RulePackOption
			{
				Type = RulePackType.TT68_2023_BQP,
				Title = "Thông tư 68/2023/TT-BQP (Quân đội)",
				Subtitle = "Chuẩn Bộ Quốc phòng, Cấp bậc quân hàm, Độ khẩn/mật"
			},
			new RulePackOption
			{
				Type = RulePackType.HD29_TWDTN,
				Title = "Hướng dẫn 29-HD/TWĐTN-VP (Đoàn TNCS Hồ Chí Minh)",
				Subtitle = "Tiêu đề ĐOÀN TNCS HỒ CHÍ MINH, dấu ***, tên loại 16pt, lề 20-20-30-20mm"
			}
		};
		SelectedRulePack = RulePackType.ND30_2020;
	}

	private void InitPrintLayoutOptions()
	{
		PrintLayoutOptions = new List<PrintLayoutOption>
		{
			new PrintLayoutOption
			{
				Type = PrintLayoutMode.SingleSided,
				Title = "Gửi file (1 mặt - Lề đơn)",
				Subtitle = "Lề cố định mọi trang (Trái 30, Phải 15/20mm). Chuẩn gửi mạng, xem trên máy."
			},
			new PrintLayoutOption
			{
				Type = PrintLayoutMode.TwoSidedMirror,
				Title = "In 2 mặt (Đối xứng)",
				Subtitle = "Mirror Margins: Lề gáy so le mặt trước - mặt sau theo HĐ 05/TT 68/NĐ 30."
			}
		};
		SelectedPrintLayout = PrintLayoutMode.SingleSided;
	}

	private void InitCommands()
	{
		StandardizeAllCommand = new RelayCommand(ExecuteStandardizeAll, () => CanExecuteActions);
		StandardizeSelectionCommand = new RelayCommand(ExecuteStandardizeSelection, () => CanExecuteActions);
		InsertHeaderCommand = new RelayCommand(ExecuteInsertHeader, () => CanExecuteActions);
		InsertSignerCommand = new RelayCommand(ExecuteInsertSigner, () => CanExecuteActions);
		OpenPersonalizationCommand = new RelayCommand(ExecuteOpenPersonalization);
		OpenLicenseCommand = new RelayCommand(ExecuteOpenLicense);
		OpenSpellCheckCommand = new RelayCommand(ExecuteOpenSpellCheck);
		OpenTemplateLibraryCommand = new RelayCommand(ExecuteOpenTemplateLibrary);
	}

	private void ExecuteOpenSpellCheck()
	{
		if (OpenSpellCheckDialogAction != null)
		{
			OpenSpellCheckDialogAction();
		}
	}

	private void ExecuteOpenTemplateLibrary()
	{
		if (OpenTemplateLibraryDialogAction != null)
		{
			OpenTemplateLibraryDialogAction();
		}
	}

	private void ExecuteOpenPersonalization()
	{
		if (OpenPersonalizationDialogAction != null)
		{
			OpenPersonalizationDialogAction();
		}
	}

	private void ExecuteOpenLicense()
	{
		if (OpenLicenseDialogAction != null)
		{
			OpenLicenseDialogAction();
		}
	}

	public void RefreshLicenseStatus()
	{
		try
		{
			ApplyLicenseStatus(_licenseManager.CheckLicenseAtStartup());
		}
		catch
		{
			ApplyLicenseStatus(null);
		}
	}

	public void RefreshLicenseStatusInBackground()
	{
		if (Interlocked.Exchange(ref _licenseCheckRunning, 1) == 1)
		{
			return;
		}
		Task.Factory.StartNew(() =>
		{
			LicenseCheckResult licenseCheckResult = null;
			try
			{
				return new LicenseManager().CheckLicenseAtStartup();
			}
			catch
			{
				return (LicenseCheckResult)null;
			}
		}, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default).ContinueWith((Task<LicenseCheckResult> t) =>
		{
			LicenseCheckResult state = ((t.Status == TaskStatus.RanToCompletion) ? t.Result : null);
			QueueLicenseStatus(state);
		}, CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.Default);
	}

	private void QueueLicenseStatus(LicenseCheckResult state)
	{
		try
		{
			if (_uiDispatcher.HasShutdownStarted || _uiDispatcher.HasShutdownFinished)
			{
				return;
			}
			_uiDispatcher.BeginInvoke((Action)(() =>
			{
				try
				{
					ApplyLicenseStatus(state);
				}
				catch (Exception ex2)
				{
					Trace.WriteLine("License UI update: " + ex2);
				}
			}));
		}
		catch (Exception ex)
		{
			Trace.WriteLine("License dispatcher: " + ex);
		}
		finally
		{
			Interlocked.Exchange(ref _licenseCheckRunning, 0);
		}
	}

	public void ApplyLicenseStatus(LicenseCheckResult state)
	{
		if (state == null)
		{
			LicenseStatusBadge = "Bản dùng thử";
			IsLicenseValid = false;
			return;
		}
		IsLicenseValid = state.IsSuccess && (state.Status == LicenseStatus.Active || state.Status == LicenseStatus.OfflineActive || state.Status == LicenseStatus.GracePeriod);
		if (!IsLicenseValid)
		{
			LicenseStatusBadge = "Chưa kích hoạt / Hết hạn";
		}
		else if (state.Status == LicenseStatus.GracePeriod)
		{
			LicenseStatusBadge = $"Bản quyền Pro (ngoại tuyến, còn {state.DaysLeft} ngày)";
		}
		else if (state.DaysLeft > 3650)
		{
			LicenseStatusBadge = "Bản quyền Pro Vĩnh viễn";
		}
		else
		{
			LicenseStatusBadge = $"Bản quyền Pro ({state.DaysLeft} ngày)";
		}
	}

	public bool EnsureLicenseValid()
	{
		RefreshLicenseStatus();
		if (!IsLicenseValid)
		{
			StatusMessage = "Vui lòng kích hoạt bản quyền OfficeFormat để sử dụng tính năng này!";
			ExecuteOpenLicense();
			return false;
		}
		return true;
	}

	private void ExecuteStandardizeAll()
	{
		StandardizeOptions options = BuildStandardizeOptions();
		if (options.IsEmpty)
		{
			StatusMessage = "Hãy chọn ít nhất một mục trong \"Các mục chuẩn hóa\".";
			return;
		}
		RunRibbonAction(() =>
		{
			StandardizeAllAction(options);
		}, StandardizeAllAction != null);
	}

	private void ExecuteStandardizeSelection()
	{
		RunRibbonAction(StandardizeSelectionAction, StandardizeSelectionAction != null);
	}

	private void ExecuteInsertHeader()
	{
		RunRibbonAction(InsertHeaderAction, InsertHeaderAction != null);
	}

	private void ExecuteInsertSigner()
	{
		RunRibbonAction(InsertSignerAction, InsertSignerAction != null);
	}

	private void RunRibbonAction(Action action, bool available)
	{
		if (!available || action == null)
		{
			StatusMessage = "Chức năng chưa sẵn sàng. Vui lòng mở lại Word.";
		}
		else
		{
			if (IsProcessing)
			{
				return;
			}
			IsProcessing = true;
			try
			{
				action();
			}
			catch (Exception ex)
			{
				StatusMessage = "Lỗi: " + ex.Message;
			}
			finally
			{
				IsProcessing = false;
			}
		}
	}
}
