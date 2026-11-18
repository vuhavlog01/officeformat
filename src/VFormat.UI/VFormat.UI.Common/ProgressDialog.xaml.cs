using System;
using System.Security.Permissions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace VFormat.UI.Common;

public partial class ProgressDialog : Window, IComponentConnector
{
	public double CurrentFillRatio
	{
		get
		{
			if (colFill == null || colEmpty == null)
			{
				return 0.0;
			}
			double value = colFill.Width.Value;
			double value2 = colEmpty.Width.Value;
			double num = value + value2;
			if (!(num > 0.0))
			{
				return 0.0;
			}
			return value / num * 100.0;
		}
	}

	public bool IsCancelRequested { get; private set; }

	public event Action CancelRequested;

	public ProgressDialog()
	{
		InitializeComponent();
	}

	public ProgressDialog(string title, string statusMessage = "")
		: this()
	{
		if (!string.IsNullOrEmpty(title))
		{
			lblTitle.Text = title;
		}
		if (!string.IsNullOrEmpty(statusMessage))
		{
			lblStatus.Text = statusMessage;
		}
	}

	public void UpdateProgress(int percent, string statusMessage, string title = null)
	{
		if (!Dispatcher.CheckAccess())
		{
			Dispatcher.Invoke(() =>
			{
				UpdateProgress(percent, statusMessage, title);
			});
			return;
		}
		if (percent < 0)
		{
			percent = 0;
		}
		if (percent > 100)
		{
			percent = 100;
		}
		if (!string.IsNullOrEmpty(title))
		{
			lblTitle.Text = title;
		}
		if (!string.IsNullOrEmpty(statusMessage))
		{
			lblStatus.Text = statusMessage;
		}
		lblPercent.Text = $"Tiến trình: {percent}%";
		if (colFill != null && colEmpty != null)
		{
			if (percent <= 0)
			{
				colFill.Width = new GridLength(0.0, GridUnitType.Star);
				colEmpty.Width = new GridLength(100.0, GridUnitType.Star);
				if (barFill != null)
				{
					barFill.Visibility = Visibility.Collapsed;
				}
			}
			else
			{
				if (barFill != null)
				{
					barFill.Visibility = Visibility.Visible;
				}
				colFill.Width = new GridLength(percent, GridUnitType.Star);
				colEmpty.Width = new GridLength(Math.Max(0, 100 - percent), GridUnitType.Star);
			}
		}
		UpdateLayout();
		DoEvents();
	}

	public void EnableCancel(bool enabled)
	{
		if (btnCancel != null)
		{
			btnCancel.Visibility = ((!enabled) ? Visibility.Collapsed : Visibility.Visible);
		}
	}

	private void BtnCancel_Click(object sender, RoutedEventArgs e)
	{
		if (IsCancelRequested)
		{
			return;
		}
		IsCancelRequested = true;
		btnCancel.IsEnabled = false;
		btnCancel.Content = "Đang dừng...";
		lblHint.Text = "Sẽ dừng ngay sau bước đang chạy...";
		Action cancelRequested = CancelRequested;
		if (cancelRequested == null)
		{
			return;
		}
		try
		{
			cancelRequested();
		}
		catch
		{
		}
	}

	private void Window_MouseDown(object sender, MouseButtonEventArgs e)
	{
		if (e.LeftButton == MouseButtonState.Pressed)
		{
			try
			{
				DragMove();
			}
			catch
			{
			}
		}
	}

	public void CloseProgress()
	{
		if (!Dispatcher.CheckAccess())
		{
			Dispatcher.Invoke(CloseProgress);
			return;
		}
		try
		{
			Close();
		}
		catch
		{
		}
	}

	[SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
	public static void DoEvents()
	{
		DispatcherFrame dispatcherFrame = new DispatcherFrame();
		Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrame), dispatcherFrame);
		Dispatcher.PushFrame(dispatcherFrame);
	}

	private static object ExitFrame(object frame)
	{
		((DispatcherFrame)frame).Continue = false;
		return null;
	}
}
