#define TRACE
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace VFormat.UI.Common;

public static class ExternalLinkHelper
{
	public const string BuyLicenseFailMessage = "Chưa cấu hình tên miền (Domain) cho liên kết này.";

	public const string BuyTransferFailMessage = "Chưa cấu hình tên miền (Domain) cho liên kết này.";

	internal static Action<string> TestLauncher;

	public static string BuyLicenseUrl => "";

	public static bool IsAllowedUrl(string url)
	{
		if (string.IsNullOrWhiteSpace(url))
		{
			return false;
		}
		if (Uri.TryCreate(url, UriKind.Absolute, out var result) && result.Scheme == Uri.UriSchemeHttps && result.IsDefaultPort)
		{
			return string.IsNullOrEmpty(result.UserInfo);
		}
		return false;
	}

	public static bool OpenBuyLicensePage(Window owner = null, IntPtr ownerHwnd = default(IntPtr))
	{
		return TryOpen(BuyLicenseUrl, BuyLicenseFailMessage, owner, ownerHwnd);
	}

	public static bool TryOpen(string url, string failMessage, Window owner = null, IntPtr ownerHwnd = default(IntPtr))
	{
		if (!IsAllowedUrl(url))
		{
			return false;
		}
		try
		{
			if (TestLauncher != null)
			{
				TestLauncher(url);
			}
			else
			{
				ProcessStartInfo processStartInfo = new ProcessStartInfo(url);
				processStartInfo.UseShellExecute = true;
				ProcessStartInfo startInfo = processStartInfo;
				using (Process.Start(startInfo))
				{
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			Trace.WriteLine("Open link failed: " + ex.GetType().Name);
		}
		try
		{
			ShowCopyLinkDialog(url, failMessage, owner, ownerHwnd);
		}
		catch (Exception ex2)
		{
			Trace.WriteLine("Copy link dialog failed: " + ex2.GetType().Name);
		}
		return false;
	}

	public static void ShowCopyLinkDialog(string url, string message, Window owner = null, IntPtr ownerHwnd = default(IntPtr))
	{
		Window win = new Window
		{
			Title = "OfficeFormat 1.0.0",
			SizeToContent = SizeToContent.Height,
			Width = 520.0,
			MinWidth = 380.0,
			ResizeMode = ResizeMode.NoResize,
			WindowStartupLocation = ((owner == null) ? WindowStartupLocation.CenterScreen : WindowStartupLocation.CenterOwner),
			ShowInTaskbar = false,
			FontFamily = new FontFamily("Segoe UI, Tahoma, Arial")
		};
		if (owner != null)
		{
			win.Owner = owner;
		}
		else if (ownerHwnd != IntPtr.Zero)
		{
			try
			{
				new WindowInteropHelper(win).Owner = ownerHwnd;
			}
			catch
			{
			}
		}
		StackPanel stackPanel = new StackPanel();
		stackPanel.Margin = new Thickness(20.0);
		StackPanel stackPanel2 = stackPanel;
		stackPanel2.Children.Add(new TextBlock
		{
			Text = (message ?? ""),
			TextWrapping = TextWrapping.Wrap,
			FontSize = 13.0,
			Margin = new Thickness(0.0, 0.0, 0.0, 12.0)
		});
		TextBox box = new TextBox
		{
			Text = url,
			IsReadOnly = true,
			FontFamily = new FontFamily("Consolas, Courier New"),
			Padding = new Thickness(6.0),
			TextWrapping = TextWrapping.Wrap
		};
		stackPanel2.Children.Add(box);
		TextBlock status = new TextBlock
		{
			TextWrapping = TextWrapping.Wrap,
			FontSize = 12.0,
			Margin = new Thickness(0.0, 8.0, 0.0, 0.0),
			Foreground = Brushes.SeaGreen
		};
		stackPanel2.Children.Add(status);
		WrapPanel wrapPanel = new WrapPanel();
		wrapPanel.HorizontalAlignment = HorizontalAlignment.Right;
		wrapPanel.Margin = new Thickness(0.0, 14.0, 0.0, 0.0);
		WrapPanel wrapPanel2 = wrapPanel;
		Button button = new Button();
		button.Content = "Sao chép liên kết";
		button.Padding = new Thickness(14.0, 6.0, 14.0, 6.0);
		button.Margin = new Thickness(0.0, 0.0, 8.0, 0.0);
		button.IsDefault = true;
		Button button2 = button;
		Button button3 = new Button();
		button3.Content = "Đóng";
		button3.Padding = new Thickness(18.0, 6.0, 18.0, 6.0);
		button3.IsCancel = true;
		Button button4 = button3;
		button2.Click += (object s, RoutedEventArgs e) =>
		{
			try
			{
				Clipboard.SetText(url);
				status.Foreground = Brushes.SeaGreen;
				status.Text = "Đã sao chép liên kết. Hãy dán vào thanh địa chỉ của trình duyệt.";
			}
			catch
			{
				box.Focus();
				box.SelectAll();
				status.Foreground = Brushes.DarkOrange;
				status.Text = "Chưa sao chép được tự động. Liên kết đã được chọn, hãy nhấn Ctrl+C.";
			}
		};
		button4.Click += (object s, RoutedEventArgs e) =>
		{
			win.Close();
		};
		wrapPanel2.Children.Add(button2);
		wrapPanel2.Children.Add(button4);
		stackPanel2.Children.Add(wrapPanel2);
		win.Content = stackPanel2;
		win.ShowDialog();
	}
}
