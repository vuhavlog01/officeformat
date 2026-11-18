using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace VFormat.UI.License;

public class LicenseRequiredDialog : Window
{
	public const string DialogTitle = "Kích hoạt bản quyền OfficeFormat";

	public const string DialogMessage = "Bạn cần kích hoạt bản quyền OfficeFormat để sử dụng tính năng này.\n\nNếu đã mua, hãy chọn ‘Nhập mã bản quyền’. Nếu chưa có mã, bạn có thể chọn ‘Mua bản quyền’ để xem các gói sử dụng.";

	public LicenseRequiredChoice Choice { get; private set; }

	public LicenseRequiredDialog()
	{
		Choice = LicenseRequiredChoice.Close;
		Title = "Kích hoạt bản quyền OfficeFormat";
		Width = 500.0;
		MinWidth = 360.0;
		SizeToContent = SizeToContent.Height;
		ResizeMode = ResizeMode.NoResize;
		WindowStartupLocation = WindowStartupLocation.CenterScreen;
		ShowInTaskbar = false;
		FontFamily = new FontFamily("Segoe UI, Tahoma, Arial");
		Background = Brushes.White;
		StackPanel stackPanel = new StackPanel
		{
			Margin = new Thickness(22.0, 20.0, 22.0, 18.0)
		};
		stackPanel.Children.Add(new TextBlock
		{
			Text = "Kích hoạt bản quyền OfficeFormat",
			FontSize = 16.0,
			FontWeight = FontWeights.Bold,
			Foreground = new SolidColorBrush(Color.FromRgb(15, 23, 42)),
			Margin = new Thickness(0.0, 0.0, 0.0, 10.0),
			TextWrapping = TextWrapping.Wrap
		});
		stackPanel.Children.Add(new TextBlock
		{
			Text = "Bạn cần kích hoạt bản quyền OfficeFormat để sử dụng tính năng này.\n\nNếu đã mua, hãy chọn ‘Nhập mã bản quyền’. Nếu chưa có mã, bạn có thể chọn ‘Mua bản quyền’ để xem các gói sử dụng.",
			FontSize = 13.0,
			TextWrapping = TextWrapping.Wrap,
			LineHeight = 20.0
		});
		WrapPanel wrapPanel = new WrapPanel
		{
			HorizontalAlignment = HorizontalAlignment.Right,
			Margin = new Thickness(0.0, 18.0, 0.0, 0.0)
		};
		wrapPanel.Children.Add(MakeButton("Nhập mã bản quyền", LicenseRequiredChoice.EnterKey, isDefault: true));
		wrapPanel.Children.Add(MakeButton("Mua bản quyền", LicenseRequiredChoice.Buy, isDefault: false));
		Button button = MakeButton("Đóng", LicenseRequiredChoice.Close, isDefault: false);
		button.IsCancel = true;
		wrapPanel.Children.Add(button);
		stackPanel.Children.Add(wrapPanel);
		Content = stackPanel;
	}

	private Button MakeButton(string text, LicenseRequiredChoice choice, bool isDefault)
	{
		Button button = new Button();
		button.Content = text;
		button.Padding = new Thickness(14.0, 7.0, 14.0, 7.0);
		button.Margin = new Thickness(8.0, 6.0, 0.0, 0.0);
		button.MinWidth = 90.0;
		button.IsDefault = isDefault;
		Button button2 = button;
		button2.Click += (object s, RoutedEventArgs e) =>
		{
			Choice = choice;
			Close();
		};
		return button2;
	}

	public static LicenseRequiredChoice Ask(IntPtr ownerHwnd)
	{
		LicenseRequiredDialog licenseRequiredDialog = new LicenseRequiredDialog();
		if (ownerHwnd != IntPtr.Zero)
		{
			try
			{
				new WindowInteropHelper(licenseRequiredDialog).Owner = ownerHwnd;
			}
			catch
			{
			}
		}
		licenseRequiredDialog.ShowDialog();
		return licenseRequiredDialog.Choice;
	}
}
