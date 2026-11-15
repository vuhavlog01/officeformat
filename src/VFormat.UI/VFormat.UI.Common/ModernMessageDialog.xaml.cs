using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;

namespace VFormat.UI.Common;

public partial class ModernMessageDialog : Window, IComponentConnector
{
	public bool Confirmed { get; private set; }

	public ModernMessageDialog()
	{
		InitializeComponent();
	}

	public ModernMessageDialog(string message, string title, MessageDialogType type = MessageDialogType.Information, string okText = "Đóng", string cancelText = "Hủy Bỏ")
	{
		InitializeComponent();
		txtMessage.Text = message ?? "";
		txtTitle.Text = (string.IsNullOrEmpty(title) ? "OfficeFormat 1.0.0" : title);
		Title = txtTitle.Text;
		btnOk.Content = okText;
		btnCancel.Content = cancelText;
		ConfigureType(type);
	}

	private void ConfigureType(MessageDialogType type)
	{
		switch (type)
		{
		case MessageDialogType.Success:
			txtCategory.Text = "Thao tác thành công";
			borderIconBox.Background = new SolidColorBrush(Color.FromRgb(236, 253, 245));
			pathIcon.Data = (Geometry)FindResource("IconCheck");
			pathIcon.Fill = new SolidColorBrush(Color.FromRgb(16, 185, 129));
			break;
		case MessageDialogType.Warning:
			txtCategory.Text = "Cảnh báo & Lưu ý";
			borderIconBox.Background = new SolidColorBrush(Color.FromRgb(byte.MaxValue, 251, 235));
			pathIcon.Data = (Geometry)FindResource("IconWarning");
			pathIcon.Fill = new SolidColorBrush(Color.FromRgb(245, 158, 11));
			break;
		case MessageDialogType.Error:
			txtCategory.Text = "Lỗi hệ thống";
			borderIconBox.Background = new SolidColorBrush(Color.FromRgb(254, 242, 242));
			pathIcon.Data = (Geometry)FindResource("IconError");
			pathIcon.Fill = new SolidColorBrush(Color.FromRgb(239, 68, 68));
			break;
		case MessageDialogType.Question:
			txtCategory.Text = "Xác nhận hành động";
			borderIconBox.Background = new SolidColorBrush(Color.FromRgb(byte.MaxValue, 241, 230));
			pathIcon.Data = (Geometry)FindResource("IconQuestion");
			pathIcon.Fill = new SolidColorBrush(Color.FromRgb(224, 79, 38));
			btnCancel.Visibility = Visibility.Visible;
			if (btnOk.Content.ToString() == "Đóng")
			{
				btnOk.Content = "Đồng ý";
			}
			break;
		default:
			txtCategory.Text = "Thông tin thể thức";
			borderIconBox.Background = new SolidColorBrush(Color.FromRgb(byte.MaxValue, 241, 230));
			pathIcon.Data = (Geometry)FindResource("IconInfo");
			pathIcon.Fill = new SolidColorBrush(Color.FromRgb(224, 79, 38));
			break;
		}
	}

	private void BtnOk_Click(object sender, RoutedEventArgs e)
	{
		Confirmed = true;
		DialogResult = true;
		Close();
	}

	private void BtnCancel_Click(object sender, RoutedEventArgs e)
	{
		Confirmed = false;
		DialogResult = false;
		Close();
	}

	public static void ShowInfo(string message, string title = "OfficeFormat 1.0.0", IntPtr? ownerHwnd = null)
	{
		ShowDialogInternal(message, title, MessageDialogType.Information, "Đóng", null, ownerHwnd);
	}

	public static void ShowSuccess(string message, string title = "OfficeFormat 1.0.0", IntPtr? ownerHwnd = null)
	{
		ShowDialogInternal(message, title, MessageDialogType.Success, "Đóng", null, ownerHwnd);
	}

	public static void ShowWarning(string message, string title = "OfficeFormat 1.0.0", IntPtr? ownerHwnd = null)
	{
		ShowDialogInternal(message, title, MessageDialogType.Warning, "Đóng", null, ownerHwnd);
	}

	public static void ShowError(string message, string title = "OfficeFormat 1.0.0", IntPtr? ownerHwnd = null)
	{
		ShowDialogInternal(message, title, MessageDialogType.Error, "Đóng", null, ownerHwnd);
	}

	public static bool ShowConfirm(string message, string title = "OfficeFormat 1.0.0", string okText = "Đồng ý", string cancelText = "Hủy Bỏ", IntPtr? ownerHwnd = null)
	{
		ModernMessageDialog modernMessageDialog = new ModernMessageDialog(message, title, MessageDialogType.Question, okText, cancelText);
		SetOwnerIfProvided(modernMessageDialog, ownerHwnd);
		modernMessageDialog.ShowDialog();
		return modernMessageDialog.Confirmed;
	}

	private static void ShowDialogInternal(string message, string title, MessageDialogType type, string okText, string cancelText, IntPtr? ownerHwnd)
	{
		ModernMessageDialog modernMessageDialog = new ModernMessageDialog(message, title, type, okText, cancelText);
		SetOwnerIfProvided(modernMessageDialog, ownerHwnd);
		modernMessageDialog.ShowDialog();
	}

	private static void SetOwnerIfProvided(Window win, IntPtr? ownerHwnd)
	{
		if (ownerHwnd.HasValue && ownerHwnd.Value != IntPtr.Zero)
		{
			try
			{
				new WindowInteropHelper(win).Owner = ownerHwnd.Value;
			}
			catch
			{
			}
		}
	}
}
