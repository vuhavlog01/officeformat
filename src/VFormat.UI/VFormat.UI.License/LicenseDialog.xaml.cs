using System.Windows;
using System.Windows.Markup;

namespace VFormat.UI.License;

public partial class LicenseDialog : Window, IComponentConnector
{
	public LicenseDialog()
		: this(new LicenseViewModel(), loadTransferStatus: true)
	{
	}

	public LicenseDialog(LicenseViewModel vm)
		: this(vm ?? new LicenseViewModel(), loadTransferStatus: true)
	{
	}

	internal LicenseDialog(LicenseViewModel vm, bool loadTransferStatus)
	{
		LicenseDialog licenseDialog = this;
		InitializeComponent();
		MaxHeight = SystemParameters.WorkArea.Height;
		MaxWidth = SystemParameters.WorkArea.Width;
		LicenseTabs.SelectedIndex = (vm.IsSuccess ? 2 : 0);
		vm.CloseAction = () =>
		{
			Close();
		};
		vm.OwnerWindow = this;
		DataContext = vm;
		if (loadTransferStatus)
		{
			Loaded += (object s, RoutedEventArgs e) =>
			{
				vm.LoadTransferStatus();
			};
		}
	}
}
