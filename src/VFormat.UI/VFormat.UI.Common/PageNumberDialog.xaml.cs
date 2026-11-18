using System;
using System.Windows;
using System.Windows.Markup;

namespace VFormat.UI.Common;

public partial class PageNumberDialog : Window, IComponentConnector
{
	public PageNumberViewModel ViewModel { get; private set; }

	public bool IsApplied { get; private set; }

	public bool IsRemoveAll { get; private set; }

	public Action<int, int, bool, bool> OnApplyPageNumbering { get; set; }

	public Action OnRemovePageNumbers { get; set; }

	public PageNumberDialog()
	{
		InitializeComponent();
		ViewModel = new PageNumberViewModel();
		DataContext = ViewModel;
		IsApplied = false;
		IsRemoveAll = false;
	}

	public PageNumberDialog(PageNumberViewModel vm)
	{
		InitializeComponent();
		ViewModel = vm ?? new PageNumberViewModel();
		DataContext = ViewModel;
		IsApplied = false;
		IsRemoveAll = false;
	}

	private void BtnApply_Click(object sender, RoutedEventArgs e)
	{
		IsApplied = true;
		IsRemoveAll = false;
		if (OnApplyPageNumbering != null)
		{
			int arg = (ViewModel.FromStart ? 1 : ViewModel.StartPage);
			int arg2 = (ViewModel.FromStart ? 1 : ViewModel.StartNumber);
			bool arg3 = ViewModel.FromStart && ViewModel.HideFirstPage;
			bool isHeader = ViewModel.IsHeader;
			OnApplyPageNumbering(arg, arg2, arg3, isHeader);
		}
		DialogResult = true;
		Close();
	}

	private void BtnRemove_Click(object sender, RoutedEventArgs e)
	{
		IsApplied = false;
		IsRemoveAll = true;
		if (OnRemovePageNumbers != null)
		{
			OnRemovePageNumbers();
		}
		DialogResult = true;
		Close();
	}

	private void BtnClose_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}
}
