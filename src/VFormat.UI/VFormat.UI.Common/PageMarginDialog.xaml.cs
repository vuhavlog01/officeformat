using System;
using System.Windows;
using System.Windows.Markup;
using VFormat.Core.Enums;

namespace VFormat.UI.Common;

public partial class PageMarginDialog : Window, IComponentConnector
{
	public PageMarginViewModel ViewModel { get; private set; }

	public bool IsApplied { get; private set; }

	public Action<float, float, float, float, float, float, bool, bool, bool> OnApplyMargins { get; set; }

	public PageMarginDialog()
	{
		InitializeComponent();
		ViewModel = new PageMarginViewModel();
		DataContext = ViewModel;
		IsApplied = false;
	}

	public PageMarginDialog(RulePackType rulePack)
		: this()
	{
		ViewModel.SelectedRulePack = rulePack;
	}

	public PageMarginDialog(RulePackType rulePack, PrintLayoutMode printLayoutMode)
		: this()
	{
		ViewModel.SelectedRulePack = rulePack;
		ViewModel.MirrorMargins = printLayoutMode == PrintLayoutMode.TwoSidedMirror;
	}

	public PageMarginDialog(PageMarginViewModel vm)
	{
		InitializeComponent();
		ViewModel = vm ?? new PageMarginViewModel();
		DataContext = ViewModel;
		IsApplied = false;
	}

	private void BtnApply_Click(object sender, RoutedEventArgs e)
	{
		IsApplied = true;
		if (OnApplyMargins != null)
		{
			OnApplyMargins(ViewModel.TopMargin, ViewModel.BottomMargin, ViewModel.LeftMargin, ViewModel.RightMargin, ViewModel.HeaderDistance, ViewModel.FooterDistance, ViewModel.IsLandscape, ViewModel.CurrentSectionOnly, ViewModel.MirrorMargins);
		}
		DialogResult = true;
		Close();
	}

	private void BtnResetQC_Click(object sender, RoutedEventArgs e)
	{
		ViewModel.ResetToStandard();
	}

	private void BtnClose_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}
}
