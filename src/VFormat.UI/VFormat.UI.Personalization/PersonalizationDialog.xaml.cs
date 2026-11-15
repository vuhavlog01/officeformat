using System;
using System.Windows;
using System.Windows.Markup;

namespace VFormat.UI.Personalization;

public partial class PersonalizationDialog : Window, IComponentConnector
{
	public PersonalizationViewModel ViewModel => DataContext as PersonalizationViewModel;

	public PersonalizationDialog()
	{
		InitializeComponent();
		PersonalizationViewModel personalizationViewModel = new PersonalizationViewModel();
		Action closeAction = () =>
		{
			Close();
		};
		personalizationViewModel.CloseAction = closeAction;
		DataContext = personalizationViewModel;
	}

	public PersonalizationDialog(PersonalizationViewModel vm)
		: this()
	{
		if (vm != null)
		{
			vm.CloseAction = () =>
			{
				Close();
			};
			DataContext = vm;
		}
	}

	private void IndentTextBox_LostFocus(object sender, RoutedEventArgs e)
	{
		if (ViewModel != null)
		{
			ViewModel.NormalizeIndentText();
		}
	}

	private void SpacingBeforeTextBox_LostFocus(object sender, RoutedEventArgs e)
	{
		if (ViewModel != null)
		{
			ViewModel.NormalizeSpacingBeforeText();
		}
	}

	private void SpacingAfterTextBox_LostFocus(object sender, RoutedEventArgs e)
	{
		if (ViewModel != null)
		{
			ViewModel.NormalizeSpacingAfterText();
		}
	}

	private void LineSpacingAtTextBox_LostFocus(object sender, RoutedEventArgs e)
	{
		if (ViewModel != null)
		{
			ViewModel.NormalizeLineSpacingAtText();
		}
	}
}
