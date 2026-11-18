using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using VFormat.Core.Templates;

namespace VFormat.UI.Templates;

public partial class TemplateLibraryControl : UserControl, IComponentConnector, IStyleConnector
{
	public TemplateLibraryViewModel ViewModel
	{
		get
		{
			return DataContext as TemplateLibraryViewModel;
		}
		set
		{
			DataContext = value;
		}
	}

	public TemplateLibraryControl()
	{
		InitializeComponent();
	}

	public TemplateLibraryControl(TemplateLibraryViewModel viewModel)
		: this()
	{
		ViewModel = viewModel;
	}

	private void CategoryCard_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
	{
		if (sender is FrameworkElement { DataContext: TemplateCategoryItem dataContext } && ViewModel != null)
		{
			ViewModel.SelectCategory(dataContext);
		}
	}
}
