using System;
using System.Linq;
using System.Windows;
using System.Windows.Markup;
using VFormat.Core.Templates;

namespace VFormat.UI.Templates;

public partial class TemplateLibraryWindow : Window, IComponentConnector
{
	public TemplateLibraryViewModel ViewModel { get; private set; }

	public Action<TemplateItem> OpenTemplateCallback { get; set; }

	public TemplateLibraryWindow()
	{
		InitializeComponent();
		ViewModel = new TemplateLibraryViewModel();
		TemplateLibraryViewModel viewModel = ViewModel;
		Action<TemplateItem> openTemplateCallback = (TemplateItem item) =>
		{
			if (OpenTemplateCallback != null)
			{
				OpenTemplateCallback(item);
			}
		};
		viewModel.OpenTemplateCallback = openTemplateCallback;
		ViewModel.RequestCloseCallback = () =>
		{
			Close();
		};
		libraryControl.ViewModel = ViewModel;
	}

	public TemplateLibraryWindow(string initialCategoryName)
		: this()
	{
		if (!string.IsNullOrEmpty(initialCategoryName))
		{
			TemplateCategoryItem templateCategoryItem = ViewModel.Categories.FirstOrDefault((TemplateCategoryItem c) => string.Equals(c.Name, initialCategoryName, StringComparison.OrdinalIgnoreCase));
			if (templateCategoryItem != null)
			{
				ViewModel.SelectCategory(templateCategoryItem);
			}
		}
	}
}
