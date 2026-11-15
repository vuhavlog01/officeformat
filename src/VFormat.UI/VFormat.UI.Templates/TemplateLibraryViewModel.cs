using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using VFormat.Core.Templates;
using VFormat.UI.Common;

namespace VFormat.UI.Templates;

public class TemplateLibraryViewModel : ViewModelBase
{
	private string _searchText = "";

	private TemplateCategoryItem _selectedCategory;

	private ObservableCollection<TemplateCategoryItem> _categories;

	private ObservableCollection<TemplateGroupSection> _groupSections;

	private string _headerTitle = "THƯ VIỆN BIỂU MẪU";

	private string _headerSubtitle = "Kho biểu mẫu chuẩn Quốc gia & Cá nhân";

	private string _searchWatermark = "\ud83d\udd0d Tìm kiếm biểu mẫu, số mẫu, từ khóa...";

	private bool _isCategoryView = true;

	private bool _isDetailView;

	private bool _isSearching;

	private bool _isEmpty;

	private bool _canGoBack;

	public Action<TemplateItem> OpenTemplateCallback { get; set; }

	public Action RequestCloseCallback { get; set; }

	public ICommand SelectCategoryCommand { get; private set; }

	public ICommand GoBackCommand { get; private set; }

	public ICommand ClearSearchCommand { get; private set; }

	public ICommand UseTemplateCommand { get; private set; }

	public ICommand RefreshCommand { get; private set; }

	public string SearchText
	{
		get
		{
			return _searchText;
		}
		set
		{
			if (SetProperty(ref _searchText, value, "SearchText"))
			{
				OnSearchTextChanged();
			}
		}
	}

	public TemplateCategoryItem SelectedCategory
	{
		get
		{
			return _selectedCategory;
		}
		set
		{
			SetProperty(ref _selectedCategory, value, "SelectedCategory");
		}
	}

	public ObservableCollection<TemplateCategoryItem> Categories
	{
		get
		{
			return _categories;
		}
		set
		{
			SetProperty(ref _categories, value, "Categories");
		}
	}

	public ObservableCollection<TemplateGroupSection> GroupSections
	{
		get
		{
			return _groupSections;
		}
		set
		{
			SetProperty(ref _groupSections, value, "GroupSections");
		}
	}

	public string HeaderTitle
	{
		get
		{
			return _headerTitle;
		}
		set
		{
			SetProperty(ref _headerTitle, value, "HeaderTitle");
		}
	}

	public string HeaderSubtitle
	{
		get
		{
			return _headerSubtitle;
		}
		set
		{
			SetProperty(ref _headerSubtitle, value, "HeaderSubtitle");
		}
	}

	public string SearchWatermark
	{
		get
		{
			return _searchWatermark;
		}
		set
		{
			SetProperty(ref _searchWatermark, value, "SearchWatermark");
		}
	}

	public bool IsCategoryView
	{
		get
		{
			return _isCategoryView;
		}
		set
		{
			SetProperty(ref _isCategoryView, value, "IsCategoryView");
		}
	}

	public bool IsDetailView
	{
		get
		{
			return _isDetailView;
		}
		set
		{
			SetProperty(ref _isDetailView, value, "IsDetailView");
		}
	}

	public bool IsSearching
	{
		get
		{
			return _isSearching;
		}
		set
		{
			SetProperty(ref _isSearching, value, "IsSearching");
		}
	}

	public bool IsEmpty
	{
		get
		{
			return _isEmpty;
		}
		set
		{
			SetProperty(ref _isEmpty, value, "IsEmpty");
		}
	}

	public bool CanGoBack
	{
		get
		{
			return _canGoBack;
		}
		set
		{
			SetProperty(ref _canGoBack, value, "CanGoBack");
		}
	}

	public TemplateLibraryViewModel()
	{
		_categories = new ObservableCollection<TemplateCategoryItem>();
		_groupSections = new ObservableCollection<TemplateGroupSection>();
		Action<object> execute = (object param) =>
		{
			if (param is TemplateCategoryItem category)
			{
				SelectCategory(category);
			}
		};
		SelectCategoryCommand = new RelayCommand(execute);
		GoBackCommand = new RelayCommand(() =>
		{
			GoBack();
		});
		ClearSearchCommand = new RelayCommand(() =>
		{
			ClearSearch();
		});
		UseTemplateCommand = new RelayCommand((object param) =>
		{
			if (param is TemplateItem item)
			{
				UseTemplate(item);
			}
		});
		RefreshCommand = new RelayCommand(() =>
		{
			TemplateLibraryService.Instance.Reload();
			LoadCategories();
			if (_selectedCategory != null)
			{
				SelectCategory(_selectedCategory);
			}
		});
		LoadCategories();
	}

	public void LoadCategories()
	{
		_categories.Clear();
		IReadOnlyList<TemplateCategoryItem> categories = TemplateLibraryService.Instance.GetCategories();
		foreach (TemplateCategoryItem item in categories)
		{
			_categories.Add(item);
		}
	}

	public void SelectCategory(TemplateCategoryItem category)
	{
		if (category != null)
		{
			_selectedCategory = category;
			IsCategoryView = false;
			IsDetailView = true;
			IsSearching = false;
			CanGoBack = true;
			HeaderTitle = category.Name;
			HeaderSubtitle = $"{category.BadgeText} ({category.Description})";
			SearchWatermark = $"\ud83d\udd0d Tìm trong {category.Name}...";
			RefreshGroupSections();
		}
	}

	public void GoBack()
	{
		if (IsSearching && _selectedCategory != null)
		{
			_searchText = "";
			OnPropertyChanged("SearchText");
			IsSearching = false;
			RefreshGroupSections();
			return;
		}
		_selectedCategory = null;
		_searchText = "";
		OnPropertyChanged("SearchText");
		IsCategoryView = true;
		IsDetailView = false;
		IsSearching = false;
		IsEmpty = false;
		CanGoBack = false;
		HeaderTitle = "THƯ VIỆN BIỂU MẪU";
		HeaderSubtitle = "Kho biểu mẫu chuẩn Quốc gia & Cá nhân";
		SearchWatermark = "\ud83d\udd0d Tìm kiếm biểu mẫu, số mẫu, từ khóa...";
		_groupSections.Clear();
	}

	public void ClearSearch()
	{
		SearchText = "";
	}

	private void OnSearchTextChanged()
	{
		if (string.IsNullOrWhiteSpace(_searchText))
		{
			if (_selectedCategory != null)
			{
				IsCategoryView = false;
				IsDetailView = true;
				IsSearching = false;
				HeaderTitle = _selectedCategory.Name;
				HeaderSubtitle = $"{_selectedCategory.BadgeText} ({_selectedCategory.Description})";
				RefreshGroupSections();
			}
			else
			{
				IsCategoryView = true;
				IsDetailView = false;
				IsSearching = false;
				IsEmpty = false;
				CanGoBack = false;
				HeaderTitle = "THƯ VIỆN BIỂU MẪU";
				HeaderSubtitle = "Kho biểu mẫu chuẩn Quốc gia & Cá nhân";
				_groupSections.Clear();
			}
		}
		else
		{
			IsCategoryView = false;
			IsDetailView = false;
			IsSearching = true;
			CanGoBack = true;
			if (_selectedCategory != null)
			{
				HeaderTitle = $"Tìm trong: {_selectedCategory.Name}";
				IReadOnlyList<TemplateGroupSection> groupSectionsByCategory = TemplateLibraryService.Instance.GetGroupSectionsByCategory(_selectedCategory.Name, _searchText);
				UpdateGroupSections(groupSectionsByCategory);
			}
			else
			{
				HeaderTitle = "KẾT QUẢ TÌM KIẾM";
				IReadOnlyList<TemplateGroupSection> sections = TemplateLibraryService.Instance.SearchAll(_searchText);
				UpdateGroupSections(sections);
			}
			int num = _groupSections.Sum((TemplateGroupSection g) => g.Count);
			HeaderSubtitle = $"Tìm thấy {num} biểu mẫu phù hợp";
			IsEmpty = num == 0;
		}
	}

	private void RefreshGroupSections()
	{
		if (_selectedCategory != null)
		{
			IReadOnlyList<TemplateGroupSection> groupSectionsByCategory = TemplateLibraryService.Instance.GetGroupSectionsByCategory(_selectedCategory.Name, null);
			UpdateGroupSections(groupSectionsByCategory);
			IsEmpty = _groupSections.Count == 0;
		}
	}

	private void UpdateGroupSections(IReadOnlyList<TemplateGroupSection> sections)
	{
		_groupSections.Clear();
		if (sections == null)
		{
			return;
		}
		foreach (TemplateGroupSection section in sections)
		{
			_groupSections.Add(section);
		}
	}

	private void UseTemplate(TemplateItem item)
	{
		if (item != null)
		{
			if (OpenTemplateCallback != null)
			{
				OpenTemplateCallback(item);
			}
			if (RequestCloseCallback != null)
			{
				RequestCloseCallback();
			}
		}
	}
}
