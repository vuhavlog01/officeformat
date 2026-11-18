using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using Microsoft.Win32;
using VFormat.Core.Templates;

namespace VFormat.UI.Templates;

public partial class MyTemplatesManagerWindow : Window, IComponentConnector
{
	private List<UserTemplateItem> _allTemplates = new List<UserTemplateItem>();

	public Action<UserTemplateItem> OpenTemplateCallback { get; set; }

	public Action TemplatesChangedCallback { get; set; }

	public MyTemplatesManagerWindow()
	{
		InitializeComponent();
		LoadTemplates();
	}

	public void LoadTemplates()
	{
		try
		{
			_allTemplates = UserTemplateService.Instance.GetAllTemplates();
			ApplyFilter();
		}
		catch (Exception ex)
		{
			MessageBox.Show(this, "Không thể tải danh sách biểu mẫu: " + ex.Message, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Hand);
		}
	}

	private void ApplyFilter()
	{
		string keyword = ((txtSearch != null) ? (txtSearch.Text ?? "").Trim().ToLowerInvariant() : "");
		List<UserTemplateItem> list = ((!string.IsNullOrEmpty(keyword)) ? _allTemplates.Where((UserTemplateItem x) => (x.Name != null && x.Name.ToLowerInvariant().Contains(keyword)) || (x.FileName != null && x.FileName.ToLowerInvariant().Contains(keyword)) || (x.Description != null && x.Description.ToLowerInvariant().Contains(keyword))).ToList() : _allTemplates);
		lvTemplates.ItemsSource = list;
		int count = list.Count;
		lblItemCount.Text = $"Tổng số: {count} biểu mẫu";
		if (_allTemplates.Count == 0)
		{
			pnlEmptyState.Visibility = Visibility.Visible;
			lvTemplates.Visibility = Visibility.Collapsed;
		}
		else
		{
			pnlEmptyState.Visibility = Visibility.Collapsed;
			lvTemplates.Visibility = Visibility.Visible;
		}
		UpdateButtonsState();
	}

	private void UpdateButtonsState()
	{
		bool isEnabled = lvTemplates.SelectedItem is UserTemplateItem;
		btnUseTemplate.IsEnabled = isEnabled;
		btnReplace.IsEnabled = isEnabled;
		btnRename.IsEnabled = isEnabled;
		btnDelete.IsEnabled = isEnabled;
	}

	private void LvTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		UpdateButtonsState();
	}

	private void LvTemplates_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		ExecuteUseTemplate();
	}

	private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
	{
		if (lblSearchWatermark != null)
		{
			lblSearchWatermark.Visibility = ((!string.IsNullOrEmpty(txtSearch.Text)) ? Visibility.Collapsed : Visibility.Visible);
		}
		ApplyFilter();
	}

	private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			string templateStorageDirectory = UserTemplateService.Instance.TemplateStorageDirectory;
			if (!Directory.Exists(templateStorageDirectory))
			{
				Directory.CreateDirectory(templateStorageDirectory);
			}
			Process.Start("explorer.exe", templateStorageDirectory);
		}
		catch (Exception ex)
		{
			MessageBox.Show(this, "Không thể mở thư mục lưu trữ: " + ex.Message, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Exclamation);
		}
	}

	private void BtnAdd_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			AddUserTemplateDialog dlg = new AddUserTemplateDialog();
			dlg.Owner = this;
			if (dlg.ShowDialog() != true)
			{
				return;
			}
			LoadTemplates();
			if (dlg.CreatedTemplateItem != null)
			{
				lvTemplates.SelectedItem = _allTemplates.FirstOrDefault((UserTemplateItem x) => x.Id == dlg.CreatedTemplateItem.Id);
			}
			if (TemplatesChangedCallback != null)
			{
				TemplatesChangedCallback();
			}
		}
		catch (Exception ex)
		{
			MessageBox.Show(this, "Lỗi khi mở hộp thoại thêm mẫu: " + ex.Message, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Hand);
		}
	}

	private void BtnReplace_Click(object sender, RoutedEventArgs e)
	{
		UserTemplateItem selected = lvTemplates.SelectedItem as UserTemplateItem;
		if (selected == null)
		{
			return;
		}
		OpenFileDialog openFileDialog = new OpenFileDialog();
		openFileDialog.Title = $"Chọn tệp mới để thay thế cho biểu mẫu '{selected.Name}'";
		openFileDialog.Filter = "Tài liệu Word (*.docx;*.doc;*.dotx)|*.docx;*.doc;*.dotx|Tất cả các tệp (*.*)|*.*";
		openFileDialog.CheckFileExists = true;
		OpenFileDialog openFileDialog2 = openFileDialog;
		if (openFileDialog2.ShowDialog(this) != true)
		{
			return;
		}
		string fileName = openFileDialog2.FileName;
		string messageBoxText = $"Bạn có chắc chắn muốn thay thế nội dung của biểu mẫu '{selected.Name}' bằng tệp:\n'{Path.GetFileName(fileName)}'?";
		if (MessageBox.Show(this, messageBoxText, "Xác nhận thay thế biểu mẫu", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
		{
			return;
		}
		if (UserTemplateService.Instance.ReplaceTemplateFile(selected.Id, fileName, out var errorMessage))
		{
			LoadTemplates();
			lvTemplates.SelectedItem = _allTemplates.FirstOrDefault((UserTemplateItem x) => x.Id == selected.Id);
			if (TemplatesChangedCallback != null)
			{
				TemplatesChangedCallback();
			}
			MessageBox.Show(this, "Đã cập nhật nội dung mới cho biểu mẫu thành công!", "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Asterisk);
		}
		else
		{
			MessageBox.Show(this, "Không thể thay thế biểu mẫu: " + errorMessage, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Hand);
		}
	}

	private void BtnRename_Click(object sender, RoutedEventArgs e)
	{
		UserTemplateItem selected = lvTemplates.SelectedItem as UserTemplateItem;
		if (selected == null)
		{
			return;
		}
		string text = PromptRename(selected.Name);
		if (string.IsNullOrEmpty(text) || string.Equals(text, selected.Name, StringComparison.Ordinal))
		{
			return;
		}
		if (UserTemplateService.Instance.RenameTemplate(selected.Id, text, out var errorMessage))
		{
			LoadTemplates();
			lvTemplates.SelectedItem = _allTemplates.FirstOrDefault((UserTemplateItem x) => x.Id == selected.Id);
			if (TemplatesChangedCallback != null)
			{
				TemplatesChangedCallback();
			}
		}
		else
		{
			MessageBox.Show(this, "Không thể đổi tên biểu mẫu: " + errorMessage, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Hand);
		}
	}

	private void BtnDelete_Click(object sender, RoutedEventArgs e)
	{
		if (!(lvTemplates.SelectedItem is UserTemplateItem userTemplateItem))
		{
			return;
		}
		string messageBoxText = $"Bạn có chắc chắn muốn xóa biểu mẫu '{userTemplateItem.Name}' khỏi thư viện cá nhân?\n(Tệp mẫu tương ứng trên đĩa cũng sẽ được xóa bỏ).";
		if (MessageBox.Show(this, messageBoxText, "Xác nhận xóa biểu mẫu", MessageBoxButton.YesNo, MessageBoxImage.Exclamation) != MessageBoxResult.Yes)
		{
			return;
		}
		if (UserTemplateService.Instance.DeleteTemplate(userTemplateItem.Id, out var errorMessage))
		{
			LoadTemplates();
			if (TemplatesChangedCallback != null)
			{
				TemplatesChangedCallback();
			}
		}
		else
		{
			MessageBox.Show(this, "Không thể xóa biểu mẫu: " + errorMessage, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Hand);
		}
	}

	private void BtnUseTemplate_Click(object sender, RoutedEventArgs e)
	{
		ExecuteUseTemplate();
	}

	private void ExecuteUseTemplate()
	{
		if (lvTemplates.SelectedItem is UserTemplateItem obj && OpenTemplateCallback != null)
		{
			OpenTemplateCallback(obj);
			Close();
		}
	}

	private void BtnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private string PromptRename(string currentName)
	{
		Window prompt = new Window
		{
			Title = "Đổi tên biểu mẫu",
			Width = 420.0,
			Height = 180.0,
			WindowStartupLocation = WindowStartupLocation.CenterOwner,
			Owner = this,
			ResizeMode = ResizeMode.NoResize,
			Background = Brushes.White,
			FontFamily = FontFamily
		};
		Grid grid = new Grid();
		grid.Margin = new Thickness(16.0);
		Grid grid2 = grid;
		grid2.RowDefinitions.Add(new RowDefinition
		{
			Height = GridLength.Auto
		});
		grid2.RowDefinitions.Add(new RowDefinition
		{
			Height = GridLength.Auto
		});
		grid2.RowDefinitions.Add(new RowDefinition
		{
			Height = new GridLength(1.0, GridUnitType.Star)
		});
		grid2.RowDefinitions.Add(new RowDefinition
		{
			Height = GridLength.Auto
		});
		TextBlock textBlock = new TextBlock();
		textBlock.Text = "Nhập tên mới cho biểu mẫu:";
		textBlock.FontWeight = FontWeights.SemiBold;
		textBlock.Margin = new Thickness(0.0, 0.0, 0.0, 8.0);
		TextBlock element = textBlock;
		Grid.SetRow(element, 0);
		TextBox txt = new TextBox
		{
			Text = currentName,
			Height = 30.0,
			VerticalContentAlignment = VerticalAlignment.Center,
			Padding = new Thickness(6.0, 0.0, 6.0, 0.0)
		};
		Grid.SetRow(txt, 1);
		StackPanel stackPanel = new StackPanel();
		stackPanel.Orientation = Orientation.Horizontal;
		stackPanel.HorizontalAlignment = HorizontalAlignment.Right;
		stackPanel.Margin = new Thickness(0.0, 12.0, 0.0, 0.0);
		StackPanel stackPanel2 = stackPanel;
		Grid.SetRow(stackPanel2, 3);
		Button button = new Button();
		button.Content = "Hủy";
		button.Width = 75.0;
		button.Height = 28.0;
		button.Margin = new Thickness(0.0, 0.0, 8.0, 0.0);
		button.IsCancel = true;
		Button button2 = button;
		Button button3 = new Button();
		button3.Content = "Đồng ý";
		button3.Width = 80.0;
		button3.Height = 28.0;
		button3.IsDefault = true;
		Button button4 = button3;
		button2.Click += (object s, RoutedEventArgs e) =>
		{
			prompt.DialogResult = false;
			prompt.Close();
		};
		button4.Click += (object s, RoutedEventArgs e) =>
		{
			prompt.DialogResult = true;
			prompt.Close();
		};
		stackPanel2.Children.Add(button2);
		stackPanel2.Children.Add(button4);
		grid2.Children.Add(element);
		grid2.Children.Add(txt);
		grid2.Children.Add(stackPanel2);
		prompt.Content = grid2;
		prompt.Loaded += (object s, RoutedEventArgs e) =>
		{
			txt.Focus();
			txt.SelectAll();
		};
		if (prompt.ShowDialog() == true)
		{
			return txt.Text.Trim();
		}
		return null;
	}
}
