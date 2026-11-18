using System.IO;
using System.Windows;
using System.Windows.Markup;
using Microsoft.Win32;
using VFormat.Core.Templates;

namespace VFormat.UI.Templates;

public partial class AddUserTemplateDialog : Window, IComponentConnector
{
	public UserTemplateItem CreatedTemplateItem { get; private set; }

	public AddUserTemplateDialog(string initialFilePath = null)
	{
		InitializeComponent();
		if (!string.IsNullOrEmpty(initialFilePath) && File.Exists(initialFilePath))
		{
			SetFilePathAndSuggestName(initialFilePath);
			return;
		}
		Loaded += (object s, RoutedEventArgs e) =>
		{
			PromptBrowseFile();
		};
	}

	private void PromptBrowseFile()
	{
		OpenFileDialog openFileDialog = new OpenFileDialog();
		openFileDialog.Title = "Chọn tệp văn bản mẫu Word";
		openFileDialog.Filter = "Tài liệu Word (*.docx;*.doc;*.dotx)|*.docx;*.doc;*.dotx|Tất cả các tệp (*.*)|*.*";
		openFileDialog.CheckFileExists = true;
		OpenFileDialog openFileDialog2 = openFileDialog;
		if (openFileDialog2.ShowDialog(this) == true)
		{
			SetFilePathAndSuggestName(openFileDialog2.FileName);
		}
	}

	private void SetFilePathAndSuggestName(string filePath)
	{
		txtFilePath.Text = filePath;
		string text = UserTemplateService.Instance.SuggestTemplateName(filePath);
		txtTemplateName.Text = text;
		txtTemplateName.Focus();
		txtTemplateName.SelectAll();
	}

	private void BtnBrowse_Click(object sender, RoutedEventArgs e)
	{
		PromptBrowseFile();
	}

	private void BtnSave_Click(object sender, RoutedEventArgs e)
	{
		string text = txtFilePath.Text.Trim();
		if (string.IsNullOrEmpty(text) || !File.Exists(text))
		{
			MessageBox.Show(this, "Vui lòng chọn một tệp Word mẫu hợp lệ trên máy tính.", "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Exclamation);
			return;
		}
		string text2 = txtTemplateName.Text.Trim();
		if (string.IsNullOrEmpty(text2))
		{
			MessageBox.Show(this, "Vui lòng nhập tên hiển thị cho biểu mẫu.", "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Exclamation);
			txtTemplateName.Focus();
			return;
		}
		string description = txtDescription.Text.Trim();
		UserTemplateItem userTemplateItem = UserTemplateService.Instance.ImportTemplate(text, text2, description, out var errorMessage);
		if (userTemplateItem == null)
		{
			MessageBox.Show(this, "Không thể thêm biểu mẫu: " + errorMessage, "OfficeFormat", MessageBoxButton.OK, MessageBoxImage.Hand);
			return;
		}
		CreatedTemplateItem = userTemplateItem;
		DialogResult = true;
		Close();
	}

	private void BtnCancel_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}
}
