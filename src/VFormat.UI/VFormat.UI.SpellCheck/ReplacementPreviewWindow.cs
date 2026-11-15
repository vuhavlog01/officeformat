using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;

namespace VFormat.UI.SpellCheck;

public sealed class ReplacementPreviewWindow : Window
{
	private static Brush Color(string value)
	{
		return (Brush)new BrushConverter().ConvertFromString(value);
	}

	private static TextBlock Text(string value, double size, string color)
	{
		TextBlock textBlock = new TextBlock();
		textBlock.Text = value;
		textBlock.FontSize = size;
		textBlock.Foreground = Color(color);
		textBlock.TextWrapping = TextWrapping.Wrap;
		return textBlock;
	}

	public ReplacementPreviewWindow(IList<ReplacementPreviewItem> items, int skipped, bool all)
	{
		Title = "Xem trước thay thế — OfficeFormat";
		Width = 850.0;
		Height = 700.0;
		MinWidth = 560.0;
		MinHeight = 420.0;
		MaxWidth = SystemParameters.WorkArea.Width;
		MaxHeight = SystemParameters.WorkArea.Height;
		WindowStartupLocation = WindowStartupLocation.CenterOwner;
		FontFamily = new FontFamily("Segoe UI");
		Background = Color("#FDF8F4");
		DockPanel dockPanel = (DockPanel)(Content = new DockPanel());
		StackPanel stackPanel = new StackPanel
		{
			Margin = new Thickness(24.0, 20.0, 24.0, 18.0)
		};
		TextBlock textBlock = Text("Kiểm tra thay đổi trước khi áp dụng", 23.0, "#27150E");
		textBlock.FontWeight = FontWeights.SemiBold;
		stackPanel.Children.Add(textBlock);
		WrapPanel wrapPanel = new WrapPanel
		{
			Margin = new Thickness(0.0, 12.0, 0.0, 10.0)
		};
		wrapPanel.Children.Add(Badge(items.Count + " lỗi sẽ sửa", "#FFF1E6", "#E04F26"));
		if (skipped > 0)
		{
			wrapPanel.Children.Add(Badge(skipped + " lỗi giữ nguyên", "#FFF3D9", "#956000"));
		}
		stackPanel.Children.Add(wrapPanel);
		stackPanel.Children.Add(Text(all ? "Mỗi lỗi dùng gợi ý riêng · Áp dụng toàn bộ danh sách, kể cả mục đang ẩn bởi bộ lọc." : "Kiểm tra từ thay thế và ngữ cảnh trước khi xác nhận.", 13.0, "#6B4C3E"));
		if (skipped > 0)
		{
			stackPanel.Children.Add(Text("Giữ nguyên các lỗi thiếu gợi ý, không tạo thay đổi hoặc chồng vị trí.", 13.0, "#956000"));
		}
		DockPanel.SetDock(stackPanel, Dock.Top);
		dockPanel.Children.Add(stackPanel);
		Grid grid = new Grid
		{
			Margin = new Thickness(24.0, 16.0, 24.0, 18.0)
		};
		grid.RowDefinitions.Add(new RowDefinition
		{
			Height = GridLength.Auto
		});
		grid.RowDefinitions.Add(new RowDefinition
		{
			Height = GridLength.Auto
		});
		grid.Children.Add(Text("Chỉ sửa sau khi xác nhận. Có thể hoàn tác cả lượt bằng Ctrl+Z trong Word.", 12.0, "#6B4C3E"));
		StackPanel stackPanel2 = new StackPanel
		{
			Orientation = Orientation.Horizontal,
			HorizontalAlignment = HorizontalAlignment.Right,
			Margin = new Thickness(0.0, 12.0, 0.0, 0.0)
		};
		Button button = Action("Hủy", primary: false);
		button.IsCancel = true;
		button.IsDefault = true;
		button.Click += (object sender, RoutedEventArgs e) =>
		{
			DialogResult = false;
		};
		Button button2 = Action(all ? ("Xác nhận sửa " + items.Count + " lỗi") : "Xác nhận thay từ này", primary: true);
		button2.Click += (object sender, RoutedEventArgs e) =>
		{
			DialogResult = true;
		};
		stackPanel2.Children.Add(button);
		stackPanel2.Children.Add(button2);
		Grid.SetRow(stackPanel2, 1);
		grid.Children.Add(stackPanel2);
		DockPanel.SetDock(grid, Dock.Bottom);
		dockPanel.Children.Add(grid);
		StackPanel stackPanel3 = new StackPanel
		{
			Margin = new Thickness(24.0, 0.0, 16.0, 0.0)
		};
		for (int num = 0; num < items.Count; num++)
		{
			stackPanel3.Children.Add(Card(items[num], num + 1));
		}
		dockPanel.Children.Add(new ScrollViewer
		{
			VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
			HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
			Content = stackPanel3
		});
	}

	private static Border Badge(string value, string background, string foreground)
	{
		Border border = new Border();
		border.Background = Color(background);
		border.CornerRadius = new CornerRadius(6.0);
		border.Padding = new Thickness(10.0, 5.0, 10.0, 5.0);
		border.Margin = new Thickness(0.0, 0.0, 8.0, 0.0);
		border.Child = Text(value, 13.0, foreground);
		return border;
	}

	private static Button Action(string label, bool primary)
	{
		Button button = new Button();
		button.Content = label;
		button.FontSize = 14.0;
		button.FontWeight = FontWeights.SemiBold;
		button.Padding = new Thickness(20.0, 11.0, 20.0, 11.0);
		button.Margin = new Thickness(10.0, 0.0, 0.0, 0.0);
		button.Background = Color(primary ? "#E04F26" : "#FFFFFF");
		button.Foreground = Color(primary ? "#FFFFFF" : "#27150E");
		button.BorderBrush = Color(primary ? "#E04F26" : "#E2C8B5");
		Button button2 = button;
		button2.Template = (ControlTemplate)XamlReader.Parse("<ControlTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' TargetType='Button'><Border Background='{TemplateBinding Background}' BorderBrush='{TemplateBinding BorderBrush}' BorderThickness='1' CornerRadius='8'><ContentPresenter Margin='{TemplateBinding Padding}' HorizontalAlignment='Center' VerticalAlignment='Center'/></Border><ControlTemplate.Triggers><Trigger Property='IsMouseOver' Value='True'><Setter Property='Opacity' Value='0.88'/></Trigger><Trigger Property='IsKeyboardFocused' Value='True'><Setter Property='BorderBrush' Value='#A63010'/></Trigger><Trigger Property='IsEnabled' Value='False'><Setter Property='Opacity' Value='0.5'/></Trigger></ControlTemplate.Triggers></ControlTemplate>");
		return button2;
	}

	private static Border Card(ReplacementPreviewItem item, int number)
	{
		StackPanel stackPanel = new StackPanel();
		WrapPanel wrapPanel = new WrapPanel();
		wrapPanel.Margin = new Thickness(0.0, 0.0, 0.0, 12.0);
		WrapPanel wrapPanel2 = wrapPanel;
		wrapPanel2.Children.Add(Badge("#" + number + " · Đoạn " + item.Paragraph, "#FAF0E7", "#6B4C3E"));
		TextBlock textBlock = new TextBlock();
		textBlock.FontSize = 17.0;
		textBlock.FontWeight = FontWeights.SemiBold;
		textBlock.TextWrapping = TextWrapping.Wrap;
		textBlock.Margin = new Thickness(0.0, 3.0, 0.0, 0.0);
		TextBlock textBlock2 = textBlock;
		textBlock2.Inlines.Add(new Run(item.Original)
		{
			Foreground = Color("#C4343D")
		});
		textBlock2.Inlines.Add(new Run("  →  ")
		{
			Foreground = Color("#9E7E6F")
		});
		textBlock2.Inlines.Add(new Run(item.Replacement)
		{
			Foreground = Color("#06764B")
		});
		wrapPanel2.Children.Add(textBlock2);
		stackPanel.Children.Add(wrapPanel2);
		int i = 0;
		int j = 0;
		string text = item.Before ?? "";
		string text2;
		for (text2 = item.After ?? ""; i < text.Length && i < text2.Length && text[i] == text2[i]; i++)
		{
		}
		for (; j < text.Length - i && j < text2.Length - i && text[text.Length - 1 - j] == text2[text2.Length - 1 - j]; j++)
		{
		}
		stackPanel.Children.Add(Context("TRƯỚC", text, i, j, after: false));
		stackPanel.Children.Add(Context("SAU", text2, i, j, after: true));
		Border border = new Border();
		border.Background = Brushes.White;
		border.BorderBrush = Color("#EFE0D3");
		border.BorderThickness = new Thickness(1.0);
		border.CornerRadius = new CornerRadius(12.0);
		border.Padding = new Thickness(18.0);
		border.Margin = new Thickness(0.0, 0.0, 0.0, 14.0);
		border.Child = stackPanel;
		return border;
	}

	private static Border Context(string label, string value, int prefix, int suffix, bool after)
	{
		StackPanel stackPanel = new StackPanel();
		TextBlock textBlock = Text(label, 11.0, after ? "#087C50" : "#AD3E48");
		textBlock.FontWeight = FontWeights.Bold;
		textBlock.Margin = new Thickness(0.0, 0.0, 0.0, 5.0);
		stackPanel.Children.Add(textBlock);
		TextBlock textBlock2 = Text("", 14.0, "#27150E");
		textBlock2.LineHeight = 23.0;
		textBlock2.Inlines.Add(new Run(value.Substring(0, prefix)));
		textBlock2.Inlines.Add(new Run(value.Substring(prefix, value.Length - prefix - suffix))
		{
			FontWeight = FontWeights.Bold,
			Foreground = Color(after ? "#06764B" : "#C4343D"),
			Background = Color(after ? "#D5F4E4" : "#FFE0E3")
		});
		textBlock2.Inlines.Add(new Run(value.Substring(value.Length - suffix)));
		stackPanel.Children.Add(textBlock2);
		Border border = new Border();
		border.Background = Color(after ? "#F0FAF5" : "#FFF5F5");
		border.CornerRadius = new CornerRadius(7.0);
		border.Padding = new Thickness(12.0, 9.0, 12.0, 9.0);
		border.Margin = new Thickness(0.0, 0.0, 0.0, 7.0);
		border.Child = stackPanel;
		return border;
	}
}
