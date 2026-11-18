using System.Windows.Controls;
using System.Windows.Markup;

namespace VFormat.UI.TaskPane;

public partial class TaskPaneView : UserControl, IComponentConnector
{
	public TaskPaneView()
	{
		InitializeComponent();
	}

	public TaskPaneView(TaskPaneViewModel viewModel)
		: this()
	{
		DataContext = viewModel;
	}
}
