using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Microsoft.Office.Interop.Word;
using VFormat.UI.TaskPane;

namespace VFormat.Addin;

[Guid("6B68585D-7BF3-4C0E-9214-B9FD286E9CD5")]
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
[ProgId("VFormat.Addin.TaskPaneControl")]
public class TaskPaneControl : UserControl
{
	private ElementHost _elementHost;

	private TaskPaneView _taskPaneView;

	private TaskPaneViewModel _viewModel;

	private Microsoft.Office.Interop.Word.Application _wordApp;

	public TaskPaneViewModel ViewModel => _viewModel;

	public TaskPaneView View => _taskPaneView;

	public TaskPaneControl()
	{
		InitializeComponent();
	}

	private void InitializeComponent()
	{
		base.SuspendLayout();
		base.Name = "TaskPaneControl";
		base.Size = new System.Drawing.Size(360, 800);
		this.BackColor = System.Drawing.SystemColors.Window;
		base.ResumeLayout(false);
	}

	public void Initialize(Microsoft.Office.Interop.Word.Application wordApp, Action onOpenPersonalization, Action onOpenLicense, Action onOpenSpellCheck = null, Action onOpenTemplateLibrary = null)
	{
		_wordApp = wordApp;
		Addin.EnsureWpfApplication();
		_viewModel = new TaskPaneViewModel
		{
			GetActiveDocument = () =>
			{
				try
				{
					return (_wordApp != null) ? _wordApp.ActiveDocument : null;
				}
				catch
				{
					return (object)null;
				}
			},
			GetActiveSelection = () =>
			{
				try
				{
					return (_wordApp != null) ? _wordApp.Selection : null;
				}
				catch
				{
					return (object)null;
				}
			},
			OpenPersonalizationDialogAction = onOpenPersonalization,
			OpenLicenseDialogAction = onOpenLicense,
			OpenSpellCheckDialogAction = onOpenSpellCheck,
			OpenTemplateLibraryDialogAction = onOpenTemplateLibrary
		};
		_taskPaneView = new TaskPaneView
		{
			DataContext = _viewModel
		};
		_elementHost = new ElementHost
		{
			Dock = DockStyle.Fill,
			Child = _taskPaneView
		};
		Controls.Clear();
		Controls.Add(_elementHost);
		_viewModel.RefreshLicenseStatusInBackground();
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && _elementHost != null)
		{
			_elementHost.Dispose();
			_elementHost = null;
		}
		base.Dispose(disposing);
	}
}
