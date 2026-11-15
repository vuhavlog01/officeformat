using System.Windows;
using System.Windows.Markup;

namespace VFormat.UI.Common;

public partial class ReportDialog : Window, IComponentConnector
{
	public bool KeepChanges { get; private set; }

	public bool KeepFormatting => KeepChanges;

	public ReportDialog()
	{
		InitializeComponent();
		KeepChanges = true;
	}

	public ReportDialog(string text)
		: this()
	{
		SetReportText(text);
	}

	public void SetReportText(string text)
	{
		txtReport.Text = text;
	}

	private void BtnKeep_Click(object sender, RoutedEventArgs e)
	{
		KeepChanges = true;
		DialogResult = true;
		Close();
	}

	private void BtnUndo_Click(object sender, RoutedEventArgs e)
	{
		KeepChanges = false;
		DialogResult = false;
		Close();
	}
}
