using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Markup;

namespace VFormat.UI.Common;

public partial class AuditReportDialog : Window, IComponentConnector
{
	public AuditReportDialog()
	{
		InitializeComponent();
	}

	public AuditReportDialog(string reportText, int commentsCount = 0, int incompleteCount = 0)
	{
		InitializeComponent();
		txtReportContent.Text = reportText ?? "";
		if (commentsCount > 0)
		{
			txtBadge.Text = $"{commentsCount} lỗi/cảnh báo trong lượt rà soát";
		}
		else
		{
			txtBadge.Text = ((incompleteCount > 0) ? "Còn mục chưa kiểm tra được" : "Không phát hiện lỗi trong mục đã kiểm tra");
		}
	}

	private void BtnCopy_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			Clipboard.SetText(txtReportContent.Text);
			txtCopyFeedback.Text = "✓ Đã sao chép nội dung báo cáo vào bộ nhớ tạm!";
		}
		catch (Exception ex)
		{
			txtCopyFeedback.Text = "Không thể sao chép: " + ex.Message;
		}
	}

	private void BtnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	public static void ShowReport(string reportText, int commentsCount = 0, IntPtr? ownerHwnd = null, int incompleteCount = 0)
	{
		AuditReportDialog auditReportDialog = new AuditReportDialog(reportText, commentsCount, incompleteCount);
		if (ownerHwnd.HasValue && ownerHwnd.Value != IntPtr.Zero)
		{
			try
			{
				new WindowInteropHelper(auditReportDialog).Owner = ownerHwnd.Value;
			}
			catch
			{
			}
		}
		auditReportDialog.ShowDialog();
	}
}
