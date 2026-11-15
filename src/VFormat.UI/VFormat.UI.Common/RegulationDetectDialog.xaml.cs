using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Markup;
using VFormat.Core.Enums;

namespace VFormat.UI.Common;

public partial class RegulationDetectDialog : Window, IComponentConnector
{
	public bool Applied { get; private set; }

	public RegulationDetectDialog()
	{
		InitializeComponent();
	}

	public RegulationDetectDialog(RulePackType rulePack, string agency = null, string docNumber = null)
	{
		InitializeComponent();
		Applied = false;
		txtAgency.Text = (string.IsNullOrEmpty(agency) ? "(Được nhận diện tự động từ tài liệu)" : agency);
		txtDocNumber.Text = (string.IsNullOrEmpty(docNumber) ? "(Được nhận diện tự động từ tài liệu)" : docNumber);
		switch (rulePack)
		{
		case RulePackType.HD05_VPTW:
			txtRegulationName.Text = "Văn bản Đảng bộ (Hướng dẫn 05-HD/VPTW)";
			txtTypography.Text = "Font Times New Roman 14pt, Dãn dòng Exactly 20pt, Cách đoạn 0pt, Trích yếu 5 gạch";
			txtMargins.Text = "Trên 20, Dưới 20, Trái 30, Phải 15 mm (Chuẩn HĐ 05)";
			break;
		case RulePackType.TT68_2023_BQP:
			txtRegulationName.Text = "Văn bản Quân đội (Thông tư 68/2023/TT-BQP)";
			txtTypography.Text = "Font Times New Roman 14pt, Dãn dòng 1.15 lines, Cấp bậc quân hàm, Độ Mật/Khẩn";
			txtMargins.Text = "Trên 20, Dưới 20, Trái 30, Phải 15 mm (Chuẩn TT 68 BQP)";
			break;
		case RulePackType.HD29_TWDTN:
			txtRegulationName.Text = "Văn bản Đoàn TNCS Hồ Chí Minh (Hướng dẫn 29-HD/TWĐTN-VP)";
			txtTypography.Text = "Font Times New Roman 14pt, Tiêu đề Đoàn 15pt, Tên loại 16pt, Thụt dòng 1,27cm, Cách đoạn tối đa 6pt";
			txtMargins.Text = "Trên 20, Dưới 20, Trái 30, Phải 20 mm (Chuẩn HD 29)";
			break;
		default:
			txtRegulationName.Text = "Hành chính Nhà nước (Nghị định 30/2020/NĐ-CP)";
			txtTypography.Text = "Font Times New Roman 14pt, Dãn dòng 1.15 lines, Cách đoạn 6pt, Thụt dòng 10mm";
			txtMargins.Text = "Trên 20, Dưới 20, Trái 30, Phải 15-20 mm (Chuẩn NĐ 30)";
			break;
		}
	}

	private void BtnApply_Click(object sender, RoutedEventArgs e)
	{
		Applied = true;
		DialogResult = true;
		Close();
	}

	private void BtnClose_Click(object sender, RoutedEventArgs e)
	{
		Applied = false;
		DialogResult = false;
		Close();
	}

	public static bool ShowDetectResult(RulePackType rulePack, string agency = null, string docNumber = null, IntPtr? ownerHwnd = null)
	{
		RegulationDetectDialog regulationDetectDialog = new RegulationDetectDialog(rulePack, agency, docNumber);
		if (ownerHwnd.HasValue && ownerHwnd.Value != IntPtr.Zero)
		{
			try
			{
				new WindowInteropHelper(regulationDetectDialog).Owner = ownerHwnd.Value;
			}
			catch
			{
			}
		}
		regulationDetectDialog.ShowDialog();
		return regulationDetectDialog.Applied;
	}
}
