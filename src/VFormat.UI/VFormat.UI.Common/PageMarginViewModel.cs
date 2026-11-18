using System.Collections.ObjectModel;
using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.UI.Common;

public class PageMarginViewModel : ViewModelBase
{
	private int _selectedRegulationIndex;

	private float _topMarginMm = 20f;

	private float _bottomMarginMm = 20f;

	private float _leftMarginMm = 30f;

	private float _rightMarginMm = 20f;

	private float _headerDistanceMm = 10f;

	private float _footerDistanceMm = 10f;

	private bool _isWholeDocument = true;

	private bool _isPortrait = true;

	private bool _mirrorMargins;

	public ObservableCollection<string> Regulations { get; set; }

	public int SelectedRegulationIndex
	{
		get
		{
			return _selectedRegulationIndex;
		}
		set
		{
			if (SetProperty(ref _selectedRegulationIndex, value, "SelectedRegulationIndex"))
			{
				OnRegulationChanged();
			}
		}
	}

	public float TopMarginMm
	{
		get
		{
			return _topMarginMm;
		}
		set
		{
			SetProperty(ref _topMarginMm, value, "TopMarginMm");
		}
	}

	public float BottomMarginMm
	{
		get
		{
			return _bottomMarginMm;
		}
		set
		{
			SetProperty(ref _bottomMarginMm, value, "BottomMarginMm");
		}
	}

	public float LeftMarginMm
	{
		get
		{
			return _leftMarginMm;
		}
		set
		{
			SetProperty(ref _leftMarginMm, value, "LeftMarginMm");
		}
	}

	public float RightMarginMm
	{
		get
		{
			return _rightMarginMm;
		}
		set
		{
			SetProperty(ref _rightMarginMm, value, "RightMarginMm");
		}
	}

	public float HeaderDistanceMm
	{
		get
		{
			return _headerDistanceMm;
		}
		set
		{
			SetProperty(ref _headerDistanceMm, value, "HeaderDistanceMm");
		}
	}

	public float FooterDistanceMm
	{
		get
		{
			return _footerDistanceMm;
		}
		set
		{
			SetProperty(ref _footerDistanceMm, value, "FooterDistanceMm");
		}
	}

	public bool IsWholeDocument
	{
		get
		{
			return _isWholeDocument;
		}
		set
		{
			if (SetProperty(ref _isWholeDocument, value, "IsWholeDocument"))
			{
				OnPropertyChanged("CurrentSectionOnly");
			}
		}
	}

	public bool CurrentSectionOnly
	{
		get
		{
			return !_isWholeDocument;
		}
		set
		{
			if (value)
			{
				IsWholeDocument = false;
			}
			else
			{
				IsWholeDocument = true;
			}
		}
	}

	public bool IsPortrait
	{
		get
		{
			return _isPortrait;
		}
		set
		{
			if (SetProperty(ref _isPortrait, value, "IsPortrait"))
			{
				OnPropertyChanged("IsLandscape");
			}
		}
	}

	public bool IsLandscape
	{
		get
		{
			return !_isPortrait;
		}
		set
		{
			if (value)
			{
				IsPortrait = false;
			}
			else
			{
				IsPortrait = true;
			}
		}
	}

	public bool MirrorMargins
	{
		get
		{
			return _mirrorMargins;
		}
		set
		{
			SetProperty(ref _mirrorMargins, value, "MirrorMargins");
		}
	}

	public RulePackType SelectedRulePack
	{
		get
		{
			if (_selectedRegulationIndex == 1)
			{
				return RulePackType.HD05_VPTW;
			}
			if (_selectedRegulationIndex == 2)
			{
				return RulePackType.TT68_2023_BQP;
			}
			if (_selectedRegulationIndex == 3)
			{
				return RulePackType.HD29_TWDTN;
			}
			return RulePackType.ND30_2020;
		}
		set
		{
			switch (value)
			{
			case RulePackType.HD05_VPTW:
				SelectedRegulationIndex = 1;
				break;
			case RulePackType.TT68_2023_BQP:
				SelectedRegulationIndex = 2;
				break;
			case RulePackType.HD29_TWDTN:
				SelectedRegulationIndex = 3;
				break;
			default:
				SelectedRegulationIndex = 0;
				break;
			}
		}
	}

	public float TopMargin
	{
		get
		{
			return TopMarginMm;
		}
		set
		{
			TopMarginMm = value;
		}
	}

	public float BottomMargin
	{
		get
		{
			return BottomMarginMm;
		}
		set
		{
			BottomMarginMm = value;
		}
	}

	public float LeftMargin
	{
		get
		{
			return LeftMarginMm;
		}
		set
		{
			LeftMarginMm = value;
		}
	}

	public float RightMargin
	{
		get
		{
			return RightMarginMm;
		}
		set
		{
			RightMarginMm = value;
		}
	}

	public float HeaderDistance
	{
		get
		{
			return HeaderDistanceMm;
		}
		set
		{
			HeaderDistanceMm = value;
		}
	}

	public float FooterDistance
	{
		get
		{
			return FooterDistanceMm;
		}
		set
		{
			FooterDistanceMm = value;
		}
	}

	public PageMarginViewModel()
	{
		Regulations = new ObservableCollection<string> { "Nghị định 30/2020/NĐ-CP (Hành chính Nhà nước)", "Hướng dẫn 05-HD/VPTW (Văn bản Đảng)", "Thông tư 68/2023/TT-BQP (Bộ Quốc phòng)", "Hướng dẫn 29-HD/TWĐTN-VP (Văn bản Đoàn TNCS Hồ Chí Minh)" };
		SelectedRegulationIndex = 0;
	}

	private void OnRegulationChanged()
	{
		TopMarginMm = 20f;
		BottomMarginMm = 20f;
		LeftMarginMm = 30f;
		HeaderDistanceMm = 10f;
		FooterDistanceMm = 10f;
		if (_selectedRegulationIndex == 3)
		{
			RightMarginMm = 20f;
		}
		else if (_selectedRegulationIndex == 1 || _selectedRegulationIndex == 2)
		{
			RightMarginMm = 15f;
		}
		else
		{
			RightMarginMm = 20f;
		}
	}

	public void ResetToStandard()
	{
		OnRegulationChanged();
	}

	public LayoutPreset ToLayoutPreset()
	{
		LayoutPreset layoutPreset = new LayoutPreset();
		layoutPreset.MarginTopMm = TopMarginMm;
		layoutPreset.MarginBottomMm = BottomMarginMm;
		layoutPreset.MarginLeftMm = LeftMarginMm;
		layoutPreset.MarginRightMm = RightMarginMm;
		layoutPreset.HeaderDistanceMm = HeaderDistanceMm;
		layoutPreset.FooterDistanceMm = FooterDistanceMm;
		layoutPreset.MirrorMargins = MirrorMargins;
		layoutPreset.Orientation = (IsPortrait ? PaperOrientation.Portrait : PaperOrientation.Landscape);
		return layoutPreset;
	}
}
