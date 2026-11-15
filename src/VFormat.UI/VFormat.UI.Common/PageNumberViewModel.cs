using VFormat.Core.Enums;
using VFormat.Core.Models;

namespace VFormat.UI.Common;

public class PageNumberViewModel : ViewModelBase
{
	private bool _isFromStart = true;

	private bool _hideFirstPage = true;

	private int _customStartPage = 2;

	private int _customStartNumber = 1;

	private bool _continuousNumbering = true;

	private PageNumberPosition _position = PageNumberPosition.TopCenter;

	public bool IsFromStart
	{
		get
		{
			return _isFromStart;
		}
		set
		{
			if (SetProperty(ref _isFromStart, value, "IsFromStart"))
			{
				OnPropertyChanged("IsCustomStart");
			}
		}
	}

	public bool IsCustomStart
	{
		get
		{
			return !_isFromStart;
		}
		set
		{
			if (value)
			{
				IsFromStart = false;
			}
			else
			{
				IsFromStart = true;
			}
		}
	}

	public bool HideFirstPage
	{
		get
		{
			return _hideFirstPage;
		}
		set
		{
			SetProperty(ref _hideFirstPage, value, "HideFirstPage");
		}
	}

	public int CustomStartPage
	{
		get
		{
			return _customStartPage;
		}
		set
		{
			SetProperty(ref _customStartPage, value, "CustomStartPage");
		}
	}

	public int CustomStartNumber
	{
		get
		{
			return _customStartNumber;
		}
		set
		{
			SetProperty(ref _customStartNumber, value, "CustomStartNumber");
		}
	}

	public bool ContinuousNumbering
	{
		get
		{
			return _continuousNumbering;
		}
		set
		{
			SetProperty(ref _continuousNumbering, value, "ContinuousNumbering");
		}
	}

	public PageNumberPosition Position
	{
		get
		{
			return _position;
		}
		set
		{
			SetProperty(ref _position, value, "Position");
		}
	}

	public bool IsTopCenter
	{
		get
		{
			return _position == PageNumberPosition.TopCenter;
		}
		set
		{
			if (value)
			{
				Position = PageNumberPosition.TopCenter;
			}
		}
	}

	public bool IsBottomCenter
	{
		get
		{
			return _position == PageNumberPosition.BottomCenter;
		}
		set
		{
			if (value)
			{
				Position = PageNumberPosition.BottomCenter;
			}
		}
	}

	public bool IsBottomRight
	{
		get
		{
			return _position == PageNumberPosition.BottomRight;
		}
		set
		{
			if (value)
			{
				Position = PageNumberPosition.BottomRight;
			}
		}
	}

	public bool FromStart
	{
		get
		{
			return IsFromStart;
		}
		set
		{
			IsFromStart = value;
		}
	}

	public int StartPage
	{
		get
		{
			return CustomStartPage;
		}
		set
		{
			CustomStartPage = value;
		}
	}

	public int StartNumber
	{
		get
		{
			return CustomStartNumber;
		}
		set
		{
			CustomStartNumber = value;
		}
	}

	public bool IsHeader
	{
		get
		{
			return IsTopCenter;
		}
		set
		{
			IsTopCenter = value;
		}
	}

	public PageNumberPreset ToPreset()
	{
		PageNumberPreset pageNumberPreset = new PageNumberPreset();
		pageNumberPreset.SizePt = 14f;
		pageNumberPreset.Position = Position;
		pageNumberPreset.DistanceFromTopMm = 10f;
		pageNumberPreset.HideFirstPage = HideFirstPage;
		return pageNumberPreset;
	}
}
