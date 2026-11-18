using System;
using VFormat.Core.Common;
using VFormat.Core.Enums;

namespace VFormat.Core.Models;

public class PageNumberPreset
{
	private float _sizePt = 14f;

	public float SizePt
	{
		get
		{
			return _sizePt;
		}
		set
		{
			_sizePt = (float)Math.Round(value);
		}
	}

	public PageNumberPosition Position { get; set; }

	public float DistanceFromTopMm { get; set; }

	public bool CenterVerticallyInTopMargin { get; set; }

	public bool HideFirstPage { get; set; }

	public string NumberingFormat { get; set; }

	public bool Bold { get; set; }

	public bool Italic { get; set; }

	public float DistanceFromTopPt => UnitConverter.MmToPt(DistanceFromTopMm);

	public PageNumberPreset()
	{
		SizePt = 14f;
		Position = PageNumberPosition.TopCenter;
		DistanceFromTopMm = 10f;
		HideFirstPage = true;
		NumberingFormat = "Arabic";
		Bold = false;
		Italic = false;
	}

	public PageNumberPreset Clone()
	{
		PageNumberPreset pageNumberPreset = new PageNumberPreset();
		pageNumberPreset.SizePt = SizePt;
		pageNumberPreset.Position = Position;
		pageNumberPreset.DistanceFromTopMm = DistanceFromTopMm;
		pageNumberPreset.CenterVerticallyInTopMargin = CenterVerticallyInTopMargin;
		pageNumberPreset.HideFirstPage = HideFirstPage;
		pageNumberPreset.NumberingFormat = NumberingFormat;
		pageNumberPreset.Bold = Bold;
		pageNumberPreset.Italic = Italic;
		return pageNumberPreset;
	}
}
