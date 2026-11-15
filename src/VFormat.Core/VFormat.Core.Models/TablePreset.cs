using System;
using VFormat.Core.Common;

namespace VFormat.Core.Models;

public class TablePreset
{
	private float _headerFontSizePt = 13f;

	private float _bodyFontSizePt = 13f;

	public float HeaderFontSizePt
	{
		get
		{
			return _headerFontSizePt;
		}
		set
		{
			_headerFontSizePt = (float)Math.Round(value);
		}
	}

	public bool HeaderBold { get; set; }

	public float BodyFontSizePt
	{
		get
		{
			return _bodyFontSizePt;
		}
		set
		{
			_bodyFontSizePt = (float)Math.Round(value);
		}
	}

	public float CellPaddingTopMm { get; set; }

	public float CellPaddingBottomMm { get; set; }

	public float CellPaddingLeftMm { get; set; }

	public float CellPaddingRightMm { get; set; }

	public bool AutoFit { get; set; }

	public bool RepeatHeader { get; set; }

	public bool CantSplit { get; set; }

	public float CellPaddingTopPt => UnitConverter.MmToPt(CellPaddingTopMm);

	public float CellPaddingBottomPt => UnitConverter.MmToPt(CellPaddingBottomMm);

	public float CellPaddingLeftPt => UnitConverter.MmToPt(CellPaddingLeftMm);

	public float CellPaddingRightPt => UnitConverter.MmToPt(CellPaddingRightMm);

	public TablePreset()
	{
		HeaderFontSizePt = 13f;
		HeaderBold = true;
		BodyFontSizePt = 13f;
		CellPaddingTopMm = 1.5f;
		CellPaddingBottomMm = 1.5f;
		CellPaddingLeftMm = 2f;
		CellPaddingRightMm = 2f;
		AutoFit = true;
		RepeatHeader = true;
		CantSplit = true;
	}

	public TablePreset Clone()
	{
		TablePreset tablePreset = new TablePreset();
		tablePreset.HeaderFontSizePt = HeaderFontSizePt;
		tablePreset.HeaderBold = HeaderBold;
		tablePreset.BodyFontSizePt = BodyFontSizePt;
		tablePreset.CellPaddingTopMm = CellPaddingTopMm;
		tablePreset.CellPaddingBottomMm = CellPaddingBottomMm;
		tablePreset.CellPaddingLeftMm = CellPaddingLeftMm;
		tablePreset.CellPaddingRightMm = CellPaddingRightMm;
		tablePreset.AutoFit = AutoFit;
		tablePreset.RepeatHeader = RepeatHeader;
		tablePreset.CantSplit = CantSplit;
		return tablePreset;
	}
}
