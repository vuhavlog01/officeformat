using VFormat.Core.Common;
using VFormat.Core.Enums;

namespace VFormat.Core.Models;

public class LayoutPreset
{
	public string PaperSizeName { get; set; }

	public PaperOrientation Orientation { get; set; }

	public float WidthMm { get; set; }

	public float HeightMm { get; set; }

	public float MarginTopMm { get; set; }

	public float MarginBottomMm { get; set; }

	public float MarginLeftMm { get; set; }

	public float MarginRightMm { get; set; }

	public bool MirrorMargins { get; set; }

	public PrintLayoutMode PrintLayoutMode
	{
		get
		{
			if (!MirrorMargins)
			{
				return PrintLayoutMode.SingleSided;
			}
			return PrintLayoutMode.TwoSidedMirror;
		}
		set
		{
			MirrorMargins = value == PrintLayoutMode.TwoSidedMirror;
		}
	}

	public float DuplexMarginLeftMm { get; set; }

	public float DuplexMarginRightMm { get; set; }

	public float HeaderDistanceMm { get; set; }

	public float FooterDistanceMm { get; set; }

	public float MarginTopPt => UnitConverter.MmToPt(MarginTopMm);

	public float MarginBottomPt => UnitConverter.MmToPt(MarginBottomMm);

	public float MarginLeftPt => UnitConverter.MmToPt(MarginLeftMm);

	public float MarginRightPt => UnitConverter.MmToPt(MarginRightMm);

	public float HeaderDistancePt => UnitConverter.MmToPt(HeaderDistanceMm);

	public float FooterDistancePt => UnitConverter.MmToPt(FooterDistanceMm);

	public LayoutPreset()
	{
		PaperSizeName = "A4";
		Orientation = PaperOrientation.Portrait;
		WidthMm = 210f;
		HeightMm = 297f;
		MarginTopMm = 20f;
		MarginBottomMm = 20f;
		MarginLeftMm = 30f;
		MarginRightMm = 15f;
		MirrorMargins = false;
		DuplexMarginLeftMm = 15f;
		DuplexMarginRightMm = 30f;
		HeaderDistanceMm = 10f;
		FooterDistanceMm = 10f;
	}

	public LayoutPreset Clone()
	{
		LayoutPreset layoutPreset = new LayoutPreset();
		layoutPreset.PaperSizeName = PaperSizeName;
		layoutPreset.Orientation = Orientation;
		layoutPreset.WidthMm = WidthMm;
		layoutPreset.HeightMm = HeightMm;
		layoutPreset.MarginTopMm = MarginTopMm;
		layoutPreset.MarginBottomMm = MarginBottomMm;
		layoutPreset.MarginLeftMm = MarginLeftMm;
		layoutPreset.MarginRightMm = MarginRightMm;
		layoutPreset.MirrorMargins = MirrorMargins;
		layoutPreset.DuplexMarginLeftMm = DuplexMarginLeftMm;
		layoutPreset.DuplexMarginRightMm = DuplexMarginRightMm;
		layoutPreset.HeaderDistanceMm = HeaderDistanceMm;
		layoutPreset.FooterDistanceMm = FooterDistanceMm;
		return layoutPreset;
	}
}
