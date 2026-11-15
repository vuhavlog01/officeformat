using System;
using VFormat.Core.Common;
using VFormat.Core.Enums;

namespace VFormat.Core.Models;

public class TypographyPreset
{
	private float _bodySizePt = 14f;

	public string FontName { get; set; }

	public string Encoding { get; set; }

	public string FontColorHex { get; set; }

	public float BodySizePt
	{
		get
		{
			return _bodySizePt;
		}
		set
		{
			_bodySizePt = (float)Math.Round(value);
		}
	}

	public LineSpacingType LineSpacingRule { get; set; }

	public float LineSpacingValue { get; set; }

	public float FirstLineIndentMm { get; set; }

	public float SpacingBeforePt { get; set; }

	public float SpacingAfterPt { get; set; }

	public AlignmentType Alignment { get; set; }

	public bool NoSpaceBetweenParagraphsOfSameStyle { get; set; }

	public float FirstLineIndentPt => UnitConverter.MmToPt(FirstLineIndentMm);

	public TypographyPreset()
	{
		FontName = "Times New Roman";
		Encoding = "Unicode TCVN 6909:2001";
		FontColorHex = "#000000";
		BodySizePt = 14f;
		LineSpacingRule = LineSpacingType.Multiple;
		LineSpacingValue = 1.15f;
		FirstLineIndentMm = 10f;
		SpacingBeforePt = 0f;
		SpacingAfterPt = 6f;
		Alignment = AlignmentType.Justify;
	}

	public TypographyPreset Clone()
	{
		TypographyPreset typographyPreset = new TypographyPreset();
		typographyPreset.FontName = FontName;
		typographyPreset.Encoding = Encoding;
		typographyPreset.FontColorHex = FontColorHex;
		typographyPreset.BodySizePt = BodySizePt;
		typographyPreset.LineSpacingRule = LineSpacingRule;
		typographyPreset.LineSpacingValue = LineSpacingValue;
		typographyPreset.FirstLineIndentMm = FirstLineIndentMm;
		typographyPreset.SpacingBeforePt = SpacingBeforePt;
		typographyPreset.SpacingAfterPt = SpacingAfterPt;
		typographyPreset.Alignment = Alignment;
		typographyPreset.NoSpaceBetweenParagraphsOfSameStyle = NoSpaceBetweenParagraphsOfSameStyle;
		return typographyPreset;
	}
}
