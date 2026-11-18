using System;
using VFormat.Core.Common;
using VFormat.Core.Enums;

namespace VFormat.Core.Models;

public class ComponentStyle
{
	private float _sizePt = 14f;

	public SemanticRole Role { get; set; }

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

	public bool Bold { get; set; }

	public bool Italic { get; set; }

	public UnderlineRuleType UnderlineRule { get; set; }

	public AlignmentType Alignment { get; set; }

	public bool IsUppercase { get; set; }

	public bool IsTitleCase { get; set; }

	public string Prefix { get; set; }

	public string Suffix { get; set; }

	public float? SpacingBeforePt { get; set; }

	public float? SpacingAfterPt { get; set; }

	public LineSpacingType? LineSpacingRule { get; set; }

	public float? LineSpacingValue { get; set; }

	public float? FirstLineIndentMm { get; set; }

	public bool? KeepWithNext { get; set; }

	public bool? KeepLinesTogether { get; set; }

	public bool HasBorderBox { get; set; }

	public bool? NoSpaceBetweenParagraphsOfSameStyle { get; set; }

	public float? FirstLineIndentPt
	{
		get
		{
			if (!FirstLineIndentMm.HasValue)
			{
				return null;
			}
			return UnitConverter.MmToPt(FirstLineIndentMm.Value);
		}
	}

	public ComponentStyle()
	{
		SizePt = 14f;
		Bold = false;
		Italic = false;
		UnderlineRule = UnderlineRuleType.None;
		Alignment = AlignmentType.Left;
		IsUppercase = false;
		IsTitleCase = false;
		Prefix = string.Empty;
		Suffix = string.Empty;
		HasBorderBox = false;
	}

	public ComponentStyle(SemanticRole role, float sizePt, bool bold = false, bool italic = false, AlignmentType alignment = AlignmentType.Left, UnderlineRuleType underline = UnderlineRuleType.None)
	{
		Role = role;
		SizePt = sizePt;
		Bold = bold;
		Italic = italic;
		Alignment = alignment;
		UnderlineRule = underline;
		IsUppercase = false;
		IsTitleCase = false;
		Prefix = string.Empty;
		Suffix = string.Empty;
		HasBorderBox = false;
	}

	public ComponentStyle Clone()
	{
		ComponentStyle componentStyle = new ComponentStyle();
		componentStyle.Role = Role;
		componentStyle.SizePt = SizePt;
		componentStyle.Bold = Bold;
		componentStyle.Italic = Italic;
		componentStyle.UnderlineRule = UnderlineRule;
		componentStyle.Alignment = Alignment;
		componentStyle.IsUppercase = IsUppercase;
		componentStyle.IsTitleCase = IsTitleCase;
		componentStyle.Prefix = Prefix;
		componentStyle.Suffix = Suffix;
		componentStyle.SpacingBeforePt = SpacingBeforePt;
		componentStyle.SpacingAfterPt = SpacingAfterPt;
		componentStyle.LineSpacingRule = LineSpacingRule;
		componentStyle.LineSpacingValue = LineSpacingValue;
		componentStyle.FirstLineIndentMm = FirstLineIndentMm;
		componentStyle.KeepWithNext = KeepWithNext;
		componentStyle.KeepLinesTogether = KeepLinesTogether;
		componentStyle.HasBorderBox = HasBorderBox;
		componentStyle.NoSpaceBetweenParagraphsOfSameStyle = NoSpaceBetweenParagraphsOfSameStyle;
		return componentStyle;
	}
}
