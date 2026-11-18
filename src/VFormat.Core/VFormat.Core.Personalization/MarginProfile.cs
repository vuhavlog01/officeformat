namespace VFormat.Core.Personalization;

public class MarginProfile
{
	private float _firstLineIndentCm = 1f;

	public float MarginTopMm { get; set; }

	public float MarginBottomMm { get; set; }

	public float MarginLeftMm { get; set; }

	public float MarginRightMm { get; set; }

	public float FirstLineIndentCm
	{
		get
		{
			return _firstLineIndentCm;
		}
		set
		{
			_firstLineIndentCm = ((value >= 50f) ? (value / 100f) : value);
		}
	}

	public float SpacingBeforePt { get; set; }

	public float SpacingAfterPt { get; set; }

	public string LineSpacingRule { get; set; }

	public float LineSpacingAtPt { get; set; }

	public bool NoSpaceBetweenParagraphsOfSameStyle { get; set; }

	public MarginProfile()
	{
		MarginTopMm = 20f;
		MarginBottomMm = 20f;
		MarginLeftMm = 30f;
		MarginRightMm = 15f;
		FirstLineIndentCm = 1f;
		SpacingBeforePt = 0f;
		SpacingAfterPt = 6f;
		LineSpacingRule = "1.15 lines";
		LineSpacingAtPt = 0f;
		NoSpaceBetweenParagraphsOfSameStyle = false;
	}

	public MarginProfile(float topMm, float bottomMm, float leftMm, float rightMm, float indentCm, float spacingAfterPt, string lineSpacingRule, float spacingBeforePt = 0f, float lineSpacingAtPt = 0f, bool noSpaceBetweenParagraphsOfSameStyle = false)
	{
		MarginTopMm = topMm;
		MarginBottomMm = bottomMm;
		MarginLeftMm = leftMm;
		MarginRightMm = rightMm;
		FirstLineIndentCm = indentCm;
		SpacingBeforePt = spacingBeforePt;
		SpacingAfterPt = spacingAfterPt;
		LineSpacingRule = lineSpacingRule ?? "1.15 lines";
		LineSpacingAtPt = lineSpacingAtPt;
		NoSpaceBetweenParagraphsOfSameStyle = noSpaceBetweenParagraphsOfSameStyle;
	}

	public MarginProfile Clone()
	{
		return new MarginProfile(MarginTopMm, MarginBottomMm, MarginLeftMm, MarginRightMm, FirstLineIndentCm, SpacingAfterPt, LineSpacingRule, SpacingBeforePt, LineSpacingAtPt, NoSpaceBetweenParagraphsOfSameStyle);
	}

	public static MarginProfile CreateND30Default()
	{
		return new MarginProfile(20f, 20f, 30f, 15f, 1f, 6f, "1.15 lines", 0f, 0f, false);
	}

	public static MarginProfile CreateHD05Default()
	{
		return new MarginProfile(20f, 20f, 30f, 15f, 1f, 6f, "Exactly", 6f, 18f, false);
	}

	public static MarginProfile CreateTT68Default()
	{
		return new MarginProfile(20f, 20f, 30f, 15f, 1f, 6f, "1.15 lines", 0f, 0f, false);
	}

	public static MarginProfile CreateHD29Default()
	{
		return new MarginProfile(20f, 20f, 30f, 20f, 1.27f, 6f, "1.15 lines", 0f, 0f, false);
	}
}
