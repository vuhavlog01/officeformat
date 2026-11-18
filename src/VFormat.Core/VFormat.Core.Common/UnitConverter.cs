using System;

namespace VFormat.Core.Common;

public static class UnitConverter
{
	public const float MmToPtFactor = 2.8346457f;

	public const float PtToMmFactor = 0.35277778f;

	public const float CmToPtFactor = 28.346457f;

	public const float PtToCmFactor = 0.035277776f;

	public const int TwipsPerPoint = 20;

	public static float MmToPt(float mm)
	{
		return mm * 2.8346457f;
	}

	public static double MmToPt(double mm)
	{
		return mm * 2.834645669291339;
	}

	public static float PtToMm(float pt)
	{
		return pt * 0.35277778f;
	}

	public static double PtToMm(double pt)
	{
		return pt * 0.35277777777777775;
	}

	public static float CmToPt(float cm)
	{
		return cm * 28.346457f;
	}

	public static double CmToPt(double cm)
	{
		return cm * 28.346456692913385;
	}

	public static float PtToCm(float pt)
	{
		return pt * 0.035277776f;
	}

	public static double PtToCm(double pt)
	{
		return pt * 0.035277777777777776;
	}

	public static int PtToTwips(float pt)
	{
		return (int)Math.Round(pt * 20f);
	}

	public static float TwipsToPt(int twips)
	{
		return (float)twips / 20f;
	}

	public static int MmToTwips(float mm)
	{
		return PtToTwips(MmToPt(mm));
	}

	public static float RoundToHalfPoint(float pt)
	{
		return (float)Math.Round(pt * 2f) / 2f;
	}
}
