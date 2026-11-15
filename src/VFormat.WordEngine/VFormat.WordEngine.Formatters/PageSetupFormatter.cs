using System;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Formatters;

public static class PageSetupFormatter
{
	public static void ApplyPageSetup(Document doc, ProfileModel profile)
	{
		if (doc == null || profile == null)
		{
			return;
		}
		LayoutPreset pageSetup = profile.PageSetup;
		if (pageSetup == null)
		{
			return;
		}
		Sections comObj = null;
		try
		{
			comObj = doc.Sections;
			int count = comObj.Count;
			for (int i = 1; i <= count; i++)
			{
				OperationProgress.Step();
				Section comObj2 = null;
				PageSetup comObj3 = null;
				try
				{
					comObj2 = comObj[i];
					comObj3 = comObj2.PageSetup;
					bool flag = pageSetup.Orientation == PaperOrientation.Landscape || (comObj3.Orientation == WdOrientation.wdOrientLandscape && comObj2.Range.Tables.Count > 0);
					bool flag2 = (profile.RulePack == RulePackType.HD05_VPTW || profile.RulePack == RulePackType.HD29_TWDTN) && string.Equals(pageSetup.PaperSizeName, "A5", StringComparison.OrdinalIgnoreCase);
					if (!flag2 && profile.RulePack == RulePackType.HD29_TWDTN)
					{
						try
						{
							flag2 = comObj3.PaperSize == WdPaperSize.wdPaperA5;
						}
						catch
						{
						}
					}
					comObj3.Orientation = (flag ? WdOrientation.wdOrientLandscape : WdOrientation.wdOrientPortrait);
					try
					{
						comObj3.PaperSize = (flag2 ? WdPaperSize.wdPaperA5 : WdPaperSize.wdPaperA4);
					}
					catch
					{
					}
					try
					{
						float num = (flag2 ? 419.53f : 595.28f);
						float num2 = (flag2 ? 595.28f : 841.89f);
						if (flag)
						{
							comObj3.PageWidth = num2;
							comObj3.PageHeight = num;
						}
						else
						{
							comObj3.PageWidth = num;
							comObj3.PageHeight = num2;
						}
					}
					catch
					{
					}
					comObj3.TopMargin = pageSetup.MarginTopPt;
					comObj3.BottomMargin = pageSetup.MarginBottomPt;
					comObj3.LeftMargin = pageSetup.MarginLeftPt;
					comObj3.RightMargin = pageSetup.MarginRightPt;
					comObj3.HeaderDistance = pageSetup.HeaderDistancePt;
					comObj3.FooterDistance = pageSetup.FooterDistancePt;
					if (pageSetup.MirrorMargins)
					{
						comObj3.MirrorMargins = -1;
					}
					else
					{
						comObj3.MirrorMargins = 0;
					}
					comObj3.DifferentFirstPageHeaderFooter = ((profile.PageNumber != null && profile.PageNumber.HideFirstPage) ? (-1) : 0);
				}
				catch
				{
				}
				finally
				{
					ComReleaser.Release(ref comObj3);
					ComReleaser.Release(ref comObj2);
				}
			}
		}
		catch
		{
		}
		finally
		{
			ComReleaser.Release(ref comObj);
		}
	}
}
