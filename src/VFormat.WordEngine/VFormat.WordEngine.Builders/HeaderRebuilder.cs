using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Rules;
using VFormat.WordEngine.Formatters;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine.Builders;

public static class HeaderRebuilder
{
	private static string ConvertToUnsigned(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return "";
		}
		string pattern = "\\p{IsCombiningDiacriticalMarks}+";
		string input = text.Normalize(NormalizationForm.FormD);
		return Regex.Replace(input, pattern, string.Empty).Replace('đ', 'd').Replace('Đ', 'D');
	}

	public static bool IsDocumentTitle(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = ConvertToUnsigned(text).Trim().ToUpperInvariant();
		if (!(text2 == "QUYET DINH") && !text2.StartsWith("QUYET DINH ") && !text2.StartsWith("QUYET DINH:") && !(text2 == "NGHI QUYET") && !text2.StartsWith("NGHI QUYET ") && !text2.StartsWith("NGHI QUYET:") && !(text2 == "TO TRINH") && !text2.StartsWith("TO TRINH ") && !text2.StartsWith("TO TRINH:") && !(text2 == "KE HOACH") && !text2.StartsWith("KE HOACH ") && !text2.StartsWith("KE HOACH:") && !(text2 == "BAO CAO") && !text2.StartsWith("BAO CAO ") && !text2.StartsWith("BAO CAO:") && !(text2 == "THONG BAO") && !text2.StartsWith("THONG BAO ") && !text2.StartsWith("THONG BAO:") && !(text2 == "CHI THI") && !text2.StartsWith("CHI THI ") && !text2.StartsWith("CHI THI:") && !(text2 == "QUY DINH") && !text2.StartsWith("QUY DINH ") && !text2.StartsWith("QUY DINH:") && !(text2 == "HUONG DAN") && !text2.StartsWith("HUONG DAN ") && !text2.StartsWith("HUONG DAN:") && !(text2 == "CONG DIEN") && !text2.StartsWith("CONG DIEN ") && !(text2 == "GIAY MOI") && !text2.StartsWith("GIAY MOI ") && !(text2 == "BIEN BAN") && !text2.StartsWith("BIEN BAN ") && !(text2 == "GIAY GIOI THIEU") && !text2.StartsWith("GIAY GIOI THIEU ") && !(text2 == "GIAY DI DUONG") && !text2.StartsWith("GIAY DI DUONG ") && !(text2 == "GIAY XAC NHAN") && !text2.StartsWith("GIAY XAC NHAN ") && !(text2 == "GIAY NGHIEP PHAP") && !text2.StartsWith("GIAY NGHI PHEP") && !(text2 == "DON") && !text2.StartsWith("DON ") && !text2.StartsWith("DON:") && !(text2 == "BAN DANG KY") && !text2.StartsWith("BAN DANG KY ") && !text2.StartsWith("BAN DANG KY:") && !(text2 == "BAN CAM KET") && !text2.StartsWith("BAN CAM KET ") && !text2.StartsWith("BAN CAM KET:") && !(text2 == "BAN KIEM DIEM") && !text2.StartsWith("BAN KIEM DIEM ") && !text2.StartsWith("BAN KIEM DIEM:") && !(text2 == "BAN TU KIEM DIEM") && !text2.StartsWith("BAN TU KIEM DIEM ") && !text2.StartsWith("BAN TU KIEM DIEM:") && !(text2 == "BAN TU NHAN XET") && !text2.StartsWith("BAN TU NHAN XET ") && !(text2 == "BAN TUONG TRINH") && !text2.StartsWith("BAN TUONG TRINH ") && !(text2 == "BAN GHI NHO") && !text2.StartsWith("BAN GHI NHO ") && !(text2 == "BAN THOA THUAN") && !text2.StartsWith("BAN THOA THUAN ") && !(text2 == "BAN KHAI") && !text2.StartsWith("BAN KHAI ") && !(text2 == "CHUONG TRINH") && !text2.StartsWith("CHUONG TRINH ") && !text2.StartsWith("CHUONG TRINH:") && !(text2 == "DE AN") && !text2.StartsWith("DE AN ") && !text2.StartsWith("DE AN:") && !(text2 == "DU AN") && !text2.StartsWith("DU AN ") && !text2.StartsWith("DU AN:") && !(text2 == "PHUONG AN") && !text2.StartsWith("PHUONG AN ") && !text2.StartsWith("PHUONG AN:") && !(text2 == "QUY CHE") && !text2.StartsWith("QUY CHE ") && !text2.StartsWith("QUY CHE:") && !(text2 == "THONG CAO") && !text2.StartsWith("THONG CAO ") && !(text2 == "PHIEU DANH GIA") && !text2.StartsWith("PHIEU DANH GIA ") && !(text2 == "PHIEU TRINH") && !text2.StartsWith("PHIEU TRINH ") && !(text2 == "PHIEU XIN Y KIEN") && !text2.StartsWith("PHIEU XIN Y KIEN "))
		{
			return text2.StartsWith("GIAY UY QUYEN");
		}
		return true;
	}

	public static bool IsDocNumberLine(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = text.Trim();
		string text3 = text2.ToLowerInvariant();
		if (text3.StartsWith("sở ") || text3.StartsWith("sở:") || text3.StartsWith("sở-") || text3.StartsWith("sở -") || text3.StartsWith("văn phòng") || text3.StartsWith("chi nhánh") || text3.StartsWith("ban ") || text3.StartsWith("ủy ban"))
		{
			return false;
		}
		if (text3.StartsWith("số:") || text3.StartsWith("số :") || text3.StartsWith("số.") || text3.StartsWith("so:") || text3.StartsWith("so :"))
		{
			return true;
		}
		string text4 = ConvertToUnsigned(text2).ToUpperInvariant();
		if (text4.StartsWith("SO:") || text4.StartsWith("SO: ") || text4.StartsWith("SO. ") || text4.StartsWith("SO :"))
		{
			return true;
		}
		if (text4.StartsWith("SO ") && (text4.StartsWith("SO /") || text4.StartsWith("SO -") || Regex.IsMatch(text4, "^SO\\s+(\\d+|[A-Z]+)?[\\/\\-]")))
		{
			return true;
		}
		return false;
	}

	public static bool IsSeparatorLine(string text)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return false;
		}
		string text2 = text.Replace("\r", "").Replace("\n", "").Replace("\a", "")
			.Trim();
		if (text2.Length < 2)
		{
			return false;
		}
		string text3 = text2.Replace("-", "").Replace("_", "").Replace("=", "")
			.Replace("*", "")
			.Replace("–", "")
			.Replace("–", "")
			.Replace("—", "")
			.Replace("—", "")
			.Replace("―", "")
			.Replace("─", "")
			.Replace("―", "")
			.Replace("─", "")
			.Replace("~", "")
			.Replace(" ", "");
		return text3.Length == 0;
	}

	public static bool IsAsteriskLine(string text)
	{
		return RulePackTraits.IsAsteriskSeparator(text);
	}

	public static bool ContainsOrganizationTitle(string unsignedUpper)
	{
		if (string.IsNullOrEmpty(unsignedUpper))
		{
			return false;
		}
		if (!unsignedUpper.Contains("DANG CONG SAN") && !unsignedUpper.Contains("DOAN TNCS HO CHI MINH"))
		{
			return unsignedUpper.Contains("DOAN THANH NIEN CONG SAN HO CHI MINH");
		}
		return true;
	}

	public static string EmptyNumberPlaceholder(RulePackType rulePack)
	{
		if (RulePackTraits.IsYouthUnion(rulePack))
		{
			return "Số:     -...";
		}
		if (RulePackTraits.IsParty(rulePack))
		{
			return "Số     -...";
		}
		return "Số:      /...";
	}

	public static bool IsSingleAgencySplitInTwo(string lineA, string lineB)
	{
		if (string.IsNullOrWhiteSpace(lineA) || string.IsNullOrWhiteSpace(lineB))
		{
			return false;
		}
		string text = ConvertToUnsigned(lineA).Trim().ToUpperInvariant();
		string text2 = ConvertToUnsigned(lineB).Trim().ToUpperInvariant();
		if (text2.StartsWith("VA ") || text2.StartsWith("& ") || text2.StartsWith("- ") || text2.StartsWith("— "))
		{
			return true;
		}
		if (text.EndsWith(" TINH") || text.EndsWith(" THANH PHO") || text.EndsWith(" TP") || text.EndsWith(" TP.") || text.EndsWith(" HUYEN") || text.EndsWith(" QUAN") || text.EndsWith(" THI XA") || text.EndsWith(" TX") || text.EndsWith(" TX.") || text.EndsWith(" XA") || text.EndsWith(" PHUONG") || text.EndsWith(" THI TRAN") || text.EndsWith(" TT") || text.EndsWith(" TT.") || text.EndsWith(" VA") || text.EndsWith(" &") || text.EndsWith(" -"))
		{
			return true;
		}
		if (lineA.Trim().EndsWith(",") || lineA.Trim().EndsWith(";") || lineA.Trim().EndsWith(" -") || lineA.Trim().EndsWith("-"))
		{
			return true;
		}
		if (text2.StartsWith("TRUYEN HINH") || text2.StartsWith("PHAT THANH") || text2.StartsWith("THONG TAN") || text2.StartsWith("BAO CHI"))
		{
			return true;
		}
		bool flag = text2.StartsWith("XA ") || text2.StartsWith("XA.") || text2.StartsWith("XA:") || text2.StartsWith("PHUONG ") || text2.StartsWith("PHUONG.") || text2.StartsWith("THI TRAN ") || text2.StartsWith("TT ") || text2.StartsWith("TT.") || text2.StartsWith("HUYEN ") || text2.StartsWith("QUAN ") || text2.StartsWith("THI XA ") || text2.StartsWith("TX ") || text2.StartsWith("TX.") || text2.StartsWith("TINH ") || text2.StartsWith("THANH PHO ") || text2.StartsWith("TP ") || text2.StartsWith("TP.") || text2.StartsWith("KHU VUC ") || text2.StartsWith("CO SO ") || text2.StartsWith("DOANH NGHIEP ") || text2.StartsWith("CONG TY ") || text2.StartsWith("CHI NHANH ") || text2.StartsWith("...") || text2.StartsWith("......") || text2.Contains(".....");
		int num;
		switch (text)
		{
		default:
			if (!text.StartsWith("BAN CHI HUY") && !text.StartsWith("UY BAN"))
			{
				num = (text.StartsWith("HOI DONG") ? 1 : 0);
				break;
			}
			goto case "BAN CHI HUY QUAN SU";
		case "BAN CHI HUY QUAN SU":
		case "BAN CHQS":
		case "BCH QUAN SU":
		case "BCHQS":
		case "BO CHI HUY QUAN SU":
		case "BO CHQS":
		case "UY BAN NHAN DAN":
		case "UBND":
		case "HOI DONG NHAN DAN":
		case "HDND":
		case "BO CHI HUY BIEN PHONG":
		case "BCH BDBP":
		case "BDBP":
		case "CONG AN":
		case "DANG UY":
		case "CHI BO":
		case "CHI UY":
		case "TRUONG":
		case "TRUONG TIEU HOC":
		case "TRUONG THCS":
		case "TRUONG THPT":
		case "TRUONG MAM NON":
		case "TRUONG DAI HOC":
		case "TRUONG CAO DANG":
		case "LIEN DOI TRUONG":
		case "LIEN DOI TRUONG TIEU HOC":
		case "LIEN DOI TRUONG THCS":
		case "LIEN DOI TRUONG THPT":
		case "TRUONG TRUNG CAP":
		case "HOC VIEN":
		case "TRAM Y TE":
		case "TRUNG TAM Y TE":
		case "BENH VIEN":
		case "PHONG GIAO DUC VA DAO TAO":
		case "PHONG GD&DT":
		case "PHONG TAI NGUYEN VA MOI TRUONG":
		case "PHONG TN&MT":
		case "PHONG TAI CHINH - KE HOACH":
		case "PHONG TC-KH":
		case "PHONG NOI VU":
		case "PHONG LAO DONG - THUONG BINH VA XA HOI":
		case "PHONG Y TE":
		case "PHONG TU PHAP":
		case "PHONG KINH TE":
		case "PHONG VAN HOA VA THONG TIN":
			num = 1;
			break;
		}
		bool flag2 = (byte)num != 0;
		Match match = RegexPatterns.YouthUnionAgency.Match(lineB.Trim());
		if (match.Success && match.Index == 0)
		{
			return false;
		}
		if (flag)
		{
			return true;
		}
		if (flag2 && !ContainsLocality(text))
		{
			return true;
		}
		return false;
	}

	private static bool ContainsLocality(string uText)
	{
		if (string.IsNullOrEmpty(uText))
		{
			return false;
		}
		if (!uText.Contains("TINH ") && !uText.Contains("THANH PHO ") && !uText.Contains("TP.") && !uText.Contains("TP ") && !uText.Contains("HUYEN ") && !uText.Contains("QUAN ") && !uText.Contains("THI XA ") && !uText.Contains("TX ") && !uText.Contains("XA ") && !uText.Contains("PHUONG ") && !uText.Contains("THI TRAN ") && !uText.Contains("TT."))
		{
			return uText.Contains("...");
		}
		return true;
	}

	public static HashSet<int> IdentifyIssuingAgencyParaIndices(Cell c1, List<int> agencyIndices)
	{
		HashSet<int> hashSet = new HashSet<int>();
		if (agencyIndices == null || agencyIndices.Count == 0)
		{
			return hashSet;
		}
		if (agencyIndices.Count == 1)
		{
			hashSet.Add(agencyIndices[0]);
			return hashSet;
		}
		List<string> list = new List<string>();
		for (int i = 0; i < agencyIndices.Count; i++)
		{
			list.Add((c1.Range.Paragraphs[agencyIndices[i]].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim());
		}
		if (agencyIndices.Count == 2)
		{
			if (IsSingleAgencySplitInTwo(list[0], list[1]))
			{
				hashSet.Add(agencyIndices[0]);
				hashSet.Add(agencyIndices[1]);
			}
			else
			{
				hashSet.Add(agencyIndices[1]);
			}
			return hashSet;
		}
		if (agencyIndices.Count == 3)
		{
			if (IsSingleAgencySplitInTwo(list[0], list[1]))
			{
				hashSet.Add(agencyIndices[2]);
				return hashSet;
			}
			if (IsSingleAgencySplitInTwo(list[1], list[2]))
			{
				hashSet.Add(agencyIndices[1]);
				hashSet.Add(agencyIndices[2]);
				return hashSet;
			}
			hashSet.Add(agencyIndices[2]);
			return hashSet;
		}
		string lineA = list[agencyIndices.Count - 2];
		string lineB = list[agencyIndices.Count - 1];
		if (IsSingleAgencySplitInTwo(lineA, lineB))
		{
			hashSet.Add(agencyIndices[agencyIndices.Count - 2]);
			hashSet.Add(agencyIndices[agencyIndices.Count - 1]);
			return hashSet;
		}
		hashSet.Add(agencyIndices[agencyIndices.Count - 1]);
		return hashSet;
	}

	public static string NormalizeDocNumber(string rawNum, RulePackType rulePack)
	{
		if (RulePackTraits.IsYouthUnion(rulePack))
		{
			string input = (rawNum ?? "").Trim();
			Match match = Regex.Match(input, "^(?:S[ốoÔôỐ]\\s*[:.]?\\s*)?(\\d{0,4})\\s*/\\s*([A-ZĐ]{1,6})\\s*$", RegexOptions.IgnoreCase);
			if (match.Success)
			{
				string text = match.Groups[1].Value;
				if (text.Length == 1)
				{
					text = "0" + text;
				}
				return "Số: " + ((text.Length == 0) ? "     " : text) + "/" + match.Groups[2].Value.ToUpperInvariant();
			}
			string text2 = NormalizeDocNumber(rawNum, RulePackType.HD05_VPTW);
			if (string.IsNullOrEmpty(text2))
			{
				return text2;
			}
			if (text2.StartsWith("Số "))
			{
				return "Số: " + text2.Substring(3);
			}
			return text2;
		}
		string text3 = (rawNum ?? "").Trim();
		bool flag = rulePack == RulePackType.HD05_VPTW;
		if (string.IsNullOrEmpty(text3))
		{
			return "";
		}
		string text4 = (flag ? "Số " : "Số: ");
		string text5 = text3;
		string text6 = ConvertToUnsigned(text5).ToUpperInvariant();
		if (text6.StartsWith("SO: ") || text6.StartsWith("SO. ") || text6.StartsWith("SO : "))
		{
			text5 = text5.Substring(4).Trim();
		}
		else if (text6.StartsWith("SO:") || text6.StartsWith("SO "))
		{
			text5 = text5.Substring(3).Trim();
		}
		while (text5.Contains("  "))
		{
			text5 = text5.Replace("  ", " ");
		}
		text5 = Regex.Replace(text5, "(\\d)\\s+/\\s*", "$1/");
		if (flag)
		{
			string text7 = text5.TrimStart();
			if (text7.StartsWith("/") || text7.StartsWith("-"))
			{
				string text8 = text7.TrimStart('/', '-', ' ').Trim();
				if (text8.Contains("-") && !text8.Contains("/"))
				{
					int num = text8.IndexOf('-');
					text8 = text8.Substring(0, num) + "/" + text8.Substring(num + 1);
				}
				if (string.IsNullOrEmpty(text8))
				{
					text8 = "...";
				}
				return "Số     -" + text8;
			}
			int num2 = text5.IndexOf('-');
			int num3 = text5.IndexOf('/');
			int num4 = -1;
			if (num2 >= 0 && num3 >= 0)
			{
				num4 = Math.Min(num2, num3);
			}
			else if (num2 >= 0)
			{
				num4 = num2;
			}
			else if (num3 >= 0)
			{
				num4 = num3;
			}
			if (num4 > 0)
			{
				string text9 = text5.Substring(0, num4).Trim();
				string text10 = text5.Substring(num4 + 1).Trim();
				if (text10.Contains("-") && !text10.Contains("/"))
				{
					int num5 = text10.IndexOf('-');
					text10 = text10.Substring(0, num5) + "/" + text10.Substring(num5 + 1);
				}
				if (string.IsNullOrEmpty(text10))
				{
					text10 = "...";
				}
				if (int.TryParse(text9, out var result) && result >= 1 && result <= 9 && text9.Length == 1)
				{
					text9 = "0" + result;
				}
				if (string.IsNullOrEmpty(text9) || text9 == "..." || !Regex.IsMatch(text9, "\\d"))
				{
					return "Số     -" + text10;
				}
				return "Số " + text9 + "-" + text10;
			}
			if (int.TryParse(text5, out var result2) && result2 >= 1 && result2 <= 9 && text5.Length == 1)
			{
				text5 = "0" + result2;
			}
			if (string.IsNullOrEmpty(text5) || text5 == "...")
			{
				return "Số     -...";
			}
			if (Regex.IsMatch(text5, "^\\d+$"))
			{
				return "Số " + text5 + "-...";
			}
			return "Số     -" + text5.TrimStart('/', '-', ' ');
		}
		string text11 = text5.TrimStart();
		if (text11.StartsWith("/") || text11.StartsWith("-"))
		{
			string text12 = text11.TrimStart('/', '-', ' ').Trim();
			return text4 + "      /" + text12;
		}
		int num6 = text5.IndexOf('/');
		if (num6 > 0)
		{
			string text13 = text5.Substring(0, num6).Trim();
			string text14 = text5.Substring(num6 + 1).Trim();
			if (string.IsNullOrEmpty(text13))
			{
				text13 = "      ";
			}
			if (string.IsNullOrEmpty(text14))
			{
				text14 = "...";
			}
			return text4 + text13 + "/" + text14;
		}
		return text4 + text5;
	}

	public static string NormalizeDocDate(string rawDate)
	{
		string text = (rawDate ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
			.Replace("\v", "")
			.Trim();
		if (string.IsNullOrEmpty(text))
		{
			return "..., ngày     tháng     năm 20...";
		}
		string text2 = ConvertToUnsigned(text).ToUpperInvariant();
		int num = text2.IndexOf("NGAY");
		if (num < 0)
		{
			return text;
		}
		string text3 = text.Substring(0, num).Trim();
		if (text3.EndsWith(","))
		{
			text3 = text3.Substring(0, text3.Length - 1).Trim();
		}
		if (string.IsNullOrEmpty(text3))
		{
			text3 = "...";
		}
		int num2 = text2.IndexOf("THANG", num + 4);
		if (num2 < 0)
		{
			string text4 = text.Substring(num + 4).Trim();
			return (text3 + ", ngày " + text4).Trim();
		}
		string text5 = text.Substring(num + 4, num2 - (num + 4)).Trim();
		string text6;
		if (text5.Contains(".") || text5.Contains("…"))
		{
			text6 = " " + text5 + " ";
		}
		else
		{
			Match match = Regex.Match(text5, "\\d+");
			text6 = ((!match.Success || !int.TryParse(match.Value, out var result) || result < 1 || result > 31) ? "     " : (" " + result.ToString("D2") + " "));
		}
		int num3 = text2.IndexOf("NAM", num2 + 5);
		if (num3 < 0)
		{
			string text7 = text.Substring(num2 + 5).Trim();
			return (text3 + ", ngày" + text6 + "tháng " + text7).Trim();
		}
		string text8 = text.Substring(num2 + 5, num3 - (num2 + 5)).Trim();
		string text9;
		if (text8.Contains(".") || text8.Contains("…"))
		{
			text9 = " " + text8 + " ";
		}
		else
		{
			Match match2 = Regex.Match(text8, "\\d+");
			if (match2.Success && int.TryParse(match2.Value, out var result2) && result2 >= 1 && result2 <= 12)
			{
				string text10 = ((result2 <= 2) ? result2.ToString("D2") : result2.ToString());
				text9 = " " + text10 + " ";
			}
			else
			{
				text9 = "     ";
			}
		}
		string text11 = text.Substring(num3 + 3).Trim();
		string text12 = ((!string.IsNullOrEmpty(text11)) ? (" " + text11) : " 20...");
		return (text3 + ", ngày" + text6 + "tháng" + text9 + "năm" + text12).Trim();
	}

	public static Table FindExistingHeaderTable(Document doc)
	{
		if (doc == null || doc.Tables.Count == 0)
		{
			return null;
		}
		int num = doc.Content.End;
		int num2 = 0;
		foreach (Paragraph paragraph in doc.Paragraphs)
		{
			if (paragraph.Range.Tables.Count == 0 && ++num2 > 25)
			{
				num = paragraph.Range.Start;
				break;
			}
		}
		int count = doc.Tables.Count;
		for (int i = 1; i <= count; i++)
		{
			Table table = doc.Tables[i];
			if (table.Range.Start >= num)
			{
				break;
			}
			if (SignatureRebuilder.HasAtLeastTwoColumns(table))
			{
				string text = table.Range.Text;
				string text2 = ConvertToUnsigned(text).ToUpperInvariant();
				if (text2.Contains("CONG HOA XA HOI") || ContainsOrganizationTitle(text2) || text2.Contains("DOC LAP - TU DO") || (text2.Contains("SO ") && text2.Contains("NGAY ")) || (text2.Contains("UBND") && (text2.Contains("SO:") || text2.Contains("SO "))))
				{
					return table;
				}
			}
		}
		return null;
	}

	public static bool IsHeaderDateInTopCell(ProfileModel profile = null)
	{
		if (profile != null)
		{
			return profile.HeaderDateInTopCell;
		}
		try
		{
			return new VFormat.Core.Serialization.JsonProfileRepository().LoadPersonalization()?.HeaderDateInTopCell ?? false;
		}
		catch
		{
			return false;
		}
	}

	public static bool NeedsRebuild(Table tbl, RulePackType rulePack)
	{
		if (tbl == null)
		{
			return false;
		}
		try
		{
			if (tbl.Rows.Count != 2)
			{
				return true;
			}
			if (tbl.Columns.Count < 2)
			{
				return true;
			}
			Cell cell = tbl.Cell(2, 1);
			string text = (cell.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			string text2 = ConvertToUnsigned(text).ToUpperInvariant();
			if (!IsDocNumberLine(text) && (RegexPatterns.PartyAgency.IsMatch(text) || RegexPatterns.AdminAgency.IsMatch(text) || RegexPatterns.MilitaryAgency.IsMatch(text) || text2.StartsWith("CHI BO") || text2.StartsWith("DANG BO") || text2.StartsWith("DANG UY") || text2.StartsWith("CHI UY") || text2.StartsWith("BAN ") || text2.StartsWith("SO ") || text2.StartsWith("PHONG ") || text2.StartsWith("UY BAN") || text2.StartsWith("HOI DONG")))
			{
				return true;
			}
			Cell cell2 = tbl.Cell(1, 1);
			string text3 = (cell2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			if (IsDocNumberLine(text3))
			{
				return true;
			}
			return false;
		}
		catch
		{
			return true;
		}
	}

	public static bool IsPersonalDocument(Document doc)
	{
		if (doc == null || doc.Paragraphs.Count == 0)
		{
			return false;
		}
		Table table = FindExistingHeaderTable(doc);
		if (table != null)
		{
			string text = ConvertToUnsigned(table.Range.Text ?? "").ToUpperInvariant();
			if (text.Contains("UBND") || text.Contains("BO ") || text.Contains("SO ") || ContainsOrganizationTitle(text) || text.Contains("BAN ") || text.Contains("DOAN ") || text.Contains("CUC "))
			{
				return false;
			}
		}
		int num = Math.Min(doc.Paragraphs.Count, 10);
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		for (int i = 1; i <= num; i++)
		{
			Paragraph paragraph = doc.Paragraphs[i];
			if (paragraph.Range.Tables.Count > 0)
			{
				continue;
			}
			string text2 = (paragraph.Range.Text ?? "").Trim();
			if (text2.Length != 0)
			{
				string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
				if (text3.Contains("CONG HOA XA HOI") || text3.Contains("DOC LAP - TU DO"))
				{
					flag = true;
				}
				if (text3.StartsWith("DON ") || text3 == "DON" || text3.StartsWith("BAN TUONG TRINH") || text3.StartsWith("BAN CAM KET") || text3.StartsWith("GIAY UY QUYEN") || text3.StartsWith("GIAY XIN ") || text3.Contains("DON XIN ") || text3.Contains("DON DE NGHI") || text3.Contains("DON KHIEU NAI") || text3.Contains("DON TO CAO"))
				{
					flag2 = true;
				}
				if (RegexPatterns.AdminAgency.IsMatch(text2) || RegexPatterns.PartyAgency.IsMatch(text2) || text3.StartsWith("SO:") || text3.StartsWith("SO :") || text3.Contains("UBND ") || text3.Contains("TINH UY"))
				{
					flag3 = true;
				}
			}
		}
		if (flag2 && !flag3)
		{
			return true;
		}
		if (flag && !flag3 && flag2)
		{
			return true;
		}
		return false;
	}

	public static void FormatPersonalDocumentHeader(Document doc)
	{
		if (doc == null || doc.Paragraphs.Count == 0)
		{
			return;
		}
		int num = Math.Min(doc.Paragraphs.Count, 6);
		Paragraph paragraph = null;
		Paragraph paragraph2 = null;
		for (int i = 1; i <= num; i++)
		{
			Paragraph paragraph3 = doc.Paragraphs[i];
			if (paragraph3.Range.Tables.Count <= 0)
			{
				string text = (paragraph3.Range.Text ?? "").Trim();
				string text2 = ConvertToUnsigned(text).ToUpperInvariant();
				if (text2.Contains("CONG HOA XA HOI") && paragraph == null)
				{
					paragraph = paragraph3;
				}
				else if (text2.Contains("DOC LAP - TU DO") && paragraph2 == null)
				{
					paragraph2 = paragraph3;
				}
			}
		}
		if (paragraph != null)
		{
			paragraph.Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r";
			paragraph.Range.Font.Name = "Times New Roman";
			paragraph.Range.Font.Size = 12f;
			paragraph.Range.Font.Bold = 1;
			paragraph.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			paragraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			paragraph.Range.ParagraphFormat.SpaceBefore = 0f;
			paragraph.Range.ParagraphFormat.SpaceAfter = 0f;
		}
		if (paragraph2 != null)
		{
			paragraph2.Range.Text = "Độc lập - Tự do - Hạnh phúc\r";
			paragraph2.Range.Font.Name = "Times New Roman";
			paragraph2.Range.Font.Size = 13f;
			paragraph2.Range.Font.Bold = 1;
			paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
			paragraph2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			paragraph2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			paragraph2.Range.ParagraphFormat.SpaceBefore = 0f;
			paragraph2.Range.ParagraphFormat.SpaceAfter = 6f;
			UnderlineFormatter.EnsurePersonalMottoShapeLine(doc, paragraph2);
		}
	}

	internal static bool HasVisualContent(Document doc, Range range)
	{
		try
		{
			if (range.InlineShapes.Count > 0)
			{
				return true;
			}
			foreach (Shape shape in doc.Shapes)
			{
				dynamic val = shape;
				if (((int)val.Type != 9 || !(shape.Name ?? "").StartsWith("VFormat_", StringComparison.Ordinal)) && shape.Anchor.Start >= range.Start && shape.Anchor.Start < range.End)
				{
					return true;
				}
			}
			string[] symbolFontNames = WordFormatEngine.SymbolFontNames;
			foreach (string name in symbolFontNames)
			{
				Range comObj = range.Duplicate;
				Find comObj2 = comObj.Find;
				try
				{
					comObj2.ClearFormatting();
					comObj2.Text = "";
					comObj2.Format = true;
					comObj2.MatchWildcards = false;
					comObj2.Forward = true;
					comObj2.Wrap = WdFindWrap.wdFindStop;
					comObj2.Font.Name = name;
					Find find = comObj2;
					object FindText = Missing.Value;
					object MatchCase = Missing.Value;
					object MatchWholeWord = Missing.Value;
					object MatchWildcards = Missing.Value;
					object MatchSoundsLike = Missing.Value;
					object MatchAllWordForms = Missing.Value;
					object Forward = Missing.Value;
					object Wrap = Missing.Value;
					object Format = Missing.Value;
					object ReplaceWith = Missing.Value;
					object Replace = Missing.Value;
					object MatchKashida = Missing.Value;
					object MatchDiacritics = Missing.Value;
					object MatchAlefHamza = Missing.Value;
					object MatchControl = Missing.Value;
					if (find.Execute(ref FindText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl))
					{
						return true;
					}
				}
				finally
				{
					ComReleaser.Release(ref comObj2);
					ComReleaser.Release(ref comObj);
				}
			}
			return false;
		}
		catch
		{
			return true;
		}
	}

	private static int DeleteConsumedHeaderParagraphs(Document doc, List<int> consumed)
	{
		if (consumed == null || consumed.Count == 0)
		{
			return 0;
		}
		consumed.Sort();
		int result = 0;
		try
		{
			result = doc.Paragraphs[consumed[0]].Range.Start;
		}
		catch
		{
		}
		for (int num = consumed.Count - 1; num >= 0; num--)
		{
			try
			{
				if (consumed[num] <= doc.Paragraphs.Count)
				{
					Range range = doc.Paragraphs[consumed[num]].Range;
					if (range.Tables.Count <= 0)
					{
						object Unit = Missing.Value;
						object Count = Missing.Value;
						range.Delete(ref Unit, ref Count);
					}
				}
			}
			catch
			{
			}
		}
		return result;
	}

	internal static int EnsureNotAdjacentToTable(Document doc, int pos)
	{
		try
		{
			pos = Math.Min(pos, Math.Max(0, doc.Content.End - 1));
			if (pos > 0)
			{
				object Start = pos - 1;
				object End = pos;
				if (doc.Range(ref Start, ref End).Tables.Count > 0)
				{
					object Start2 = pos;
					object End2 = pos;
					doc.Range(ref Start2, ref End2).InsertParagraphBefore();
					pos++;
				}
			}
			object Start3 = pos;
			object End3 = Math.Min(pos + 1, doc.Content.End);
			if (doc.Range(ref Start3, ref End3).Tables.Count > 0)
			{
				object Start4 = pos;
				object End4 = Math.Min(pos + 1, doc.Content.End);
				Table table = doc.Range(ref Start4, ref End4).Tables[1];
				Range duplicate = doc.Application.Selection.Range.Duplicate;
				try
				{
					int start = table.Range.Start;
					object Start5 = start;
					object End5 = start;
					doc.Range(ref Start5, ref End5).Select();
					doc.Application.Selection.SplitTable();
				}
				finally
				{
					duplicate.Select();
				}
			}
		}
		catch
		{
		}
		return pos;
	}

	internal static bool IsMottoText(string unsignedUpper)
	{
		if (string.IsNullOrEmpty(unsignedUpper))
		{
			return false;
		}
		if (unsignedUpper.Contains("DOC LAP") && unsignedUpper.Contains("TU DO"))
		{
			return unsignedUpper.Contains("HANH PHUC");
		}
		return false;
	}

	internal static bool IsNationalOrMottoText(string unsignedUpper)
	{
		if (string.IsNullOrEmpty(unsignedUpper))
		{
			return false;
		}
		if (!unsignedUpper.Contains("CONG HOA XA HOI") && !ContainsOrganizationTitle(unsignedUpper) && !unsignedUpper.Contains("DOC LAP - TU DO"))
		{
			return IsMottoText(unsignedUpper);
		}
		return true;
	}

	internal static bool IsBareNumberPrefix(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		return Regex.IsMatch(ConvertToUnsigned(text.Trim()).ToUpperInvariant(), "^SO\\s*[:.]?$");
	}

	internal static bool IsHeaderPiece(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string text2 = ConvertToUnsigned(text).ToUpperInvariant();
		if (IsNationalOrMottoText(text2))
		{
			return true;
		}
		if (IsDocNumberLine(text) && !IsBareNumberPrefix(text))
		{
			return true;
		}
		if (text2.Contains("NGAY ") && text2.Contains(" THANG "))
		{
			return text2.Contains(" NAM ");
		}
		return false;
	}

	internal static void ApplyAgencyAllCaps(Table tbl)
	{
		try
		{
			foreach (Paragraph paragraph in tbl.Cell(1, 1).Range.Paragraphs)
			{
				string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				if (text.Length != 0 && !IsAsteriskLine(text) && !IsSeparatorLine(text) && !IsDocNumberLine(text) && !RegexPatterns.DispatchSubject.IsMatch(text))
				{
					paragraph.Range.Font.AllCaps = -1;
				}
			}
		}
		catch
		{
		}
	}

	internal static void ResetInheritedFormatting(Table tbl)
	{
		if (tbl == null)
		{
			return;
		}
		try
		{
			tbl.Range.Font.Reset();
		}
		catch
		{
		}
		try
		{
			tbl.Range.ParagraphFormat.Reset();
		}
		catch
		{
		}
		try
		{
			tbl.Range.HighlightColorIndex = WdColorIndex.wdAuto;
		}
		catch
		{
		}
	}

	public static bool ConvertRawHeaderToTable(Document doc, RulePackType rulePack, ProfileModel profile = null)
	{
		if (doc == null || doc.Paragraphs.Count < 2)
		{
			return false;
		}
		int num = Math.Min(doc.Paragraphs.Count, 12);
		string text = "";
		string text2 = "";
		string text3 = "";
		bool flag = false;
		int num2 = 0;
		List<int> list = new List<int>();
		int num3 = Math.Min(doc.Paragraphs.Count, 80);
		for (int i = 1; i <= num; i++)
		{
			Paragraph paragraph = doc.Paragraphs[i];
			if (paragraph.Range.Tables.Count > 0)
			{
				if (num < num3)
				{
					num++;
				}
				continue;
			}
			string text4 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
				.Trim();
			if (text4.Length <= 0)
			{
				continue;
			}
			string text5 = ConvertToUnsigned(text4).ToUpperInvariant();
			if (IsDocumentTitle(text5) || RegexPatterns.DocumentType.IsMatch(text4) || text5.StartsWith("V/V") || text5.StartsWith("VE VIEC") || text5.StartsWith("VE ") || text5.StartsWith("KINH GUI") || text5.StartsWith("NOI NHAN") || text5.StartsWith("CAN CU") || text5.StartsWith("THEO ") || SignatureRebuilder.IsSignatureOrRecipientBlock(text4) || (!string.IsNullOrEmpty(text3) && !text4.Contains("\t") && !text4.Contains("   ")))
			{
				break;
			}
			string text6 = "";
			string text7 = "";
			bool flag2 = false;
			if (text4.Contains("\t"))
			{
				string[] array = text4.Split('\t');
				string[] array2 = array;
				foreach (string text8 in array2)
				{
					string text9 = text8.Trim();
					if (text9.Length > 0)
					{
						if (string.IsNullOrEmpty(text6))
						{
							text6 = text9;
						}
						else
						{
							text7 = text9;
						}
					}
				}
				if (!string.IsNullOrEmpty(text6) && !string.IsNullOrEmpty(text7))
				{
					flag2 = true;
				}
				else if (!string.IsNullOrEmpty(text6))
				{
					text4 = text6;
					text5 = ConvertToUnsigned(text4).ToUpperInvariant();
				}
			}
			else if (text4.Contains("   "))
			{
				string text10 = text4;
				while (text10.Contains("    "))
				{
					text10 = text10.Replace("    ", "   ");
				}
				text10 = text10.Replace("   ", "\t");
				string[] array3 = text10.Split('\t');
				string[] array4 = array3;
				foreach (string text11 in array4)
				{
					string text12 = text11.Trim();
					if (text12.Length > 0)
					{
						if (string.IsNullOrEmpty(text6))
						{
							text6 = text12;
						}
						else
						{
							text7 = text12;
						}
					}
				}
				if (!string.IsNullOrEmpty(text6) && !string.IsNullOrEmpty(text7))
				{
					flag2 = true;
				}
			}
			if (flag2 && !IsHeaderPiece(text7) && (!IsHeaderPiece(text6) || IsBareNumberPrefix(text6)))
			{
				flag2 = false;
				text4 = Regex.Replace(text4.Replace("\t", " "), "\\s{2,}", " ").Trim();
				text5 = ConvertToUnsigned(text4).ToUpperInvariant();
				text6 = "";
				text7 = "";
			}
			if (flag2)
			{
				num2 = i;
				list.Add(i);
				ConvertToUnsigned(text6).ToUpperInvariant();
				string text13 = ConvertToUnsigned(text7).ToUpperInvariant();
				if (string.IsNullOrEmpty(text2) && IsDocNumberLine(text6))
				{
					text2 = text6;
				}
				else
				{
					if (text.Length > 0)
					{
						text += "\r\n";
					}
					text += text6;
				}
				if (IsNationalOrMottoText(text13))
				{
					flag = true;
				}
				else if (text13.Contains("NGAY ") && text13.Contains(" THANG ") && text13.Contains(" NAM "))
				{
					text3 = text7;
				}
				continue;
			}
			if (IsNationalOrMottoText(text5) || (flag && Regex.IsMatch(text5, "^VIET\\s*NAM\\.?$")))
			{
				flag = true;
				num2 = i;
				list.Add(i);
				continue;
			}
			if (text5.Contains("NGAY ") && text5.Contains(" THANG ") && text5.Contains(" NAM "))
			{
				text3 = text4;
				num2 = i;
				list.Add(i);
				continue;
			}
			if (string.IsNullOrEmpty(text2) && IsDocNumberLine(text4))
			{
				text2 = text4;
				num2 = i;
				list.Add(i);
				continue;
			}
			string text14 = text4.Replace("-", "").Replace("_", "").Replace(" ", "")
				.Replace("*", "");
			if (text14.Length == 0 || IsSeparatorLine(text4))
			{
				if (num2 > 0)
				{
					num2 = i;
					list.Add(i);
				}
				continue;
			}
			if (list.Count >= 5 || !string.IsNullOrEmpty(text3) || (!RegexPatterns.AdminAgency.IsMatch(text4) && !RegexPatterns.PartyAgency.IsMatch(text4) && !RegexPatterns.MilitaryAgency.IsMatch(text4) && !RegexPatterns.YouthUnionAgency.IsMatch(text4) && (!string.IsNullOrEmpty(text2) || text4.Length >= 60 || text4.Contains(".") || !Regex.IsMatch(text4, "\\p{L}{2,}"))))
			{
				break;
			}
			if (text.Length > 0)
			{
				text += "\r\n";
			}
			text += text4;
			num2 = i;
			list.Add(i);
		}
		if (!flag && string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text3))
		{
			return false;
		}
		if (num2 == 0)
		{
			return false;
		}
		foreach (int item in list)
		{
			if (HasVisualContent(doc, doc.Paragraphs[item].Range))
			{
				return false;
			}
		}
		int pos = DeleteConsumedHeaderParagraphs(doc, list);
		pos = EnsureNotAdjacentToTable(doc, pos);
		object Start = pos;
		object End = pos;
		Range range = doc.Range(ref Start, ref End);
		Tables tables = doc.Tables;
		object DefaultTableBehavior = Missing.Value;
		object AutoFitBehavior = Missing.Value;
		Table table = tables.Add(range, 2, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
		table.Borders.Enable = 0;
		text2 = NormalizeDocNumber(text2, rulePack);
		text3 = NormalizeDocDate(text3);
		if (RulePackTraits.UsesOrganizationHeader(rulePack))
		{
			string agencySeparator = RulePackTraits.GetAgencySeparator(rulePack);
			if (string.IsNullOrEmpty(text))
			{
				text = (RulePackTraits.IsYouthUnion(rulePack) ? "TỈNH ĐOÀN...\r\nBCH ĐOÀN..." : "TỈNH ỦY...\r\nĐẢNG ỦY...");
			}
			if (!text.Contains("*"))
			{
				text = text + "\r\n" + agencySeparator;
			}
			table.Cell(1, 1).Range.Text = text;
			table.Cell(2, 1).Range.Text = text2;
			table.Cell(1, 2).Range.Text = RulePackTraits.GetOrganizationHeaderText(rulePack);
			table.Cell(2, 2).Range.Text = text3;
		}
		else
		{
			if (string.IsNullOrEmpty(text))
			{
				text = "UBND TỈNH...\r\nSỞ...";
			}
			table.Cell(1, 1).Range.Text = text;
			table.Cell(2, 1).Range.Text = text2;
			table.Cell(1, 2).Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
			table.Cell(2, 2).Range.Text = text3;
		}
		ResetInheritedFormatting(table);
		FormatExistingHeaderTable(table, rulePack, profile, skipRebuildCheck: true);
		return true;
	}

	internal static float MeasureLinePt(string text, float sizePt, bool bold, bool italic)
	{
		if (string.IsNullOrEmpty(text))
		{
			return 0f;
		}
		if (sizePt <= 0f || sizePt > 200f)
		{
			sizePt = 13f;
		}
		try
		{
			FontStyle fontStyle = FontStyle.Regular;
			if (bold)
			{
				fontStyle |= FontStyle.Bold;
			}
			if (italic)
			{
				fontStyle |= FontStyle.Italic;
			}
			using Bitmap image = new Bitmap(1, 1);
			using Graphics graphics = Graphics.FromImage(image);
			using System.Drawing.Font font = new System.Drawing.Font("Times New Roman", sizePt, fontStyle, GraphicsUnit.Point);
			graphics.PageUnit = GraphicsUnit.Point;
			return graphics.MeasureString(text, font, int.MaxValue, StringFormat.GenericTypographic).Width;
		}
		catch
		{
			return (float)text.Length * sizePt * 0.62f;
		}
	}

	internal static float MaxLineWidthPt(Cell cell)
	{
		float num = 0f;
		if (cell == null)
		{
			return num;
		}
		try
		{
			foreach (Paragraph paragraph in cell.Range.Paragraphs)
			{
				string text = paragraph.Range.Text ?? "";
				float sizePt = 13f;
				bool bold = true;
				bool italic = false;
				try
				{
					float size = paragraph.Range.Font.Size;
					if (size > 0f && size < 200f)
					{
						sizePt = size;
					}
				}
				catch
				{
				}
				try
				{
					int bold2 = paragraph.Range.Font.Bold;
					bold = bold2 != 0;
				}
				catch
				{
				}
				try
				{
					int italic2 = paragraph.Range.Font.Italic;
					italic = italic2 == -1;
				}
				catch
				{
				}
				bool flag = false;
				try
				{
					flag = paragraph.Range.Font.AllCaps == -1;
				}
				catch
				{
				}
				string[] array = text.Split(new char[5] { '\r', '\a', '\v', '\n', '\v' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string text2 in array)
				{
					string text3 = text2.Trim();
					if (text3.Length != 0 && !IsSeparatorLine(text3) && !IsAsteriskLine(text3))
					{
						if (flag)
						{
							text3 = text3.ToUpperInvariant();
						}
						float num2 = MeasureLinePt(text3, sizePt, bold, italic);
						if (num2 > num)
						{
							num = num2;
						}
					}
				}
			}
		}
		catch
		{
		}
		return num;
	}

	internal static void FitColumnsToContent(Table tbl, float printW, ref float col1Width, ref float col2Width)
	{
		try
		{
			if (tbl == null || tbl.Columns.Count < 2 || printW <= 0f)
			{
				return;
			}
			float num = MaxLineWidthPt(tbl.Cell(1, 1));
			float num2 = MaxLineWidthPt(tbl.Cell(1, 2));
			if (tbl.Rows.Count >= 2)
			{
				try
				{
					num = Math.Max(num, MaxLineWidthPt(tbl.Cell(2, 1)));
				}
				catch
				{
				}
				try
				{
					num2 = Math.Max(num2, MaxLineWidthPt(tbl.Cell(2, 2)));
				}
				catch
				{
				}
			}
			if (num <= 0f)
			{
				return;
			}
			num += 3f;
			num2 += 3f;
			float num3 = num2 * 1.04f;
			if (num3 > col2Width && num3 < printW - UnitConverter.MmToPt(50f))
			{
				float num4 = printW - num3 - num;
				float num5 = ((num4 >= 0f) ? (num3 + num4 / 2f) : num3);
				float num6 = printW - num5;
				if (num6 >= UnitConverter.MmToPt(50f))
				{
					col1Width = (float)Math.Floor(num6);
					col2Width = printW - col1Width;
					tbl.PreferredWidth = col1Width + col2Width;
					try
					{
						tbl.Columns[1].Width = col1Width;
					}
					catch
					{
					}
					try
					{
						tbl.Columns[2].Width = col2Width;
					}
					catch
					{
					}
					for (int i = 1; i <= tbl.Rows.Count; i++)
					{
						try
						{
							if (tbl.Rows[i].Cells.Count >= 2)
							{
								tbl.Rows[i].Cells[1].Width = col1Width;
								tbl.Rows[i].Cells[2].Width = col2Width;
							}
						}
						catch
						{
						}
					}
					return;
				}
			}
			if (num <= col1Width || printW - num < num2)
			{
				return;
			}
			float num7 = (float)Math.Ceiling(num);
			float num8 = printW - num7;
			if (num8 < num2)
			{
				return;
			}
			col1Width = num7;
			col2Width = num8;
			tbl.PreferredWidth = col1Width + col2Width;
			try
			{
				tbl.Columns[1].Width = col1Width;
			}
			catch
			{
			}
			try
			{
				tbl.Columns[2].Width = col2Width;
			}
			catch
			{
			}
			for (int j = 1; j <= tbl.Rows.Count; j++)
			{
				try
				{
					if (tbl.Rows[j].Cells.Count >= 2)
					{
						tbl.Rows[j].Cells[1].Width = col1Width;
						tbl.Rows[j].Cells[2].Width = col2Width;
					}
				}
				catch
				{
				}
			}
		}
		catch
		{
		}
	}

	private static void ApplyDatePlacementToHeaderTable(Table tbl, RulePackType rulePack, ProfileModel profile)
	{
		if (tbl == null || tbl.Rows.Count < 2 || tbl.Columns.Count < 2)
		{
			return;
		}
		bool dateInTopCell = IsHeaderDateInTopCell(profile);
		bool isOrgHeader = RulePackTraits.UsesOrganizationHeader(rulePack);
		string dateFromTop = "";
		for (int i = 1; i <= tbl.Cell(1, 2).Range.Paragraphs.Count; i++)
		{
			string pTxt = (tbl.Cell(1, 2).Range.Paragraphs[i].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
			if (!string.IsNullOrEmpty(pTxt))
			{
				string uTxt = ConvertToUnsigned(pTxt).ToUpperInvariant();
				if (uTxt.Contains("NGAY ") && (uTxt.Contains("THANG ") || uTxt.Contains("NAM ")))
				{
					dateFromTop = pTxt;
				}
			}
		}
		string dateFromBottom = (tbl.Cell(2, 2).Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
		string effectiveDate = (!string.IsNullOrEmpty(dateFromBottom)) ? NormalizeDocDate(dateFromBottom) : ((!string.IsNullOrEmpty(dateFromTop)) ? NormalizeDocDate(dateFromTop) : "");
		if (dateInTopCell)
		{
			if (isOrgHeader)
			{
				string orgHdr = RulePackTraits.GetOrganizationHeaderText(rulePack);
				tbl.Cell(1, 2).Range.Text = (!string.IsNullOrEmpty(effectiveDate)) ? (orgHdr + "\r\n" + effectiveDate) : orgHdr;
			}
			else
			{
				string natHdr = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
				tbl.Cell(1, 2).Range.Text = (!string.IsNullOrEmpty(effectiveDate)) ? (natHdr + "\r\n" + effectiveDate) : natHdr;
			}
			tbl.Cell(2, 2).Range.Text = "";
		}
		else
		{
			if (isOrgHeader)
			{
				tbl.Cell(1, 2).Range.Text = RulePackTraits.GetOrganizationHeaderText(rulePack);
			}
			else
			{
				tbl.Cell(1, 2).Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
			}
			if (!string.IsNullOrEmpty(effectiveDate))
			{
				tbl.Cell(2, 2).Range.Text = effectiveDate;
			}
		}
	}

	public static void FormatExistingHeaderTable(Table tbl, RulePackType rulePack, ProfileModel profile = null, bool skipRebuildCheck = false)
	{
		if (tbl == null)
		{
			return;
		}
		if (HasVisualContent(tbl.Range.Document, tbl.Range))
		{
			if (profile != null)
			{
				SemanticComponentFormatter.Apply(tbl.Range.Document, profile, tbl.Range, reflowTables: false);
			}
			return;
		}
		try
		{
			tbl.Range.Font.ColorIndex = WdColorIndex.wdAuto;
			tbl.Range.HighlightColorIndex = WdColorIndex.wdAuto;
			tbl.Range.Font.StrikeThrough = 0;
			tbl.Range.Font.DoubleStrikeThrough = 0;
		}
		catch
		{
		}
		try
		{
			Range range = tbl.Cell(1, 1).Range;
			string text = range.Text ?? "";
			for (int num = text.Length - 1; num >= 0; num--)
			{
				if (text[num] == '\v')
				{
					Document document = tbl.Range.Document;
					object Start = range.Start + num;
					object End = range.Start + num + 1;
					Range range2 = document.Range(ref Start, ref End);
					range2.Text = "\r";
				}
			}
			if (!skipRebuildCheck && NeedsRebuild(tbl, rulePack))
			{
				Document document2 = tbl.Range.Document;
				HeaderData headerData = ComponentDataExtractor.ExtractHeaderData(document2, tbl, rulePack);
				if (headerData != null && (headerData.HasExtractedAgency || !string.IsNullOrEmpty(headerData.DocNumber) || !string.IsNullOrEmpty(headerData.DateText)))
				{
					RebuildHeaderCleanly(document2, headerData, rulePack, tbl, profile);
					return;
				}
			}
			tbl.Borders.Enable = 0;
			WdBorderType[] array = new WdBorderType[6]
			{
				WdBorderType.wdBorderTop,
				WdBorderType.wdBorderBottom,
				WdBorderType.wdBorderLeft,
				WdBorderType.wdBorderRight,
				WdBorderType.wdBorderHorizontal,
				WdBorderType.wdBorderVertical
			};
			foreach (WdBorderType index in array)
			{
				try
				{
					tbl.Borders[index].LineStyle = WdLineStyle.wdLineStyleNone;
				}
				catch
				{
				}
			}
			tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
			tbl.Rows.WrapAroundText = 0;
			tbl.Rows.LeftIndent = 0f;
			tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft;
			tbl.Rows.AllowBreakAcrossPages = 0;
			tbl.Range.ParagraphFormat.KeepWithNext = -1;
			tbl.LeftPadding = 0f;
			tbl.RightPadding = 0f;
			tbl.TopPadding = 0f;
			tbl.BottomPadding = 0f;
			for (int j = 1; j <= tbl.Rows.Count; j++)
			{
				Row row = tbl.Rows[j];
				try
				{
					row.Borders.Enable = 0;
				}
				catch
				{
				}
				array = new WdBorderType[5]
				{
					WdBorderType.wdBorderTop,
					WdBorderType.wdBorderBottom,
					WdBorderType.wdBorderLeft,
					WdBorderType.wdBorderRight,
					WdBorderType.wdBorderHorizontal
				};
				foreach (WdBorderType index2 in array)
				{
					try
					{
						row.Borders[index2].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
				for (int k = 1; k <= row.Cells.Count; k++)
				{
					Cell cell = row.Cells[k];
					try
					{
						cell.Borders.Enable = 0;
					}
					catch
					{
					}
					array = new WdBorderType[6]
					{
						WdBorderType.wdBorderTop,
						WdBorderType.wdBorderBottom,
						WdBorderType.wdBorderLeft,
						WdBorderType.wdBorderRight,
						WdBorderType.wdBorderDiagonalDown,
						WdBorderType.wdBorderDiagonalUp
					};
					foreach (WdBorderType index3 in array)
					{
						try
						{
							cell.Borders[index3].LineStyle = WdLineStyle.wdLineStyleNone;
						}
						catch
						{
						}
					}
					cell.LeftPadding = 0f;
					cell.RightPadding = 0f;
					cell.TopPadding = 0f;
					cell.BottomPadding = 0f;
					foreach (Paragraph paragraph14 in cell.Range.Paragraphs)
					{
						try
						{
							paragraph14.Range.ParagraphFormat.Borders.Enable = 0;
						}
						catch
						{
						}
						array = new WdBorderType[4]
						{
							WdBorderType.wdBorderTop,
							WdBorderType.wdBorderBottom,
							WdBorderType.wdBorderLeft,
							WdBorderType.wdBorderRight
						};
						foreach (WdBorderType index4 in array)
						{
							try
							{
								paragraph14.Borders[index4].LineStyle = WdLineStyle.wdLineStyleNone;
							}
							catch
							{
							}
						}
					}
				}
			}
			bool flag = RulePackTraits.UsesOrganizationHeader(rulePack);
			float num2 = 0f;
			try
			{
				num2 = tbl.Range.Document.PageSetup.PageWidth - tbl.Range.Document.PageSetup.LeftMargin - tbl.Range.Document.PageSetup.RightMargin;
			}
			catch
			{
			}
			if (num2 < UnitConverter.MmToPt(140f) || num2 > UnitConverter.MmToPt(200f))
			{
				num2 = UnitConverter.MmToPt(165f);
			}
			if (tbl.Rows.Count == 1 && tbl.Columns.Count >= 2)
			{
				List<string> list = new List<string>();
				string text2 = "";
				string text3 = "";
				bool flag2 = false;
				for (int l = 1; l <= tbl.Cell(1, 1).Range.Paragraphs.Count; l++)
				{
					string text4 = (tbl.Cell(1, 1).Range.Paragraphs[l].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (string.IsNullOrEmpty(text4))
					{
						continue;
					}
					string text5 = ConvertToUnsigned(text4).ToUpperInvariant();
					if (text5.StartsWith("KINH GUI") || text5.StartsWith("KINH TRINH") || RegexPatterns.DirectRecipient.IsMatch(text4))
					{
						continue;
					}
					if (IsDocNumberLine(text4))
					{
						if (string.IsNullOrEmpty(text2))
						{
							text2 = text4;
						}
					}
					else if (RegexPatterns.DispatchSubject.IsMatch(text4) || text5.StartsWith("V/V ") || text5.StartsWith("V/V:") || text5.StartsWith("V/V.") || !string.IsNullOrEmpty(text3))
					{
						text3 = ((!string.IsNullOrEmpty(text3)) ? (text3 + " " + text4) : text4);
					}
					else if (IsAsteriskLine(text4))
					{
						flag2 = true;
					}
					else if (!IsSeparatorLine(text4))
					{
						list.Add(text4);
					}
				}
				List<string> list2 = new List<string>();
				string text6 = "";
				for (int m = 1; m <= tbl.Cell(1, 2).Range.Paragraphs.Count; m++)
				{
					string text7 = (tbl.Cell(1, 2).Range.Paragraphs[m].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (string.IsNullOrEmpty(text7))
					{
						continue;
					}
					string text8 = ConvertToUnsigned(text7).ToUpperInvariant();
					if (text8.Contains("NGAY ") && (text8.Contains("THANG ") || text8.Contains("NAM ")))
					{
						if (string.IsNullOrEmpty(text6))
						{
							text6 = text7;
						}
					}
					else
					{
						list2.Add(text7);
					}
				}
				if (!string.IsNullOrEmpty(text2) || !string.IsNullOrEmpty(text6))
				{
					try
					{
						Rows rows = tbl.Rows;
						object Start = Missing.Value;
						Row row2 = rows.Add(ref Start);
						if (flag)
						{
							if (!flag2 && list.Count > 0)
							{
								flag2 = true;
							}
							if (flag2)
							{
								list.Add(RulePackTraits.GetAgencySeparator(rulePack));
							}
							tbl.Cell(1, 1).Range.Text = string.Join("\r", list);
							if (!IsHeaderDateInTopCell(profile))
						{
							tbl.Cell(1, 2).Range.Text = RulePackTraits.GetOrganizationHeaderText(rulePack);
						}
						}
						else
						{
							if (list.Count > 0)
							{
								list.Add(UnderlineFormatter.CreateSolidSeparator());
							}
							tbl.Cell(1, 1).Range.Text = string.Join("\r", list);
							if (list2.Count == 0)
							{
								list2.Add("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
								list2.Add("Độc lập - Tự do - Hạnh phúc");
							}
							tbl.Cell(1, 2).Range.Text = string.Join("\r", list2);
						}
						string text9 = ((!string.IsNullOrEmpty(text2)) ? NormalizeDocNumber(text2, rulePack) : EmptyNumberPlaceholder(rulePack));
						if (!string.IsNullOrEmpty(text3))
						{
							text9 = text9 + "\r" + text3;
						}
						row2.Cells[1].Range.Text = text9;
						row2.Cells[2].Range.Text = ((!string.IsNullOrEmpty(text6)) ? NormalizeDocDate(text6) : "");
					}
					catch
					{
					}
				}
			}
			string text10 = "";
			if (tbl.Rows.Count >= 2 && tbl.Columns.Count >= 2)
			{
				try
				{
					List<string> list3 = new List<string>();
					for (int n = 1; n <= tbl.Cell(1, 1).Range.Paragraphs.Count; n++)
					{
						string text11 = (tbl.Cell(1, 1).Range.Paragraphs[n].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
						if (string.IsNullOrEmpty(text11))
						{
							continue;
						}
						string text12 = ConvertToUnsigned(text11).ToUpperInvariant();
						if (!text12.StartsWith("KINH GUI") && !text12.StartsWith("KINH TRINH") && !RegexPatterns.DirectRecipient.IsMatch(text11) && !IsDocNumberLine(text11))
						{
							if (RegexPatterns.DispatchSubject.IsMatch(text11) || text12.StartsWith("V/V ") || text12.StartsWith("V/V:") || text12.StartsWith("V/V.") || !string.IsNullOrEmpty(text10))
							{
								text10 = ((!string.IsNullOrEmpty(text10)) ? (text10 + " " + text11) : text11);
							}
							else if (!IsSeparatorLine(text11) && !IsAsteriskLine(text11))
							{
								list3.Add(text11);
							}
						}
					}
					if (list3.Count > 0)
					{
						if (flag)
						{
							string agencySeparator = RulePackTraits.GetAgencySeparator(rulePack);
							if (!list3.Contains(agencySeparator))
							{
								list3.Add(agencySeparator);
							}
						}
						else
						{
							list3.Add(UnderlineFormatter.CreateSolidSeparator());
						}
						tbl.Cell(1, 1).Range.Text = string.Join("\r", list3);
					}
					ApplyDatePlacementToHeaderTable(tbl, rulePack, profile);
				}
				catch
				{
				}
				try
				{
					if (flag)
					{
						if (!IsHeaderDateInTopCell(profile))
						{
							tbl.Cell(1, 2).Range.Text = RulePackTraits.GetOrganizationHeaderText(rulePack);
						}
					}
					else
					{
						List<string> list4 = new List<string>();
						string text13 = "";
						for (int num3 = 1; num3 <= tbl.Cell(1, 2).Range.Paragraphs.Count; num3++)
						{
							string text14 = (tbl.Cell(1, 2).Range.Paragraphs[num3].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
							if (!string.IsNullOrEmpty(text14))
							{
								string text15 = ConvertToUnsigned(text14).ToUpperInvariant();
								if (text15.Contains("NGAY ") && (text15.Contains("THANG ") || text15.Contains("NAM ")))
								{
									text13 = text14;
								}
								else
								{
									list4.Add(text14);
								}
							}
						}
						if (!string.IsNullOrEmpty(text13) && tbl.Rows.Count >= 2)
						{
							string value = (tbl.Cell(2, 2).Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
							if (string.IsNullOrEmpty(value))
							{
								tbl.Cell(2, 2).Range.Text = NormalizeDocDate(text13);
							}
						}
						if (list4.Count > 0 && !IsHeaderDateInTopCell(profile))
						{
							string input = string.Join("\r", list4);
							input = Regex.Replace(input, "(CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT|CONG HOA XA HOI CHU NGHIA VIET)\\s*[\\r\\n\\v]+\\s*NAM", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", RegexOptions.IgnoreCase);
							tbl.Cell(1, 2).Range.Text = input;
						}
					}
				}
				catch
				{
				}
			}
			string text16 = "";
			try
			{
				text16 = tbl.Cell(1, 1).Range.Text ?? "";
			}
			catch
			{
			}
			bool flag3 = false;
			string[] array2 = text16.Split(new char[3] { '\r', '\n', '\a' }, StringSplitOptions.RemoveEmptyEntries);
			string[] array3 = array2;
			foreach (string text17 in array3)
			{
				string text18 = text17.Trim();
				if (text18.Length >= 23 && !IsDocNumberLine(text18) && !IsSeparatorLine(text18) && !IsAsteriskLine(text18))
				{
					flag3 = true;
					break;
				}
			}
			float col2Width;
			float col1Width;
			if (flag)
			{
				float num4 = UnitConverter.MmToPt(80f);
				if (num2 < UnitConverter.MmToPt(160f))
				{
					num4 = UnitConverter.MmToPt(78f);
				}
				col2Width = num4;
				col1Width = num2 - col2Width;
				if (col1Width < UnitConverter.MmToPt(84f))
				{
					col1Width = UnitConverter.MmToPt(84f);
					col2Width = num2 - col1Width;
				}
			}
			else
			{
				float num5 = UnitConverter.MmToPt(106f);
				if (num2 >= UnitConverter.MmToPt(165f))
				{
					col2Width = num5;
					col1Width = num2 - col2Width;
					if (flag3 && col1Width < UnitConverter.MmToPt(60f))
					{
						col1Width = UnitConverter.MmToPt(60f);
						col2Width = num2 - col1Width;
					}
				}
				else if (num2 >= UnitConverter.MmToPt(158f))
				{
					col2Width = UnitConverter.MmToPt(105f);
					col1Width = num2 - col2Width;
				}
				else
				{
					col2Width = UnitConverter.MmToPt(104f);
					col1Width = num2 - col2Width;
				}
			}
			try
			{
				tbl.Rows.LeftIndent = 0f;
			}
			catch
			{
			}
			try
			{
				tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft;
			}
			catch
			{
			}
			tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
			tbl.PreferredWidth = col1Width + col2Width;
			try
			{
				tbl.Columns[1].Width = col1Width;
			}
			catch
			{
			}
			try
			{
				tbl.Columns[2].Width = col2Width;
			}
			catch
			{
			}
			tbl.LeftPadding = 0f;
			tbl.RightPadding = 0f;
			tbl.TopPadding = 0f;
			tbl.BottomPadding = 0f;
			try
			{
				tbl.AllowAutoFit = false;
				tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
			}
			catch
			{
			}
			for (int num6 = 1; num6 <= tbl.Rows.Count; num6++)
			{
				if (tbl.Rows[num6].Cells.Count >= 2)
				{
					try
					{
						tbl.Rows[num6].Cells[1].LeftPadding = 0f;
						tbl.Rows[num6].Cells[1].RightPadding = 0f;
						tbl.Rows[num6].Cells[1].TopPadding = 0f;
						tbl.Rows[num6].Cells[1].BottomPadding = 0f;
						tbl.Rows[num6].Cells[2].LeftPadding = 0f;
						tbl.Rows[num6].Cells[2].RightPadding = 0f;
						tbl.Rows[num6].Cells[2].TopPadding = 0f;
						tbl.Rows[num6].Cells[2].BottomPadding = 0f;
						tbl.Rows[num6].Cells[1].Width = col1Width;
						tbl.Rows[num6].Cells[2].Width = col2Width;
					}
					catch
					{
					}
				}
			}
			Cell cell2 = tbl.Cell(1, 1);
			cell2.Borders.Enable = 0;
			array = new WdBorderType[4]
			{
				WdBorderType.wdBorderTop,
				WdBorderType.wdBorderBottom,
				WdBorderType.wdBorderLeft,
				WdBorderType.wdBorderRight
			};
			foreach (WdBorderType index5 in array)
			{
				try
				{
					cell2.Borders[index5].LineStyle = WdLineStyle.wdLineStyleNone;
				}
				catch
				{
				}
			}
			cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			cell2.Range.ParagraphFormat.SpaceBefore = 0f;
			cell2.Range.ParagraphFormat.SpaceAfter = 0f;
			cell2.Range.ParagraphFormat.FirstLineIndent = 0f;
			cell2.Range.ParagraphFormat.LeftIndent = 0f;
			cell2.Range.ParagraphFormat.RightIndent = 0f;
			cell2.Range.Font.Name = "Times New Roman";
			int count = cell2.Range.Paragraphs.Count;
			int num7 = -1;
			int num8 = -1;
			int num9 = -1;
			List<int> list5 = new List<int>();
			List<string> list6 = new List<string>();
			string text19 = "";
			for (int num10 = 1; num10 <= count; num10++)
			{
				string text20 = (cell2.Range.Paragraphs[num10].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				if (IsDocNumberLine(text20))
				{
					num7 = num10;
					text19 = text20;
				}
				else if (IsAsteriskLine(text20))
				{
					num8 = num10;
				}
				else if (IsSeparatorLine(text20))
				{
					num9 = num10;
				}
				else if (text20.Length > 0)
				{
					list5.Add(num10);
					list6.Add(text20);
				}
			}
			if (!flag && num9 <= 0 && list6.Count > 0)
			{
				try
				{
					List<string> list7 = new List<string>();
					for (int num11 = 0; num11 < list6.Count; num11++)
					{
						list7.Add(list6[num11]);
					}
					list7.Add(UnderlineFormatter.CreateSolidSeparator());
					if (!string.IsNullOrEmpty(text19))
					{
						list7.Add(NormalizeDocNumber(text19, rulePack));
					}
					cell2.Range.Text = string.Join("\r", list7);
					count = cell2.Range.Paragraphs.Count;
					list5.Clear();
					for (int num12 = 1; num12 <= list6.Count; num12++)
					{
						list5.Add(num12);
					}
					num9 = list6.Count + 1;
					num7 = ((!string.IsNullOrEmpty(text19)) ? (list6.Count + 2) : (-1));
				}
				catch
				{
				}
			}
			else if (flag && num8 <= 0 && list6.Count > 0)
			{
				try
				{
					List<string> list8 = new List<string>();
					for (int num13 = 0; num13 < list6.Count; num13++)
					{
						list8.Add(list6[num13]);
					}
					list8.Add(RulePackTraits.GetAgencySeparator(rulePack));
					if (!string.IsNullOrEmpty(text19))
					{
						list8.Add(NormalizeDocNumber(text19, rulePack));
					}
					cell2.Range.Text = string.Join("\r", list8);
					count = cell2.Range.Paragraphs.Count;
					list5.Clear();
					for (int num14 = 1; num14 <= list6.Count; num14++)
					{
						list5.Add(num14);
					}
					num8 = list6.Count + 1;
					num7 = ((!string.IsNullOrEmpty(text19)) ? (list6.Count + 2) : (-1));
				}
				catch
				{
				}
			}
			HashSet<int> hashSet = IdentifyIssuingAgencyParaIndices(cell2, list5);
			if (flag)
			{
				for (int num15 = 1; num15 <= cell2.Range.Paragraphs.Count; num15++)
				{
					Paragraph paragraph2 = cell2.Range.Paragraphs[num15];
					string text21 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					paragraph2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph2.Format.FirstLineIndent = 0f;
					paragraph2.Format.LeftIndent = 0f;
					paragraph2.Format.RightIndent = 0f;
					paragraph2.Format.SpaceBefore = 0f;
					paragraph2.Format.SpaceAfter = 0f;
					paragraph2.Range.Font.Name = "Times New Roman";
					paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph2.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					if (num15 == num7)
					{
						paragraph2.Range.Text = NormalizeDocNumber(text21, rulePack) + "\r";
						paragraph2.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.DocumentNumber).SizePt ?? 14f;
						paragraph2.Range.Font.Bold = 0;
						paragraph2.Range.Font.Italic = 0;
					}
					else if (num15 == num8 || IsAsteriskLine(text21))
					{
						string agencySeparator2 = RulePackTraits.GetAgencySeparator(rulePack);
						if (text21 != agencySeparator2)
						{
							paragraph2.Range.Text = agencySeparator2 + "\r";
						}
						paragraph2.Range.Font.Size = 14f;
						paragraph2.Range.Font.Bold = 0;
						paragraph2.Range.Font.Italic = 0;
					}
					else if (hashSet.Contains(num15))
					{
						paragraph2.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.IssuingAgency).SizePt ?? 14f;
						paragraph2.Range.Font.Bold = -1;
						paragraph2.Range.Font.Italic = 0;
					}
					else if (list5.Contains(num15))
					{
						paragraph2.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.SuperiorAgency).SizePt ?? 14f;
						paragraph2.Range.Font.Bold = 0;
						paragraph2.Range.Font.Italic = 0;
					}
					else if (text21.Length == 0)
					{
						try
						{
							paragraph2.Range.Text = "";
						}
						catch
						{
						}
					}
				}
			}
			else
			{
				float size = profile?.GetComponentStyle(SemanticRole.IssuingAgency).SizePt ?? ((rulePack == RulePackType.TT68_2023_BQP) ? 13f : (flag3 ? 12f : 13f));
				for (int num16 = 1; num16 <= cell2.Range.Paragraphs.Count; num16++)
				{
					Paragraph paragraph3 = cell2.Range.Paragraphs[num16];
					string text22 = (paragraph3.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					string text23 = text22.Replace("-", "").Replace("—", "").Replace("_", "")
						.Replace("*", "")
						.Replace(" ", "")
						.Replace("–", "")
						.Replace("–", "")
						.Replace("―", "")
						.Replace("─", "")
						.Replace("―", "")
						.Replace("─", "");
					if (text23.Length == 0 && num16 != num7 && !hashSet.Contains(num16) && num16 != num9 && !IsSeparatorLine(text22))
					{
						try
						{
							paragraph3.Range.Text = "";
						}
						catch
						{
						}
						continue;
					}
					paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph3.Format.FirstLineIndent = 0f;
					paragraph3.Format.LeftIndent = 0f;
					paragraph3.Format.RightIndent = 0f;
					paragraph3.Format.SpaceBefore = 0f;
					paragraph3.Format.SpaceAfter = 0f;
					paragraph3.Range.Font.Name = "Times New Roman";
					paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph3.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					if (num16 == num7)
					{
						paragraph3.Range.Text = NormalizeDocNumber(text22, rulePack) + "\r";
						paragraph3.Range.Font.Size = 13f;
						paragraph3.Range.Font.Bold = 0;
						paragraph3.Range.Font.Italic = 0;
					}
					else if (num16 == num9 || IsSeparatorLine(text22))
					{
						paragraph3.Range.Text = UnderlineFormatter.CreateSolidSeparator() + "\r";
						paragraph3.Range.Font.Size = 10f;
						paragraph3.Range.Font.Bold = -1;
						paragraph3.Range.Font.Italic = 0;
					}
					else if (hashSet.Contains(num16))
					{
						paragraph3.Range.Font.Size = size;
						paragraph3.Range.Font.Bold = -1;
						paragraph3.Range.Font.Italic = 0;
					}
					else if (list5.Contains(num16))
					{
						if (profile != null)
						{
							paragraph3.Range.Font.Size = profile.GetComponentStyle(SemanticRole.SuperiorAgency).SizePt;
						}
						else if (rulePack != RulePackType.TT68_2023_BQP && (text22.Length >= 25 || flag3))
						{
							paragraph3.Range.Font.Size = 12f;
						}
						else
						{
							paragraph3.Range.Font.Size = 13f;
						}
						paragraph3.Range.Font.Bold = 0;
						paragraph3.Range.Font.Italic = 0;
					}
				}
			}
			Cell cell3 = tbl.Cell(1, 2);
			cell3.Borders.Enable = 0;
			array = new WdBorderType[4]
			{
				WdBorderType.wdBorderTop,
				WdBorderType.wdBorderBottom,
				WdBorderType.wdBorderLeft,
				WdBorderType.wdBorderRight
			};
			foreach (WdBorderType index6 in array)
			{
				try
				{
					cell3.Borders[index6].LineStyle = WdLineStyle.wdLineStyleNone;
				}
				catch
				{
				}
			}
			cell3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell3.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			cell3.Range.ParagraphFormat.SpaceBefore = 0f;
			cell3.Range.ParagraphFormat.SpaceAfter = 0f;
			cell3.Range.ParagraphFormat.FirstLineIndent = 0f;
			cell3.Range.ParagraphFormat.LeftIndent = 0f;
			cell3.Range.ParagraphFormat.RightIndent = 0f;
			cell3.Range.Font.Name = "Times New Roman";
			string text24 = (cell3.Range.Text ?? "").Replace("\a", "");
			if (text24.Contains("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT") || text24.Contains("CONG HOA XA HOI CHU NGHIA VIET"))
			{
				string text25 = Regex.Replace(text24, "(CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT|CONG HOA XA HOI CHU NGHIA VIET)\\s*[\\r\\n\\v]+\\s*NAM", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", RegexOptions.IgnoreCase);
				if (text25 != text24)
				{
					cell3.Range.Text = text25;
				}
			}
			int count2 = cell3.Range.Paragraphs.Count;
			int num17 = -1;
			for (int num18 = 1; num18 <= count2; num18++)
			{
				string text26 = (cell3.Range.Paragraphs[num18].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				string text27 = ConvertToUnsigned(text26).ToUpperInvariant();
				if (text27.Contains("NGAY ") && (text27.Contains("THANG ") || text27.Contains("NAM ")))
				{
					num17 = num18;
					break;
				}
			}
			if (flag)
			{
				if (tbl.Rows.Count >= 2 && !IsHeaderDateInTopCell(profile))
				{
					cell3.Range.Text = RulePackTraits.GetOrganizationHeaderText(rulePack);
					cell3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					cell3.Range.ParagraphFormat.FirstLineIndent = 0f;
					cell3.Range.ParagraphFormat.LeftIndent = 0f;
					cell3.Range.ParagraphFormat.RightIndent = 0f;
					cell3.Range.ParagraphFormat.SpaceBefore = 0f;
					cell3.Range.ParagraphFormat.SpaceAfter = 0f;
					cell3.Range.Font.Name = "Times New Roman";
					cell3.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.PartyHeader).SizePt ?? 15f;
					cell3.Range.Font.Bold = -1;
					cell3.Range.Font.Italic = 0;
					cell3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				}
				else
				{
					for (int num19 = 1; num19 <= cell3.Range.Paragraphs.Count; num19++)
					{
						Paragraph paragraph4 = cell3.Range.Paragraphs[num19];
						string text28 = (paragraph4.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
						string text29 = ConvertToUnsigned(text28).ToUpperInvariant();
						paragraph4.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
						paragraph4.Format.FirstLineIndent = 0f;
						paragraph4.Format.LeftIndent = 0f;
						paragraph4.Format.RightIndent = 0f;
						paragraph4.Format.SpaceBefore = 0f;
						paragraph4.Format.SpaceAfter = 0f;
						paragraph4.Range.Font.Name = "Times New Roman";
						if (num19 == num17 || (text29.Contains("NGAY ") && (text29.Contains("THANG ") || text29.Contains("NAM "))))
						{
							if (false && !text28.Contains("ngày"))
							{
								paragraph4.Range.Text = NormalizeDocDate(text28);
							}
							paragraph4.Format.SpaceBefore = 8f;
							paragraph4.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.DocumentDate).SizePt ?? 14f;
							paragraph4.Range.Font.Bold = 0;
							paragraph4.Range.Font.Italic = -1;
							paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						}
						else if (num19 == 1 || ContainsOrganizationTitle(text29))
						{
							string organizationHeaderText = RulePackTraits.GetOrganizationHeaderText(rulePack);
							if (false && text28 != organizationHeaderText)
							{
								paragraph4.Range.Text = organizationHeaderText;
							}
							paragraph4.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.PartyHeader).SizePt ?? 15f;
							paragraph4.Range.Font.Bold = -1;
							paragraph4.Range.Font.Italic = 0;
							paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						}
						else if (text28.Length == 0)
						{
							paragraph4.Range.Font.Size = 6f;
							paragraph4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						}
					}
				}
			}
			else if (tbl.Rows.Count >= 2 && !IsHeaderDateInTopCell(profile))
			{
				cell3.Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
				if (cell3.Range.Paragraphs.Count >= 1)
				{
					Paragraph paragraph5 = cell3.Range.Paragraphs[1];
					paragraph5.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph5.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					paragraph5.Format.FirstLineIndent = 0f;
					paragraph5.Format.LeftIndent = 0f;
					paragraph5.Format.RightIndent = 0f;
					paragraph5.Format.SpaceBefore = 0f;
					paragraph5.Format.SpaceAfter = 0f;
					paragraph5.Range.Font.Name = "Times New Roman";
					paragraph5.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.NationalHeader).SizePt ?? ((rulePack == RulePackType.TT68_2023_BQP) ? 13f : 12f);
					paragraph5.Range.Font.Bold = -1;
					paragraph5.Range.Font.Italic = 0;
					paragraph5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					if (col2Width < UnitConverter.MmToPt(105f))
					{
						try
						{
							paragraph5.Range.Font.Scaling = 97;
						}
						catch
						{
						}
					}
					else
					{
						try
						{
							paragraph5.Range.Font.Scaling = 100;
						}
						catch
						{
						}
					}
					try
					{
						paragraph5.Range.ParagraphFormat.Borders.Enable = 0;
					}
					catch
					{
					}
					try
					{
						paragraph5.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					try
					{
						paragraph5.Range.NoProofing = -1;
					}
					catch
					{
					}
				}
				if (cell3.Range.Paragraphs.Count >= 2)
				{
					Paragraph paragraph6 = cell3.Range.Paragraphs[2];
					paragraph6.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph6.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					paragraph6.Format.FirstLineIndent = 0f;
					paragraph6.Format.LeftIndent = 0f;
					paragraph6.Format.RightIndent = 0f;
					paragraph6.Format.SpaceBefore = 0f;
					paragraph6.Format.SpaceAfter = 0f;
					paragraph6.Range.Font.Name = "Times New Roman";
					paragraph6.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.Motto).SizePt ?? 13f;
					paragraph6.Range.Font.Bold = -1;
					paragraph6.Range.Font.Italic = 0;
					paragraph6.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph6.Range.ParagraphFormat.Borders.Enable = 0;
					}
					catch
					{
					}
					try
					{
						paragraph6.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					try
					{
						paragraph6.Range.NoProofing = -1;
					}
					catch
					{
					}
				}
			}
			else
			{
				string text30 = "";
				if (num17 > 0 && num17 <= cell3.Range.Paragraphs.Count)
				{
					string text31 = (cell3.Range.Paragraphs[num17].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (!string.IsNullOrEmpty(text31))
					{
						text30 = NormalizeDocDate(text31);
					}
				}
				if (!string.IsNullOrEmpty(text30))
				{
					cell3.Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc\r\n" + text30;
				}
				else
				{
					cell3.Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
				}
				if (cell3.Range.Paragraphs.Count >= 1)
				{
					Paragraph paragraph7 = cell3.Range.Paragraphs[1];
					paragraph7.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph7.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					paragraph7.Format.FirstLineIndent = 0f;
					paragraph7.Format.LeftIndent = 0f;
					paragraph7.Format.RightIndent = 0f;
					paragraph7.Format.SpaceBefore = 0f;
					paragraph7.Format.SpaceAfter = 0f;
					paragraph7.Range.Font.Name = "Times New Roman";
					paragraph7.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.NationalHeader).SizePt ?? ((rulePack == RulePackType.TT68_2023_BQP) ? 13f : 12f);
					paragraph7.Range.Font.Bold = -1;
					paragraph7.Range.Font.Italic = 0;
					paragraph7.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					if (col2Width < UnitConverter.MmToPt(105f))
					{
						try
						{
							paragraph7.Range.Font.Scaling = 97;
						}
						catch
						{
						}
					}
					else
					{
						try
						{
							paragraph7.Range.Font.Scaling = 100;
						}
						catch
						{
						}
					}
					try
					{
						paragraph7.Range.ParagraphFormat.Borders.Enable = 0;
					}
					catch
					{
					}
					try
					{
						paragraph7.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					try
					{
						paragraph7.Range.NoProofing = -1;
					}
					catch
					{
					}
				}
				if (cell3.Range.Paragraphs.Count >= 2)
				{
					Paragraph paragraph8 = cell3.Range.Paragraphs[2];
					paragraph8.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph8.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					paragraph8.Format.FirstLineIndent = 0f;
					paragraph8.Format.LeftIndent = 0f;
					paragraph8.Format.RightIndent = 0f;
					paragraph8.Format.SpaceBefore = 0f;
					paragraph8.Format.SpaceAfter = 0f;
					paragraph8.Range.Font.Name = "Times New Roman";
					paragraph8.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.Motto).SizePt ?? 13f;
					paragraph8.Range.Font.Bold = -1;
					paragraph8.Range.Font.Italic = 0;
					paragraph8.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph8.Range.ParagraphFormat.Borders.Enable = 0;
					}
					catch
					{
					}
					try
					{
						paragraph8.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					try
					{
						paragraph8.Range.NoProofing = -1;
					}
					catch
					{
					}
				}
				if (cell3.Range.Paragraphs.Count >= 3)
				{
					Paragraph paragraph9 = cell3.Range.Paragraphs[3];
					paragraph9.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph9.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					paragraph9.Format.FirstLineIndent = 0f;
					paragraph9.Format.LeftIndent = 0f;
					paragraph9.Format.RightIndent = 0f;
					paragraph9.Format.SpaceBefore = 8f;
					paragraph9.Format.SpaceAfter = 0f;
					paragraph9.Range.Font.Name = "Times New Roman";
					paragraph9.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.DocumentDate).SizePt ?? ((rulePack == RulePackType.TT68_2023_BQP) ? 13f : 14f);
					paragraph9.Range.Font.Bold = 0;
					paragraph9.Range.Font.Italic = -1;
					paragraph9.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					try
					{
						paragraph9.Range.ParagraphFormat.Borders.Enable = 0;
					}
					catch
					{
					}
					try
					{
						paragraph9.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
					try
					{
						paragraph9.Range.NoProofing = -1;
					}
					catch
					{
					}
				}
			}
			if (tbl.Rows.Count >= 2)
			{
				Cell cell4 = tbl.Cell(2, 1);
				cell4.Borders.Enable = 0;
				array = new WdBorderType[4]
				{
					WdBorderType.wdBorderTop,
					WdBorderType.wdBorderBottom,
					WdBorderType.wdBorderLeft,
					WdBorderType.wdBorderRight
				};
				foreach (WdBorderType index7 in array)
				{
					try
					{
						cell4.Borders[index7].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
				cell4.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				cell4.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				cell4.Range.ParagraphFormat.SpaceBefore = 0f;
				cell4.Range.ParagraphFormat.SpaceAfter = 0f;
				cell4.Range.ParagraphFormat.FirstLineIndent = 0f;
				cell4.Range.ParagraphFormat.LeftIndent = 0f;
				cell4.Range.ParagraphFormat.RightIndent = 0f;
				cell4.Range.Font.Name = "Times New Roman";
				string text32 = "";
				string text33 = "";
				for (int num20 = 1; num20 <= cell4.Range.Paragraphs.Count; num20++)
				{
					string text34 = (cell4.Range.Paragraphs[num20].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (string.IsNullOrEmpty(text34))
					{
						continue;
					}
					string text35 = ConvertToUnsigned(text34).ToUpperInvariant();
					if (text35.StartsWith("KINH GUI") || text35.StartsWith("KINH TRINH") || RegexPatterns.DirectRecipient.IsMatch(text34))
					{
						continue;
					}
					if (IsDocNumberLine(text34) || (string.IsNullOrEmpty(text32) && (text35.StartsWith("SO") || text35.Contains("/")) && !text35.StartsWith("V/V")))
					{
						text32 = text34;
					}
					else if (RegexPatterns.DispatchSubject.IsMatch(text34) || text35.StartsWith("V/V ") || text35.StartsWith("V/V:") || text35.StartsWith("V/V.") || !string.IsNullOrEmpty(text33))
					{
						text33 = ((!string.IsNullOrEmpty(text33)) ? (text33 + " " + text34) : text34);
					}
					else if (!RegexPatterns.PartyAgency.IsMatch(text34) && !RegexPatterns.AdminAgency.IsMatch(text34) && !RegexPatterns.MilitaryAgency.IsMatch(text34) && !RegexPatterns.YouthUnionAgency.IsMatch(text34) && !text35.StartsWith("BCH ") && !text35.StartsWith("CHI BO") && !text35.StartsWith("DANG BO") && !text35.StartsWith("DANG UY") && !text35.StartsWith("CHI UY") && !text35.StartsWith("BAN ") && !text35.StartsWith("SO ") && !text35.StartsWith("PHONG ") && !text35.StartsWith("UY BAN") && !text35.StartsWith("HOI DONG"))
					{
						if (string.IsNullOrEmpty(text32))
						{
							text32 = text34;
						}
						else if (string.IsNullOrEmpty(text33))
						{
							text33 = text34;
						}
					}
				}
				if (!string.IsNullOrEmpty(text10) && string.IsNullOrEmpty(text33))
				{
					text33 = text10;
				}
				string text36 = ((!string.IsNullOrEmpty(text32)) ? NormalizeDocNumber(text32, rulePack) : EmptyNumberPlaceholder(rulePack));
				if (!string.IsNullOrEmpty(text33))
				{
					cell4.Range.Text = text36 + "\r" + text33;
					if (cell4.Range.Paragraphs.Count >= 2)
					{
						ComponentStyle componentStyle = profile?.GetComponentStyle(SemanticRole.DocumentNumber);
						Paragraph paragraph10 = cell4.Range.Paragraphs[1];
						paragraph10.Range.Font.Name = "Times New Roman";
						paragraph10.Range.Font.Size = componentStyle?.SizePt ?? (flag ? 14f : 13f);
						paragraph10.Range.Font.Bold = ((componentStyle != null) ? (componentStyle.Bold ? (-1) : 0) : 0);
						paragraph10.Range.Font.Italic = ((componentStyle != null) ? (componentStyle.Italic ? (-1) : 0) : 0);
						paragraph10.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						paragraph10.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
						paragraph10.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
						paragraph10.Format.SpaceBefore = 0f;
						paragraph10.Format.SpaceAfter = 0f;
						ComponentStyle componentStyle2 = profile?.GetComponentStyle(SemanticRole.DispatchSubject);
						Paragraph paragraph11 = cell4.Range.Paragraphs[2];
						paragraph11.Range.Font.Name = "Times New Roman";
						paragraph11.Range.Font.Size = componentStyle2?.SizePt ?? 12f;
						paragraph11.Range.Font.Bold = ((componentStyle2 != null) ? (componentStyle2.Bold ? (-1) : 0) : 0);
						paragraph11.Range.Font.Italic = ((componentStyle2 == null) ? (flag ? (-1) : 0) : (componentStyle2.Italic ? (-1) : 0));
						paragraph11.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						paragraph11.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
						paragraph11.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
						paragraph11.Format.SpaceBefore = 2f;
						paragraph11.Format.SpaceAfter = 0f;
					}
				}
				else
				{
					ComponentStyle componentStyle3 = profile?.GetComponentStyle(SemanticRole.DocumentNumber);
					cell4.Range.Text = text36;
					cell4.Range.Font.Name = "Times New Roman";
					cell4.Range.Font.Size = componentStyle3?.SizePt ?? (flag ? 14f : 13f);
					cell4.Range.Font.Bold = ((componentStyle3 != null) ? (componentStyle3.Bold ? (-1) : 0) : 0);
					cell4.Range.Font.Italic = ((componentStyle3 != null) ? (componentStyle3.Italic ? (-1) : 0) : 0);
					cell4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				}
				foreach (Paragraph paragraph15 in cell4.Range.Paragraphs)
				{
					paragraph15.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					paragraph15.Range.ParagraphFormat.Borders.Enable = 0;
					try
					{
						paragraph15.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
				Cell cell5 = tbl.Cell(2, 2);
				cell5.Borders.Enable = 0;
				array = new WdBorderType[4]
				{
					WdBorderType.wdBorderTop,
					WdBorderType.wdBorderBottom,
					WdBorderType.wdBorderLeft,
					WdBorderType.wdBorderRight
				};
				foreach (WdBorderType index8 in array)
				{
					try
					{
						cell5.Borders[index8].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
				cell5.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				cell5.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				cell5.Range.ParagraphFormat.SpaceBefore = 0f;
				cell5.Range.ParagraphFormat.SpaceAfter = 0f;
				cell5.Range.ParagraphFormat.FirstLineIndent = 0f;
				cell5.Range.ParagraphFormat.LeftIndent = 0f;
				cell5.Range.ParagraphFormat.RightIndent = 0f;
				cell5.Range.Font.Name = "Times New Roman";
				cell5.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.DocumentDate).SizePt ?? ((rulePack == RulePackType.TT68_2023_BQP) ? 13f : 14f);
				cell5.Range.Font.Bold = 0;
				cell5.Range.Font.Italic = -1;
				cell5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				if (IsHeaderDateInTopCell(profile))
				{
					cell5.Range.Text = "";
				}
				else
				{
					string text37 = (cell5.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (!string.IsNullOrEmpty(text37))
					{
						cell5.Range.Text = NormalizeDocDate(text37);
						cell5.Range.Font.Size = profile?.GetComponentStyle(SemanticRole.DocumentDate).SizePt ?? ((rulePack == RulePackType.TT68_2023_BQP) ? 13f : 14f);
						cell5.Range.Font.Bold = 0;
						cell5.Range.Font.Italic = -1;
						cell5.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					}
				}
				foreach (Paragraph paragraph16 in cell5.Range.Paragraphs)
				{
					paragraph16.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					paragraph16.Range.ParagraphFormat.Borders.Enable = 0;
					try
					{
						paragraph16.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
					}
					catch
					{
					}
				}
			}
			ApplyAgencyAllCaps(tbl);
			FitColumnsToContent(tbl, num2, ref col1Width, ref col2Width);
			UnderlineFormatter.EnsureHeaderShapeLines(tbl, rulePack, col1Width, col2Width);
		}
		catch
		{
		}
	}

	public static Table RebuildHeaderCleanly(Document doc, HeaderData data, RulePackType rulePack, Table existingTable = null, ProfileModel profile = null)
	{
		if (doc == null)
		{
			return null;
		}
		if (data == null)
		{
			data = ComponentDataExtractor.ExtractHeaderData(doc, existingTable, rulePack);
		}
		int pos = 0;
		if (existingTable != null)
		{
			if (HasVisualContent(doc, existingTable.Range))
			{
				FormatExistingHeaderTable(existingTable, rulePack, profile);
				return existingTable;
			}
			try
			{
				pos = existingTable.Range.Start;
				existingTable.Delete();
			}
			catch
			{
			}
			try
			{
				int num = Math.Min(doc.Paragraphs.Count, 8);
				for (int i = 1; i <= num && i <= doc.Paragraphs.Count; i++)
				{
					Paragraph paragraph = doc.Paragraphs[i];
					if (paragraph.Range.Tables.Count != 0)
					{
						continue;
					}
					string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
						.Trim();
					if (!string.IsNullOrEmpty(text))
					{
						if (IsDocumentTitle(text))
						{
							break;
						}
						string text2 = ConvertToUnsigned(text).ToUpperInvariant();
						bool flag = !string.IsNullOrEmpty(data.DateText) && (text == data.DateText.Trim() || (text2.Contains("NGAY ") && (text2.Contains("THANG ") || text2.Contains("NAM "))));
						bool flag2 = !string.IsNullOrEmpty(data.DocNumber) && (text == data.DocNumber.Trim() || IsDocNumberLine(text));
						if (flag || flag2)
						{
							Range range = paragraph.Range;
							object Unit = Missing.Value;
							object Count = Missing.Value;
							range.Delete(ref Unit, ref Count);
							i--;
							num--;
						}
					}
				}
			}
			catch
			{
			}
		}
		else
		{
			int num2 = Math.Min(doc.Paragraphs.Count, 12);
			int num3 = Math.Min(doc.Paragraphs.Count, 80);
			int num4 = 0;
			List<int> list = new List<int>();
			for (int j = 1; j <= num2; j++)
			{
				Paragraph paragraph2 = doc.Paragraphs[j];
				if (paragraph2.Range.Tables.Count > 0)
				{
					if (num2 < num3)
					{
						num2++;
					}
					continue;
				}
				string text3 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (text3.Length > 0)
				{
					string text4 = ConvertToUnsigned(text3).ToUpperInvariant();
					if (IsDocumentTitle(text4))
					{
						break;
					}
					bool flag3 = text4.Contains("CONG HOA XA HOI") || ContainsOrganizationTitle(text4) || text4.Contains("DOC LAP - TU DO") || IsDocNumberLine(text3) || (text4.Contains("NGAY ") && (text4.Contains("THANG ") || text4.Contains("NAM "))) || IsAsteriskLine(text3);
					bool flag4 = list.Count < 5 && (text3.Contains("\t") || text3.Contains("   ") || RegexPatterns.AdminAgency.IsMatch(text3) || RegexPatterns.PartyAgency.IsMatch(text3) || RegexPatterns.YouthUnionAgency.IsMatch(text3));
					bool flag5 = num4 > 0 && IsSeparatorLine(text3);
					if (flag3 || flag4 || flag5)
					{
						num4 = j;
						list.Add(j);
					}
				}
			}
			if (list.Count > 0)
			{
				foreach (int item in list)
				{
					if (HasVisualContent(doc, doc.Paragraphs[item].Range))
					{
						return null;
					}
				}
				try
				{
					pos = DeleteConsumedHeaderParagraphs(doc, list);
				}
				catch
				{
				}
			}
		}
		int num5 = EnsureNotAdjacentToTable(doc, pos);
		object Start = num5;
		object End = num5;
		Range range2 = doc.Range(ref Start, ref End);
		Tables tables = doc.Tables;
		object DefaultTableBehavior = Missing.Value;
		object AutoFitBehavior = Missing.Value;
		Table table = tables.Add(range2, 2, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
		table.Borders.Enable = 0;
		bool flag6 = RulePackTraits.UsesOrganizationHeader(rulePack);
		string text5 = "";
		if (!string.IsNullOrEmpty(data.SuperiorAgency) && !string.IsNullOrEmpty(data.IssuingAgency))
		{
			text5 = data.SuperiorAgency + "\r\n" + data.IssuingAgency;
		}
		else if (string.IsNullOrEmpty(data.IssuingAgency))
		{
			if (string.IsNullOrEmpty(data.SuperiorAgency))
			{
				if (RulePackTraits.IsYouthUnion(rulePack))
				{
					text5 = "TỈNH ĐOÀN...\r\nBCH ĐOÀN...";
				}
				else
				{
					text5 = (flag6 ? "TỈNH ỦY BÌNH DƯƠNG\r\nĐẢNG ỦY KHỐI CƠ QUAN" : "ỦY BAN NHÂN DÂN\r\nTỈNH BÌNH DƯƠNG");
				}
			}
			else
			{
				text5 = data.SuperiorAgency;
			}
		}
		else
		{
			text5 = data.IssuingAgency;
		}
		text5 = ((!flag6) ? (text5 + "\r\n" + UnderlineFormatter.CreateSolidSeparator()) : (text5 + "\r\n" + RulePackTraits.GetAgencySeparator(rulePack)));
		table.Cell(1, 1).Range.Text = text5;
		string text6 = ((!string.IsNullOrEmpty(data.DocNumber)) ? NormalizeDocNumber(data.DocNumber, rulePack) : EmptyNumberPlaceholder(rulePack));
		if (!string.IsNullOrEmpty(data.DispatchSubject))
		{
			text6 = text6 + "\r\n" + data.DispatchSubject;
		}
		table.Cell(2, 1).Range.Text = text6;
		if (flag6)
		{
			table.Cell(1, 2).Range.Text = RulePackTraits.GetOrganizationHeaderText(rulePack);
		}
		else
		{
			table.Cell(1, 2).Range.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc";
		}
		table.Cell(2, 2).Range.Text = ((!string.IsNullOrEmpty(data.DateText)) ? NormalizeDocDate(data.DateText) : "..., ngày ... tháng ... năm 20...");
		ResetInheritedFormatting(table);
		FormatExistingHeaderTable(table, rulePack, profile, skipRebuildCheck: true);
		try
		{
			Range range3 = table.Range;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			range3.Collapse(ref Direction);
			if (doc.Paragraphs.Count == 0 || range3.End >= doc.Content.End - 2)
			{
				range3.InsertParagraphAfter();
			}
		}
		catch
		{
		}
		return table;
	}

	public static void AutoInsertMissingHeader(Document doc, RulePackType rulePack, ProfileModel profile = null)
	{
		if (doc != null)
		{
			RebuildHeaderCleanly(doc, null, rulePack, null, profile);
		}
	}
}
