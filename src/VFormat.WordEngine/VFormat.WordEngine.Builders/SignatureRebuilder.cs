using System;
using System.Collections.Generic;
using System.Globalization;
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

public static class SignatureRebuilder
{
	private static readonly CultureInfo ViCulture = new CultureInfo("vi-VN");

	private static readonly HashSet<string> RecipientAcronyms = new HashSet<string>(new string[62]
	{
		"UBND", "HĐND", "MTTQ", "UBMTTQ", "UBKT", "VPUB", "VPCP", "TTCP", "TTG", "BCH",
		"BTV", "BTC", "TƯ", "TW", "QH", "ĐBQH", "TAND", "VKSND", "CA", "BCA",
		"BQP", "BTP", "STP", "ĐU", "ĐUQS", "LĐLĐ", "CĐ", "ĐTN", "HLHPN", "HND",
		"HCCB", "CCB", "CNTT", "GDĐT", "LĐTBXH", "TNMT", "PTNT", "VHTTDL", "KHĐT", "GTVT",
		"KHCN", "VT", "VP", "TP", "TX", "UV", "QĐ", "NĐ", "PCT", "CT",
		"GĐ", "PGĐ", "TT", "KT", "HC", "TC", "TCCB", "TCHC", "KHTC", "NV",
		"PV", "XD"
	});

	private static readonly Regex StrongSignatureLine = new Regex("^(NOI NHAN|TM\\.|KT\\.|TL\\.|TUQ\\.|Q\\.|T/M |K/T |T/L |TM |KT |TL )", RegexOptions.Compiled);

	private static readonly string[] SignerTitleLines = new string[24]
	{
		"CHU TICH", "PHO CHU TICH", "BI THU", "PHO BI THU", "GIAM DOC", "PHO GIAM DOC", "BO TRUONG", "THU TRUONG", "CUC TRUONG", "PHO CUC TRUONG",
		"VU TRUONG", "HIEU TRUONG", "VIEN TRUONG", "TRUONG BAN", "PHO TRUONG BAN", "CHANH VAN PHONG", "CHINH UY", "PHO CHINH UY", "CHI HUY TRUONG", "TU LENH",
		"TONG GIAM DOC", "TRUONG PHONG", "CHU NHIEM", "NGUOI KY"
	};

	internal static bool IsAllUpperText(string s)
	{
		if (string.IsNullOrEmpty(s))
		{
			return false;
		}
		int num = 0;
		foreach (char c in s)
		{
			if (char.IsLower(c))
			{
				return false;
			}
			if (char.IsLetter(c))
			{
				num++;
			}
		}
		return num >= 4;
	}

	private static string ToSentenceCaseKeepAcronyms(string s)
	{
		return Regex.Replace(s, "[\\p{L}\\p{M}]+", (Match m) =>
		{
			string value = m.Value;
			if (RecipientAcronyms.Contains(value))
			{
				return value;
			}
			string text = ConvertToUnsigned(value).ToUpperInvariant();
			return (text.Length >= 2 && text.IndexOfAny("AEIOUY".ToCharArray()) < 0) ? value : value.ToLower(ViCulture);
		});
	}

	internal static string NormalizeRecipientLine(string line)
	{
		if (string.IsNullOrEmpty(line))
		{
			return line;
		}
		string text = line.Trim();
		if (text.Length == 0)
		{
			return line;
		}
		string input = ConvertToUnsigned(text).ToUpperInvariant();
		if (Regex.IsMatch(input, "^NOI\\s*NHAN\\s*:?$"))
		{
			return "Nơi nhận:";
		}
		Match match = Regex.Match(text, "^[-\\u2013\\u2014+•]\\s*(.+)$");
		if (!match.Success)
		{
			return line;
		}
		string value = match.Groups[1].Value;
		value = Regex.Replace(value, "\\s{2,}", " ");
		value = Regex.Replace(value, "\\s+(?=[,;:.)])", "");
		value = Regex.Replace(value, "\\(\\s+", "(");
		value = Regex.Replace(value, "([,:])(?=\\p{L})", "$1 ");
		if (IsAllUpperText(value))
		{
			value = ToSentenceCaseKeepAcronyms(value);
		}
		if (value.Length > 0 && char.IsLower(value[0]))
		{
			value = char.ToUpper(value[0], ViCulture) + value.Substring(1);
		}
		return "- " + value.Trim();
	}

	private static void NormalizeRecipientParagraph(Paragraph p)
	{
		try
		{
			Range range = p.Range;
			string text = range.Text ?? "";
			int num = text.Length;
			while (num > 0 && (text[num - 1] == '\r' || text[num - 1] == '\a' || text[num - 1] == '\n'))
			{
				num--;
			}
			if (num == 0)
			{
				return;
			}
			string text2 = text.Substring(0, num);
			if (text2.IndexOf('\t') < 0)
			{
				string[] array = text2.Split('\v');
				for (int i = 0; i < array.Length; i++)
				{
					array[i] = NormalizeRecipientLine(array[i]);
				}
				string text3 = string.Join("\v", array);
				if (!(text3 == text2) && range.Fields.Count <= 0)
				{
					Range comObj = range.Duplicate;
					comObj.SetRange(range.Start, range.Start + num);
					comObj.Text = text3;
					ComReleaser.Release(ref comObj);
				}
			}
		}
		catch
		{
		}
	}

	public static string ConvertToUnsigned(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return "";
		}
		string pattern = "\\p{IsCombiningDiacriticalMarks}+";
		string input = text.Normalize(NormalizationForm.FormD);
		return Regex.Replace(input, pattern, string.Empty).Replace('đ', 'd').Replace('Đ', 'D');
	}

	public static bool HasAtLeastTwoColumns(Table tbl)
	{
		if (tbl == null)
		{
			return false;
		}
		try
		{
			if (tbl.Columns.Count >= 2)
			{
				return true;
			}
		}
		catch
		{
		}
		try
		{
			if (tbl.Rows.Count > 0 && tbl.Rows[1].Cells.Count >= 2)
			{
				return true;
			}
		}
		catch
		{
		}
		try
		{
			if (tbl.Range.Cells.Count >= 2)
			{
				return true;
			}
		}
		catch
		{
		}
		return false;
	}

	public static bool IsDocumentAttachedToDecision(Document doc)
	{
		if (doc == null || doc.Paragraphs.Count == 0)
		{
			return false;
		}
		int num = 0;
		for (int i = 1; i <= doc.Paragraphs.Count; i++)
		{
			Paragraph paragraph = doc.Paragraphs[i];
			if (paragraph.Range.Tables.Count <= 0)
			{
				if (++num > 10)
				{
					break;
				}
				string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (IsAttachedDecisionIndicatorLine(text))
				{
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsAttachedDecisionIndicatorLine(string text)
	{
		string text2 = (text ?? "").Trim();
		if (text2.Length == 0 || text2.Length > 300)
		{
			return false;
		}
		string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
		if (!text3.Contains("KEM THEO QUYET DINH") && !text3.Contains("KEM THEO QUY DINH") && !text3.Contains("KEM THEO NGHI QUYET") && !text3.Contains("KEM THEO THONG TU") && !text3.Contains("KEM THEO CHI THI"))
		{
			return false;
		}
		string text4 = text3.TrimStart('(', '[', ' ', '*');
		if (!text2.StartsWith("(") && !text2.StartsWith("[") && !text4.StartsWith("BAN HANH KEM THEO"))
		{
			return text4.StartsWith("KEM THEO");
		}
		return true;
	}

	public static bool HasStrongSignatureAnchor(string tableText)
	{
		if (string.IsNullOrEmpty(tableText))
		{
			return false;
		}
		string[] array = ConvertToUnsigned(tableText).ToUpperInvariant().Split(new char[5] { '\r', '\n', '\a', '\v', '\t' }, StringSplitOptions.RemoveEmptyEntries);
		string[] array2 = array;
		foreach (string text in array2)
		{
			string text2 = text.Trim();
			if (text2.Length == 0)
			{
				continue;
			}
			if (StrongSignatureLine.IsMatch(text2))
			{
				return true;
			}
			string[] signerTitleLines = SignerTitleLines;
			foreach (string text3 in signerTitleLines)
			{
				if (text2 == text3)
				{
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsLikelySignatureTable(Document doc, Table tbl)
	{
		if (tbl == null)
		{
			return false;
		}
		try
		{
			string text = tbl.Range.Text ?? "";
			string text2 = ConvertToUnsigned(text).ToUpperInvariant();
			if (text2.Contains("CONG HOA XA HOI") || HeaderRebuilder.ContainsOrganizationTitle(text2) || text2.Contains("DOC LAP - TU DO"))
			{
				return false;
			}
			if (tbl.Rows.Count > 6)
			{
				return false;
			}
			if (!HasStrongSignatureAnchor(text))
			{
				return false;
			}
			if (doc != null)
			{
				int end = doc.Content.End;
				if (end > 2000 && (double)tbl.Range.Start < (double)end * 0.3)
				{
					return false;
				}
			}
			return true;
		}
		catch
		{
			return false;
		}
	}

	private static bool IsPlaceholderSignatureText(string text)
	{
		string text2 = ConvertToUnsigned(text ?? "").ToUpperInvariant();
		string[] array = text2.Split(new char[4] { '\r', '\n', '\a', '\v' }, StringSplitOptions.RemoveEmptyEntries);
		foreach (string text3 in array)
		{
			string text4 = text3.Trim();
			if (text4 == "TRAN VAN B" || text4 == "NGUYEN VAN A")
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsOwnedPlaceholderSignature(Document doc, Table table)
	{
		if (table.Title == "VFormat_SignatureTemplate" && IsPlaceholderSignatureText(table.Range.Text))
		{
			return !HeaderRebuilder.HasVisualContent(doc, table.Range);
		}
		return false;
	}

	public static bool IsSignatureOrRecipientBlock(string rawText)
	{
		string text = (rawText ?? "").Trim();
		if (text.Length == 0)
		{
			return false;
		}
		string text2 = ConvertToUnsigned(text).ToUpperInvariant();
		if (text2.Contains("NOI NHAN") || text2.Contains("LUU: VT") || text2.Contains("LUU VT") || text2.Contains("LUU: TCHC") || text2.Contains("NHU TREN") || text2.Contains("NHU THANH PHAN") || text2.Contains("LUU PN") || text2.Contains("LUU:") || text2.StartsWith("- LUU") || text2.StartsWith("- NHU"))
		{
			return true;
		}
		if (text2.Contains("TM.") || text2.Contains("KT.") || text2.Contains("TL.") || text2.Contains("TUQ.") || text2.Contains("T/M") || text2.Contains("K/T") || text2.Contains("T/L"))
		{
			return true;
		}
		string[] array = new string[16]
		{
			"CHU TICH", "BI THU", "GIAM DOC", "HIEU TRUONG", "VIEN TRUONG", "TRUONG BAN", "TRUONG PHONG", "CHU NHIEM", "CHI HUY TRUONG", "CHINH UY",
			"CUC TRUONG", "VU TRUONG", "THU TRUONG", "TONG GIAM DOC", "NGUOI KY", "DAI DIEN"
		};
		string[] array2 = array;
		foreach (string value in array2)
		{
			if (text2.Contains(value) && (text.Length <= 60 || text.Contains("\t") || text.Contains("  ") || text2.EndsWith(value) || text2.StartsWith(value)))
			{
				return true;
			}
		}
		return false;
	}

	public static bool CleanDuplicatePlaceholderSignatures(Document doc)
	{
		if (doc == null || doc.Tables.Count < 2)
		{
			return false;
		}
		try
		{
			List<Table> list = new List<Table>();
			for (int i = 1; i <= doc.Tables.Count; i++)
			{
				Table table = doc.Tables[i];
				if (IsLikelySignatureTable(doc, table))
				{
					list.Add(table);
				}
			}
			if (list.Count > 1)
			{
				bool flag = false;
				foreach (Table item in list)
				{
					if (!IsOwnedPlaceholderSignature(doc, item))
					{
						flag = true;
						break;
					}
				}
				if (!flag)
				{
					return false;
				}
				bool result = false;
				for (int num = list.Count - 1; num >= 0; num--)
				{
					if (IsOwnedPlaceholderSignature(doc, list[num]))
					{
						try
						{
							list[num].Delete();
							result = true;
						}
						catch
						{
						}
					}
				}
				return result;
			}
		}
		catch
		{
		}
		return false;
	}

	public static int LastContentParagraphIndex(Document doc)
	{
		if (doc == null)
		{
			return 0;
		}
		int count = doc.Paragraphs.Count;
		int num = 0;
		int num2 = count;
		while (num2 >= 1 && num < 300)
		{
			try
			{
				Paragraph paragraph = doc.Paragraphs[num2];
				if (paragraph.Range.Tables.Count > 0)
				{
					return num2;
				}
				string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Replace("\f", "")
					.Replace("\v", "")
					.Replace("\t", "")
					.Trim();
				if (text.Length > 0)
				{
					return num2;
				}
				if (paragraph.Range.InlineShapes.Count > 0)
				{
					return num2;
				}
			}
			catch
			{
				return num2;
			}
			num2--;
			num++;
		}
		return count;
	}

	public static bool HasRawSignatureCueNearEnd(Document doc, int window = 30)
	{
		if (doc == null)
		{
			return false;
		}
		try
		{
			int num = LastContentParagraphIndex(doc);
			for (int num2 = num; num2 >= Math.Max(1, num - window); num2--)
			{
				Paragraph paragraph = doc.Paragraphs[num2];
				string text = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				if (text.Length != 0 && text.Length <= 80)
				{
					string text2 = ConvertToUnsigned(text).ToUpperInvariant().Trim();
					if (StrongSignatureLine.IsMatch(text2))
					{
						return true;
					}
					string[] signerTitleLines = SignerTitleLines;
					foreach (string text3 in signerTitleLines)
					{
						if (text2 == text3)
						{
							return true;
						}
					}
				}
			}
		}
		catch
		{
		}
		return false;
	}

	public static bool IsDocEndingWithSignatureBlock(Document doc)
	{
		if (doc == null || doc.Paragraphs.Count == 0)
		{
			return false;
		}
		try
		{
			if (doc.Tables.Count > 0)
			{
				Table table = doc.Tables[doc.Tables.Count];
				if ((double)table.Range.Start > (double)doc.Content.End * 0.35)
				{
					string text = ConvertToUnsigned(table.Range.Text).ToUpperInvariant();
					if ((text.Contains("NOI NHAN") || text.Contains("LUU: VT") || text.Contains("LUU VT") || text.Contains("LUU: VP") || text.Contains("LUU: CHI BO")) && (text.Contains("CHU TICH") || text.Contains("GIAM DOC") || text.Contains("BO TRUONG") || text.Contains("CUC TRUONG") || text.Contains("BI THU") || text.Contains("CHANH VAN PHONG") || text.Contains("TRUONG BAN") || text.Contains("TM.") || text.Contains("KT.") || text.Contains("TL.") || text.Contains("T/M") || text.Contains("K/T") || text.Contains("T/L") || text.Contains("CHINH UY") || text.Contains("CHI HUY TRUONG") || text.Contains("NGUOI KY")))
					{
						return true;
					}
				}
			}
			int num = LastContentParagraphIndex(doc);
			int num2 = Math.Min(num, 12);
			for (int num3 = num; num3 >= num - num2 + 1; num3--)
			{
				Paragraph paragraph = doc.Paragraphs[num3];
				if (paragraph.Range.Tables.Count != 0)
				{
					continue;
				}
				string text2 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
				if (text3.Length <= 0 || text3.Length >= 60 || text3.StartsWith("DIEU ") || text3.Contains("CHIU TRACH NHIEM") || text3.Contains("THI HANH"))
				{
					continue;
				}
				if (!text3.StartsWith("NOI NHAN") && !text3.StartsWith("- LUU:") && !text3.StartsWith("LUU: VT") && !text3.StartsWith("LUU: VP") && !text3.StartsWith("TM.") && !text3.StartsWith("KT.") && !text3.StartsWith("TL.") && !text3.StartsWith("TUQ.") && !text3.StartsWith("T/M") && !text3.StartsWith("K/T") && !text3.StartsWith("T/L"))
				{
					switch (text3)
					{
					case "CHU TICH":
					case "GIAM DOC":
					case "BO TRUONG":
					case "CUC TRUONG":
					case "BI THU":
					case "PHO BI THU":
					case "PHO CHU TICH":
					case "PHO GIAM DOC":
					case "CHINH UY":
					case "CHI HUY TRUONG":
						break;
					default:
						continue;
					}
				}
				return true;
			}
		}
		catch
		{
		}
		return false;
	}

	public static Table FindExistingSignatureTable(Document doc)
	{
		if (doc == null || doc.Tables.Count == 0)
		{
			return null;
		}
		for (int num = doc.Tables.Count; num >= 1; num--)
		{
			Table table = doc.Tables[num];
			if (HasAtLeastTwoColumns(table) && IsLikelySignatureTable(doc, table))
			{
				return table;
			}
		}
		return null;
	}

	public static bool ConvertRawSignatureToTable(Document doc, RulePackType rulePack, ProfileModel profile = null)
	{
		if (doc == null || doc.Paragraphs.Count < 5)
		{
			return false;
		}
		if (IsDocumentAttachedToDecision(doc))
		{
			return false;
		}
		int num = LastContentParagraphIndex(doc);
		int num2 = Math.Max(1, num - 12);
		string text = "";
		string text2 = "";
		int num3 = 0;
		int num4 = 0;
		List<int> list = new List<int>();
		bool flag = false;
		for (int i = num2; i <= num; i++)
		{
			Paragraph paragraph = doc.Paragraphs[i];
			if (paragraph.Range.Tables.Count != 0)
			{
				continue;
			}
			string text3 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "");
			string text4 = text3.Trim();
			if (text4.Length <= 0)
			{
				continue;
			}
			if (text4.Length > 120)
			{
				if (flag)
				{
					break;
				}
				continue;
			}
			string text5 = ConvertToUnsigned(text4).ToUpperInvariant();
			bool flag2 = text5.StartsWith("DIEU ") || text5.Contains("CHIU TRACH NHIEM") || text5.Contains("THI HANH");
			bool flag3 = text5.StartsWith("KINH GUI") || RegexPatterns.DirectRecipient.IsMatch(text4);
			if ((flag2 || flag3 || !HasStrongSignatureAnchor(text4)) && !flag)
			{
				continue;
			}
			flag = true;
			if (num3 == 0)
			{
				num3 = i;
			}
			num4 = i;
			if (HeaderRebuilder.HasVisualContent(doc, paragraph.Range))
			{
				return false;
			}
			list.Add(i);
			bool flag4 = text3.StartsWith("\t") || text3.StartsWith("   ") || paragraph.Format.Alignment == WdParagraphAlignment.wdAlignParagraphRight;
			string text6 = "";
			string text7 = "";
			string input = text3.Trim();
			input = Regex.Replace(input, "(?<=\\S) {2,3}(?=\\S)", " ");
			string[] array = Regex.Split(input, "\\t+|\\s{4,}");
			List<string> list2 = new List<string>();
			string[] array2 = array;
			foreach (string text8 in array2)
			{
				if (!string.IsNullOrEmpty(text8) && !string.IsNullOrEmpty(text8.Trim()))
				{
					list2.Add(text8.Trim());
				}
			}
			if (list2.Count >= 2 && (list2[0] == "-" || list2[0] == "+" || list2[0] == "*" || list2[0] == "–" || list2[0] == "—"))
			{
				list2[1] = list2[0] + " " + list2[1];
				list2.RemoveAt(0);
			}
			string[] array3 = list2.ToArray();
			if (array3.Length >= 2)
			{
				text6 = array3[0];
				text7 = string.Join(" ", array3, 1, array3.Length - 1);
			}
			else if (array3.Length == 1)
			{
				string text9 = array3[0];
				string text10 = ConvertToUnsigned(text9).ToUpperInvariant();
				if (text10.Contains("NOI NHAN") || text9.StartsWith("-") || text9.StartsWith("+") || text10.Contains("LUU") || text10.Contains("NHU TREN") || text10.Contains("NHU THANH PHAN"))
				{
					text6 = text9;
				}
				else if (text10.Contains("TM.") || text10.Contains("KT.") || text10.Contains("TL.") || text10.Contains("TUQ.") || text10.Contains("CHU TICH") || text10.Contains("BI THU") || text10.Contains("GIAM DOC") || flag4)
				{
					text7 = text9;
				}
				else if (text9.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length <= 5 && !text9.StartsWith("-"))
				{
					text7 = text9;
				}
				else
				{
					text6 = text9;
				}
			}
			if (!string.IsNullOrEmpty(text7))
			{
				string text11 = ConvertToUnsigned(text7).ToUpperInvariant();
				if (text11.StartsWith("LUU") || text11.Contains("LUU:") || text11.Contains("LUU ") || text11.Contains("LUU.") || text11.Contains("NHU TREN") || text11.Contains("NHU THANH PHAN") || text11.Contains("NOI NHAN") || text7.StartsWith("-") || text7.StartsWith("+"))
				{
					text6 = ((!string.IsNullOrEmpty(text6)) ? (text6 + " " + text7) : text7);
					text7 = "";
				}
			}
			if (!string.IsNullOrEmpty(text7) && (text7.ToUpperInvariant().Contains("NOI NHAN") || text7.ToLowerInvariant().Contains("nơi nhận")))
			{
				text7 = Regex.Replace(text7, "(?i)nơi\\s*nhận\\s*:?\\s*", "").Trim();
				if (!text.ToLowerInvariant().Contains("nơi nhận") && !text6.ToLowerInvariant().Contains("nơi nhận"))
				{
					text6 = ((!string.IsNullOrEmpty(text6)) ? ("Nơi nhận:\r\n" + text6) : "Nơi nhận:");
				}
			}
			if (!string.IsNullOrEmpty(text6))
			{
				if (text.Length > 0)
				{
					text += "\r\n";
				}
				text += text6;
			}
			if (!string.IsNullOrEmpty(text7))
			{
				if (text2.Length > 0)
				{
					text2 += "\r\n";
				}
				text2 += text7;
			}
		}
		int num5 = 0;
		for (int k = num4 + 1; k <= num; k++)
		{
			try
			{
				string text12 = (doc.Paragraphs[k].Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (text12.Length > 0)
				{
					num5++;
				}
			}
			catch
			{
			}
		}
		if (num3 == 0 || num4 == 0 || num5 > 2)
		{
			return false;
		}
		Range comObj = doc.Paragraphs[num3].Range.Duplicate;
		for (int num6 = list.Count - 1; num6 >= 0; num6--)
		{
			if (list[num6] != num3)
			{
				Range range = doc.Paragraphs[list[num6]].Range;
				if (range.Tables.Count == 0)
				{
					object Unit = Missing.Value;
					object Count = Missing.Value;
					range.Delete(ref Unit, ref Count);
				}
			}
		}
		comObj.Text = "\r\r\r";
		int num7 = comObj.Start + 1;
		ComReleaser.Release(ref comObj);
		object Start = num7;
		object End = num7;
		Range range2 = doc.Range(ref Start, ref End);
		Tables tables = doc.Tables;
		object DefaultTableBehavior = Missing.Value;
		object AutoFitBehavior = Missing.Value;
		Table table = tables.Add(range2, 1, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
		table.Borders.Enable = 0;
		if (string.IsNullOrEmpty(text))
		{
			text = (RulePackTraits.IsYouthUnion(rulePack) ? "Nơi nhận:\r\n- Như trên;\r\n- Lưu VP." : "Nơi nhận:\r\n- Như trên;\r\n- Lưu: VT.");
		}
		table.Cell(1, 1).Range.Text = text;
		if (string.IsNullOrEmpty(text2))
		{
			if (rulePack == RulePackType.HD05_VPTW)
			{
				text2 = "T/M BAN CHẤP HÀNH\r\nBÍ THƯ\r\n\r\n\r\n\r\nNguyễn Văn A";
			}
			else
			{
				text2 = ((!RulePackTraits.IsYouthUnion(rulePack)) ? "TM. ỦY BAN NHÂN DÂN\r\nCHỦ TỊCH\r\n\r\n\r\n\r\nTrần Văn B" : "TM. BAN THƯỜNG VỤ\r\nBÍ THƯ\r\n\r\n\r\n\r\nNguyễn Văn A");
			}
		}
		table.Cell(1, 2).Range.Text = text2;
		FormatExistingSignatureTable(table, rulePack, profile);
		return true;
	}

	public static void FormatExistingSignatureTable(Table tbl, RulePackType rulePack, ProfileModel profile = null)
	{
		if (tbl == null)
		{
			return;
		}
		if (HeaderRebuilder.HasVisualContent(tbl.Range.Document, tbl.Range))
		{
			if (profile != null)
			{
				SemanticComponentFormatter.Apply(tbl.Range.Document, profile, tbl.Range, reflowTables: false);
			}
			return;
		}
		try
		{
			tbl.Borders.Enable = 0;
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
			for (int i = 1; i <= tbl.Rows.Count; i++)
			{
				Row row = tbl.Rows[i];
				row.Borders.Enable = 0;
				for (int j = 1; j <= row.Cells.Count; j++)
				{
					Cell cell = row.Cells[j];
					cell.Borders.Enable = 0;
					cell.LeftPadding = 0f;
					cell.RightPadding = 0f;
					foreach (Paragraph paragraph4 in cell.Range.Paragraphs)
					{
						try
						{
							paragraph4.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
							paragraph4.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
							paragraph4.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
							paragraph4.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
						}
						catch
						{
						}
					}
				}
			}
			bool flag = rulePack == RulePackType.HD05_VPTW;
			bool flag2 = RulePackTraits.IsYouthUnion(rulePack);
			float num = 0f;
			try
			{
				num = tbl.Range.Document.PageSetup.PageWidth - tbl.Range.Document.PageSetup.LeftMargin - tbl.Range.Document.PageSetup.RightMargin;
			}
			catch
			{
			}
			if (num < UnitConverter.MmToPt(140f) || num > UnitConverter.MmToPt(200f))
			{
				num = UnitConverter.MmToPt(165f);
			}
			float num2 = UnitConverter.MmToPt(75f);
			float num3 = num - num2;
			if (num3 < UnitConverter.MmToPt(85f))
			{
				num3 = UnitConverter.MmToPt(85f);
			}
			tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
			tbl.PreferredWidth = num2 + num3;
			try
			{
				tbl.Columns[1].Width = num2;
			}
			catch
			{
			}
			try
			{
				tbl.Columns[2].Width = num3;
			}
			catch
			{
			}
			for (int k = 1; k <= tbl.Rows.Count; k++)
			{
				try
				{
					if (tbl.Rows[k].Cells.Count >= 2)
					{
						tbl.Rows[k].Cells[1].LeftPadding = 0f;
						tbl.Rows[k].Cells[1].RightPadding = 0f;
						tbl.Rows[k].Cells[2].LeftPadding = 0f;
						tbl.Rows[k].Cells[2].RightPadding = 0f;
						tbl.Rows[k].Cells[1].Width = num2;
						tbl.Rows[k].Cells[2].Width = num3;
					}
				}
				catch
				{
				}
			}
			for (int l = 1; l <= tbl.Rows.Count; l++)
			{
				try
				{
					if (tbl.Rows[l].Cells.Count < 1)
					{
						continue;
					}
					Cell cell2 = tbl.Rows[l].Cells[1];
					cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
					cell2.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
					cell2.Range.ParagraphFormat.SpaceBefore = 0f;
					cell2.Range.ParagraphFormat.SpaceAfter = 0f;
					cell2.Range.ParagraphFormat.FirstLineIndent = 0f;
					cell2.Range.ParagraphFormat.LeftIndent = 0f;
					cell2.Range.ParagraphFormat.RightIndent = 0f;
					cell2.Range.Font.Name = "Times New Roman";
					int count = cell2.Range.Paragraphs.Count;
					for (int m = 1; m <= count; m++)
					{
						Paragraph paragraph2 = cell2.Range.Paragraphs[m];
						paragraph2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
						paragraph2.Format.FirstLineIndent = 0f;
						paragraph2.Format.LeftIndent = 0f;
						paragraph2.Format.RightIndent = 0f;
						paragraph2.Format.SpaceBefore = 0f;
						paragraph2.Format.SpaceAfter = 0f;
						paragraph2.Range.Font.Name = "Times New Roman";
						try
						{
							paragraph2.Range.Font.AllCaps = 0;
							paragraph2.Range.Font.SmallCaps = 0;
						}
						catch
						{
						}
						NormalizeRecipientParagraph(paragraph2);
						string text = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
						string text2 = ConvertToUnsigned(text).ToUpperInvariant();
						if (text2.StartsWith("NOI NHAN"))
						{
							float headerSize = flag ? 14f : (flag2 ? 13f : 12f);
							int headerBold = flag ? 0 : (-1);
							int headerItalic = (!flag && !flag2) ? (-1) : 0;
							WdUnderline headerUnderline = flag ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone;
							string rawParaText = paragraph2.Range.Text ?? "";
							int softBreakIdx = rawParaText.IndexOf('');
							if (softBreakIdx > 0)
							{
								paragraph2.Range.Font.Size = flag ? 12f : 11f;
								paragraph2.Range.Font.Bold = 0;
								paragraph2.Range.Font.Italic = 0;
								paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
								Range headerLineRange = paragraph2.Range.Duplicate;
								headerLineRange.SetRange(paragraph2.Range.Start, paragraph2.Range.Start + softBreakIdx);
								headerLineRange.Font.Size = headerSize;
								headerLineRange.Font.Bold = headerBold;
								headerLineRange.Font.Italic = headerItalic;
								headerLineRange.Font.Underline = headerUnderline;
								ComReleaser.Release(ref headerLineRange);
							}
							else
							{
								paragraph2.Range.Font.Size = headerSize;
								paragraph2.Range.Font.Bold = headerBold;
								paragraph2.Range.Font.Italic = headerItalic;
								paragraph2.Range.Font.Underline = headerUnderline;
							}
						}
						else if (text.Length > 0)
						{
							paragraph2.Range.Font.Size = (flag ? 12f : 11f);
							paragraph2.Range.Font.Bold = 0;
							paragraph2.Range.Font.Italic = 0;
							paragraph2.Range.Font.Underline = WdUnderline.wdUnderlineNone;
						}
					}
				}
				catch
				{
				}
			}
			Cell cell3 = tbl.Cell(1, 2);
			cell3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
			cell3.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
			cell3.Range.ParagraphFormat.SpaceBefore = 0f;
			cell3.Range.ParagraphFormat.SpaceAfter = 0f;
			cell3.Range.ParagraphFormat.FirstLineIndent = 0f;
			cell3.Range.ParagraphFormat.LeftIndent = 0f;
			cell3.Range.ParagraphFormat.RightIndent = 0f;
			cell3.Range.Font.Name = "Times New Roman";
			bool flag3 = rulePack == RulePackType.TT68_2023_BQP;
			int count2 = cell3.Range.Paragraphs.Count;
			int num4 = -1;
			for (int num5 = count2; num5 >= 1; num5--)
			{
				string text3 = (cell3.Range.Paragraphs[num5].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				if (text3.Length > 0)
				{
					num4 = num5;
					break;
				}
			}
			int num6 = -1;
			if (num4 > 1)
			{
				for (int num7 = num4 - 1; num7 >= 1; num7--)
				{
					string text4 = (cell3.Range.Paragraphs[num7].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (text4.Length > 0)
					{
						num6 = num7;
						break;
					}
				}
			}
			if (num6 > 0 && num4 > num6)
			{
				int num8 = num4 - 1 - num6;
				if (num8 < 4)
				{
					int num9 = 4 - num8;
					for (int n = 0; n < num9; n++)
					{
						try
						{
							cell3.Range.Paragraphs[num6].Range.InsertParagraphAfter();
						}
						catch
						{
						}
					}
				}
				else if (num8 > 5)
				{
					int num10 = num8 - 4;
					for (int num11 = 0; num11 < num10; num11++)
					{
						try
						{
							Range range = cell3.Range.Paragraphs[num6 + 1].Range;
							object Unit = Missing.Value;
							object Count = Missing.Value;
							range.Delete(ref Unit, ref Count);
						}
						catch
						{
						}
					}
				}
				count2 = cell3.Range.Paragraphs.Count;
				for (int num12 = count2; num12 >= 1; num12--)
				{
					string text5 = (cell3.Range.Paragraphs[num12].Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
					if (text5.Length > 0)
					{
						num4 = num12;
						break;
					}
				}
			}
			for (int num13 = 1; num13 <= count2; num13++)
			{
				Paragraph paragraph3 = cell3.Range.Paragraphs[num13];
				try
				{
					paragraph3.Format.TabStops.ClearAll();
					Range range2;
					object Unit2;
					object Count2;
					for (int num14 = 0; paragraph3.Range.Characters.Count > 1 && num14++ < 30; range2 = paragraph3.Range.Characters[1], Unit2 = Missing.Value, Count2 = Missing.Value, range2.Delete(ref Unit2, ref Count2))
					{
						switch (paragraph3.Range.Characters[1].Text)
						{
						case " ":
						case "\t":
						case "\u00a0":
						case "\u3000":
							continue;
						}
						break;
					}
				}
				catch
				{
				}
				try
				{
					bool flag4 = num13 == num4;
					bool flag5 = paragraph3.Range.Font.AllCaps != 0 || paragraph3.Range.Font.SmallCaps != 0;
					string text6 = paragraph3.Range.Text ?? "";
					int num15 = text6.Length;
					while (num15 > 0 && (text6[num15 - 1] == '\r' || text6[num15 - 1] == '\a'))
					{
						num15--;
					}
					if (num15 > 0 && paragraph3.Range.Fields.Count == 0)
					{
						string text7 = text6.Substring(0, num15);
						string text8 = text7;
						if (flag4 && !flag3 && IsAllUpperText(text7))
						{
							text8 = ViCulture.TextInfo.ToTitleCase(text7.ToLower(ViCulture));
						}
						else if (!flag4 && flag5)
						{
							text8 = text7.ToUpper(ViCulture);
						}
						if (text8 != text7)
						{
							Range comObj = paragraph3.Range.Duplicate;
							comObj.SetRange(paragraph3.Range.Start, paragraph3.Range.Start + num15);
							comObj.Text = text8;
							ComReleaser.Release(ref comObj);
						}
					}
					if (flag5)
					{
						paragraph3.Range.Font.AllCaps = 0;
						paragraph3.Range.Font.SmallCaps = 0;
					}
				}
				catch
				{
				}
				paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				paragraph3.Format.FirstLineIndent = 0f;
				paragraph3.Format.LeftIndent = 0f;
				paragraph3.Format.RightIndent = 0f;
				paragraph3.Format.SpaceBefore = 0f;
				paragraph3.Format.SpaceAfter = 0f;
				paragraph3.Range.Font.Name = "Times New Roman";
				paragraph3.Range.Font.Underline = WdUnderline.wdUnderlineNone;
				string text9 = (paragraph3.Range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
				string text10 = ConvertToUnsigned(text9).ToUpperInvariant();
				if (num13 == num4 && text9.Length > 1)
				{
					paragraph3.Range.Font.Size = 14f;
					paragraph3.Range.Font.Bold = -1;
					paragraph3.Range.Font.Italic = 0;
					paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph3.Format.SpaceBefore = 18f;
					paragraph3.Format.SpaceAfter = 0f;
				}
				else if (num13 > num6 && num13 < num4 && num6 > 0)
				{
					paragraph3.Range.Font.Size = 14f;
					paragraph3.Range.Font.Bold = 0;
					paragraph3.Range.Font.Italic = 0;
					paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					paragraph3.Format.SpaceBefore = 0f;
					paragraph3.Format.SpaceAfter = 6f;
					paragraph3.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				}
				else if (flag3 && RegexPatterns.MilitaryRank.IsMatch(text9))
				{
					paragraph3.Range.Font.Size = 14f;
					paragraph3.Range.Font.Bold = -1;
					paragraph3.Range.Font.Italic = 0;
					paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				}
				else if (text10.StartsWith("TM.") || text10.StartsWith("KT.") || text10.StartsWith("TL.") || text10.StartsWith("TUQ.") || text10.StartsWith("T/M") || text10.StartsWith("K/T") || text10.StartsWith("T/L") || text10.StartsWith("Q. "))
				{
					paragraph3.Range.Font.Size = 14f;
					paragraph3.Range.Font.Bold = -1;
					paragraph3.Range.Font.Italic = 0;
					paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				}
				else if (text9.Length > 0)
				{
					paragraph3.Range.Font.Size = 14f;
					paragraph3.Range.Font.Bold = ((!flag && !flag2) ? (-1) : 0);
					paragraph3.Range.Font.Italic = 0;
					paragraph3.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				}
			}
			if (tbl.Rows.Count >= 2)
			{
				try
				{
					Row row2 = tbl.Rows[tbl.Rows.Count];
					row2.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
					row2.Height = UnitConverter.MmToPt(38f);
					if (row2.Cells.Count >= 2)
					{
						Cell cell4 = row2.Cells[2];
						cell4.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
						cell4.Range.Font.Name = "Times New Roman";
						cell4.Range.Font.Size = 14f;
						cell4.Range.Font.Bold = -1;
						cell4.Range.ParagraphFormat.SpaceBefore = UnitConverter.MmToPt(30f);
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
		if (profile != null)
		{
			SemanticComponentFormatter.Apply(tbl.Range.Document, profile, tbl.Range, reflowTables: false);
		}
	}

	public static Table RebuildSignatureCleanly(Document doc, SignatureData data, RulePackType rulePack, Table existingTable = null, int? targetInsertPos = null, ProfileModel profile = null)
	{
		if (doc == null)
		{
			return null;
		}
		if (data == null)
		{
			data = ComponentDataExtractor.ExtractSignatureData(doc, existingTable, rulePack);
		}
		int num = -1;
		if (existingTable == null)
		{
			existingTable = FindExistingSignatureTable(doc);
		}
		if (existingTable != null)
		{
			if (HeaderRebuilder.HasVisualContent(doc, existingTable.Range))
			{
				FormatExistingSignatureTable(existingTable, rulePack, profile);
				return existingTable;
			}
			try
			{
				num = existingTable.Range.Start;
				int end = existingTable.Range.End;
				Table table = null;
				try
				{
					for (int i = 1; i <= doc.Tables.Count; i++)
					{
						Table table2 = doc.Tables[i];
						if (table2.Range.Start != num && table2.Range.Start >= end && table2.Range.Start <= end + 150 && table2.Title == "VFormat_SignatureTemplate" && IsLikelySignatureTable(doc, table2))
						{
							table = table2;
							break;
						}
					}
				}
				catch
				{
				}
				existingTable.Delete();
				if (table != null)
				{
					try
					{
						table.Delete();
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
		Range range;
		if (num >= 0)
		{
			int num2 = Math.Min(num, Math.Max(0, doc.Content.End - 1));
			object Start = num2;
			object End = num2;
			range = doc.Range(ref Start, ref End);
		}
		else if (targetInsertPos.HasValue)
		{
			int num3 = Math.Min(targetInsertPos.Value, Math.Max(0, doc.Content.End - 1));
			object Start2 = num3;
			object End2 = num3;
			range = doc.Range(ref Start2, ref End2);
		}
		else
		{
			Range content = doc.Content;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			content.Collapse(ref Direction);
			content.InsertParagraphAfter();
			range = doc.Content;
			Range range2 = range;
			object Direction2 = WdCollapseDirection.wdCollapseEnd;
			range2.Collapse(ref Direction2);
		}
		int num4 = HeaderRebuilder.EnsureNotAdjacentToTable(doc, range.Start);
		object Start3 = num4;
		object End3 = num4;
		range = doc.Range(ref Start3, ref End3);
		Tables tables = doc.Tables;
		Range range3 = range;
		object DefaultTableBehavior = Missing.Value;
		object AutoFitBehavior = Missing.Value;
		Table table3 = tables.Add(range3, 1, 2, ref DefaultTableBehavior, ref AutoFitBehavior);
		table3.Borders.Enable = 0;
		List<string> list = new List<string>();
		list.Add("Nơi nhận:");
		if (data.Recipients != null && data.Recipients.Count > 0)
		{
			list.AddRange(data.Recipients);
		}
		else
		{
			list.Add("- Như trên;");
			list.Add(RulePackTraits.IsYouthUnion(rulePack) ? "- Lưu VP." : "- Lưu: VT.");
		}
		table3.Cell(1, 1).Range.Text = string.Join("\r\n", list.ToArray());
		List<string> list2 = new List<string>();
		if (!string.IsNullOrEmpty(data.Authority))
		{
			list2.Add(data.Authority);
		}
		if (!string.IsNullOrEmpty(data.Title))
		{
			list2.Add(data.Title);
		}
		list2.Add("");
		list2.Add("");
		list2.Add("");
		list2.Add("");
		string text;
		if (string.IsNullOrEmpty(data.SignerName))
		{
			text = ((rulePack == RulePackType.HD05_VPTW || RulePackTraits.IsYouthUnion(rulePack)) ? "Nguyễn Văn A" : "Trần Văn B");
		}
		else
		{
			text = data.SignerName;
		}
		if (rulePack == RulePackType.TT68_2023_BQP && !string.IsNullOrEmpty(data.MilitaryRank) && !text.ToLowerInvariant().Contains(data.MilitaryRank.ToLowerInvariant()))
		{
			text = data.MilitaryRank.Trim() + " " + text.Trim();
		}
		list2.Add(text);
		table3.Cell(1, 2).Range.Text = string.Join("\r\n", list2.ToArray());
		FormatExistingSignatureTable(table3, rulePack, profile);
		return table3;
	}

	public static void AutoInsertMissingSignature(Document doc, RulePackType rulePack, ProfileModel profile = null)
	{
		if (doc == null)
		{
			return;
		}
		if (IsDocumentAttachedToDecision(doc))
		{
			Range content = doc.Content;
			object Direction = WdCollapseDirection.wdCollapseEnd;
			content.Collapse(ref Direction);
			content.InsertAfter("\r\n./.");
		}
		else if (FindExistingSignatureTable(doc) == null && !IsDocEndingWithSignatureBlock(doc) && !HasRawSignatureCueNearEnd(doc))
		{
			Table table = RebuildSignatureCleanly(doc, null, rulePack, null, null, profile);
			if (table != null)
			{
				table.Title = "VFormat_SignatureTemplate";
			}
		}
	}
}
