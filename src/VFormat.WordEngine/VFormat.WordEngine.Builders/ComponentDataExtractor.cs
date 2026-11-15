using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Rules;

namespace VFormat.WordEngine.Builders;

public static class ComponentDataExtractor
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

	public static HeaderData ExtractHeaderData(Document doc, Table existingTable, RulePackType rulePack)
	{
		HeaderData headerData = new HeaderData();
		if (doc == null)
		{
			return headerData;
		}
		List<string> list = new List<string>();
		string text = string.Empty;
		string text2 = string.Empty;
		if (existingTable != null)
		{
			try
			{
				int count = existingTable.Rows.Count;
				_ = existingTable.Columns.Count;
				for (int i = 1; i <= count; i++)
				{
					Row row = existingTable.Rows[i];
					for (int j = 1; j <= row.Cells.Count; j++)
					{
						Cell cell = row.Cells[j];
						string text3 = (cell.Range.Text ?? "").Replace("\a", "").Trim();
						if (string.IsNullOrEmpty(text3))
						{
							continue;
						}
						string[] array = text3.Split(new char[3] { '\r', '\n', '\v' }, StringSplitOptions.RemoveEmptyEntries);
						string[] array2 = array;
						foreach (string text4 in array2)
						{
							string text5 = text4.Trim();
							if (string.IsNullOrEmpty(text5))
							{
								continue;
							}
							string text6 = ConvertToUnsigned(text5).ToUpperInvariant();
							if (text6.Contains("CONG HOA XA HOI") || text6.Contains("DOC LAP - TU DO") || HeaderRebuilder.ContainsOrganizationTitle(text6))
							{
								continue;
							}
							if (HeaderRebuilder.IsAsteriskLine(text5))
							{
								headerData.HasAsterisk = true;
							}
							else
							{
								if (HeaderRebuilder.IsSeparatorLine(text5))
								{
									continue;
								}
								if (HeaderRebuilder.IsDocNumberLine(text5) || RegexPatterns.DocumentNumberGeneral.IsMatch(text5))
								{
									if (string.IsNullOrEmpty(text))
									{
										text = text5;
									}
								}
								else if (text6.Contains("NGAY ") && (text6.Contains("THANG ") || text6.Contains("NAM ")))
								{
									if (string.IsNullOrEmpty(text2))
									{
										text2 = text5;
									}
								}
								else
								{
									if (text6.StartsWith("KINH GUI") || text6.StartsWith("KINH TRINH") || RegexPatterns.DirectRecipient.IsMatch(text5))
									{
										continue;
									}
									if (RegexPatterns.DispatchSubject.IsMatch(text5) || text6.StartsWith("V/V ") || text6.StartsWith("V/V:") || text6.StartsWith("V/V.") || (!string.IsNullOrEmpty(headerData.DispatchSubject) && j == 1))
									{
										if (string.IsNullOrEmpty(headerData.DispatchSubject))
										{
											headerData.DispatchSubject = text5;
										}
										else
										{
											headerData.DispatchSubject = headerData.DispatchSubject + " " + text5;
										}
									}
									else if (j == 1 && string.IsNullOrEmpty(headerData.DispatchSubject) && !list.Contains(text5))
									{
										list.Add(text5);
									}
								}
							}
						}
					}
				}
				if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2))
				{
					int end = existingTable.Range.End;
					for (int l = 1; l <= Math.Min(doc.Paragraphs.Count, 15); l++)
					{
						Paragraph paragraph = doc.Paragraphs[l];
						if (paragraph.Range.Start < end || paragraph.Range.Tables.Count != 0)
						{
							continue;
						}
						string text7 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
							.Trim();
						if (!string.IsNullOrEmpty(text7))
						{
							if (HeaderRebuilder.IsDocumentTitle(text7))
							{
								break;
							}
							string text8 = ConvertToUnsigned(text7).ToUpperInvariant();
							if (string.IsNullOrEmpty(text) && (HeaderRebuilder.IsDocNumberLine(text7) || RegexPatterns.DocumentNumberGeneral.IsMatch(text7)))
							{
								text = text7;
							}
							else if (string.IsNullOrEmpty(text2) && text8.Contains("NGAY ") && (text8.Contains("THANG ") || text8.Contains("NAM ")))
							{
								text2 = text7;
							}
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
			int num = Math.Min(doc.Paragraphs.Count, 12);
			int num2 = Math.Min(doc.Paragraphs.Count, 80);
			for (int m = 1; m <= num; m++)
			{
				Paragraph paragraph2 = doc.Paragraphs[m];
				if (paragraph2.Range.Tables.Count > 0)
				{
					if (num < num2)
					{
						num++;
					}
					continue;
				}
				string text9 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (string.IsNullOrEmpty(text9))
				{
					continue;
				}
				string text10 = ConvertToUnsigned(text9).ToUpperInvariant();
				if (HeaderRebuilder.IsDocumentTitle(text10))
				{
					break;
				}
				string[] array3 = (text9.Contains("\t") ? text9.Split('\t') : new string[1] { text9 });
				string[] array4 = array3;
				foreach (string text11 in array4)
				{
					string text12 = text11.Trim();
					if (string.IsNullOrEmpty(text12))
					{
						continue;
					}
					string text13 = ConvertToUnsigned(text12).ToUpperInvariant();
					if (HeaderRebuilder.IsNationalOrMottoText(text13) || Regex.IsMatch(text13, "^VIET\\s*NAM\\.?$"))
					{
						continue;
					}
					if (HeaderRebuilder.IsAsteriskLine(text12))
					{
						headerData.HasAsterisk = true;
					}
					else
					{
						if (HeaderRebuilder.IsSeparatorLine(text12))
						{
							continue;
						}
						if (HeaderRebuilder.IsDocNumberLine(text12) || RegexPatterns.DocumentNumberGeneral.IsMatch(text12))
						{
							if (string.IsNullOrEmpty(text))
							{
								text = text12;
							}
						}
						else if (text13.Contains("NGAY ") && (text13.Contains("THANG ") || text13.Contains("NAM ")))
						{
							if (string.IsNullOrEmpty(text2))
							{
								text2 = text12;
							}
						}
						else
						{
							if (text13.StartsWith("KINH GUI") || text13.StartsWith("KINH TRINH") || RegexPatterns.DirectRecipient.IsMatch(text12))
							{
								continue;
							}
							if (RegexPatterns.DispatchSubject.IsMatch(text12) || text13.StartsWith("V/V ") || text13.StartsWith("V/V:") || text13.StartsWith("V/V.") || !string.IsNullOrEmpty(headerData.DispatchSubject))
							{
								if (string.IsNullOrEmpty(headerData.DispatchSubject))
								{
									headerData.DispatchSubject = text12;
								}
								else
								{
									headerData.DispatchSubject = headerData.DispatchSubject + " " + text12;
								}
							}
							else if (string.IsNullOrEmpty(headerData.DispatchSubject) && (RegexPatterns.PartyAgency.IsMatch(text12) || RegexPatterns.YouthUnionAgency.IsMatch(text12) || RegexPatterns.MilitaryAgency.IsMatch(text12) || RegexPatterns.AdminAgency.IsMatch(text12) || list.Count < 2) && !list.Contains(text12))
							{
								list.Add(text12);
							}
						}
					}
				}
			}
		}
		if (list.Count == 1)
		{
			headerData.IssuingAgency = list[0].Trim();
			headerData.SuperiorAgency = "";
			headerData.HasExtractedAgency = true;
		}
		else if (list.Count == 2)
		{
			if (HeaderRebuilder.IsSingleAgencySplitInTwo(list[0], list[1]))
			{
				headerData.SuperiorAgency = "";
				headerData.IssuingAgency = list[0].Trim() + "\r\n" + list[1].Trim();
			}
			else
			{
				headerData.SuperiorAgency = list[0].Trim();
				headerData.IssuingAgency = list[1].Trim();
			}
			headerData.HasExtractedAgency = true;
		}
		else if (list.Count == 3)
		{
			if (HeaderRebuilder.IsSingleAgencySplitInTwo(list[0], list[1]))
			{
				headerData.SuperiorAgency = list[0].Trim() + "\r\n" + list[1].Trim();
				headerData.IssuingAgency = list[2].Trim();
			}
			else if (HeaderRebuilder.IsSingleAgencySplitInTwo(list[1], list[2]))
			{
				headerData.SuperiorAgency = list[0].Trim();
				headerData.IssuingAgency = list[1].Trim() + "\r\n" + list[2].Trim();
			}
			else
			{
				headerData.SuperiorAgency = list[0].Trim() + "\r\n" + list[1].Trim();
				headerData.IssuingAgency = list[2].Trim();
			}
			headerData.HasExtractedAgency = true;
		}
		else if (list.Count >= 4)
		{
			if (HeaderRebuilder.IsSingleAgencySplitInTwo(list[list.Count - 2], list[list.Count - 1]))
			{
				headerData.IssuingAgency = list[list.Count - 2].Trim() + "\r\n" + list[list.Count - 1].Trim();
				headerData.SuperiorAgency = string.Join("\r\n", list.GetRange(0, list.Count - 2).ToArray()).Trim();
			}
			else
			{
				headerData.IssuingAgency = list[list.Count - 1].Trim();
				headerData.SuperiorAgency = string.Join("\r\n", list.GetRange(0, list.Count - 1).ToArray()).Trim();
			}
			headerData.HasExtractedAgency = true;
		}
		if (string.IsNullOrEmpty(headerData.IssuingAgency))
		{
			if (rulePack == RulePackType.HD05_VPTW)
			{
				headerData.SuperiorAgency = "TỈNH ỦY...";
				headerData.IssuingAgency = "ĐẢNG ỦY...";
				headerData.HasAsterisk = true;
			}
			else if (RulePackTraits.IsYouthUnion(rulePack))
			{
				headerData.SuperiorAgency = "TỈNH ĐOÀN...";
				headerData.IssuingAgency = "BCH ĐOÀN...";
				headerData.HasAsterisk = true;
			}
			else if (rulePack == RulePackType.TT68_2023_BQP)
			{
				headerData.SuperiorAgency = "BỘ QUỐC PHÒNG";
				headerData.IssuingAgency = "BỘ CHỈ HUY QUÂN SỰ...";
			}
			else
			{
				headerData.SuperiorAgency = "ỦY BAN NHÂN DÂN TỈNH...";
				headerData.IssuingAgency = "SỞ...";
			}
		}
		headerData.DocNumber = HeaderRebuilder.NormalizeDocNumber(text, rulePack);
		headerData.DateText = HeaderRebuilder.NormalizeDocDate(text2);
		return headerData;
	}

	public static SignatureData ExtractSignatureData(Document doc, Table existingTable, RulePackType rulePack)
	{
		SignatureData signatureData = new SignatureData();
		if (doc == null)
		{
			return signatureData;
		}
		List<string> list = new List<string>();
		List<string> list2 = new List<string>();
		if (existingTable != null)
		{
			try
			{
				int count = existingTable.Rows.Count;
				for (int i = 1; i <= count; i++)
				{
					Row row = existingTable.Rows[i];
					int count2 = row.Cells.Count;
					string[] array2;
					if (count2 >= 1)
					{
						string text = (row.Cells[1].Range.Text ?? "").Replace("\a", "").Trim();
						string[] array = text.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
						array2 = array;
						foreach (string text2 in array2)
						{
							string text3 = text2.Trim();
							if (text3.Length != 0)
							{
								string text4 = ConvertToUnsigned(text3).ToUpperInvariant();
								if (count2 == 1 && (text4.StartsWith("TM.") || text4.StartsWith("KT.") || text4.StartsWith("CHU TICH") || text4.Contains("CHU TICH") || text4.Contains("GIAM DOC") || (!text3.StartsWith("-") && !text3.StartsWith("+") && !text4.Contains("LUU") && !text4.Contains("NOI NHAN"))))
								{
									list2.Add(text3);
								}
								else if (!list.Contains(text3))
								{
									list.Add(text3);
								}
							}
						}
					}
					if (count2 < 2)
					{
						continue;
					}
					string text5 = (row.Cells[count2].Range.Text ?? "").Replace("\a", "").Trim();
					string[] array3 = text5.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
					array2 = array3;
					foreach (string text6 in array2)
					{
						string text7 = text6.Trim();
						if (text7.Length == 0)
						{
							continue;
						}
						string text8 = ConvertToUnsigned(text7).ToUpperInvariant();
						if (text7.StartsWith("-") || text7.StartsWith("+") || text8.StartsWith("NOI NHAN") || text8.StartsWith("LUU") || text8.Contains("LUU:") || text8.Contains("NHU TREN") || text8.Contains("NHU THANH PHAN"))
						{
							if (!list.Contains(text7))
							{
								list.Add(text7);
							}
						}
						else
						{
							list2.Add(text7);
						}
					}
				}
				bool flag = false;
				foreach (string item in list2)
				{
					string text9 = ConvertToUnsigned(item).ToUpperInvariant();
					if (!text9.StartsWith("TM.") && !text9.StartsWith("KT.") && !text9.StartsWith("TL.") && !text9.StartsWith("TUQ.") && !text9.Contains("CHU TICH") && !text9.Contains("BI THU") && !text9.Contains("GIAM DOC") && item.Length >= 4 && item.Length <= 40)
					{
						flag = true;
						break;
					}
				}
				if (!flag)
				{
					try
					{
						int end = existingTable.Range.End;
						for (int k = 1; k <= doc.Tables.Count; k++)
						{
							Table table = doc.Tables[k];
							if (table.Range.Start == existingTable.Range.Start || table.Range.Start < end || table.Range.Start > end + 150 || !SignatureRebuilder.IsLikelySignatureTable(doc, table))
							{
								continue;
							}
							for (int l = 1; l <= table.Rows.Count; l++)
							{
								Row row2 = table.Rows[l];
								int count3 = row2.Cells.Count;
								if (count3 >= 1)
								{
									string text10 = (row2.Cells[count3].Range.Text ?? "").Replace("\a", "").Trim();
									if (text10.Length > 0 && !list2.Contains(text10))
									{
										list2.Add(text10);
									}
								}
							}
							break;
						}
					}
					catch
					{
					}
					try
					{
						int end2 = existingTable.Range.End;
						int count4 = doc.Paragraphs.Count;
						for (int m = Math.Max(1, count4 - 6); m <= count4; m++)
						{
							Paragraph paragraph = doc.Paragraphs[m];
							if (paragraph.Range.Start >= end2 && paragraph.Range.Tables.Count == 0)
							{
								string text11 = (paragraph.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
									.Trim();
								if (text11.Length >= 4 && text11.Length <= 40 && !text11.Contains(":") && !text11.StartsWith("-") && !list2.Contains(text11))
								{
									list2.Add(text11);
								}
							}
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
		else
		{
			int num = SignatureRebuilder.LastContentParagraphIndex(doc);
			int num2 = Math.Max(1, num - 12);
			bool flag2 = false;
			for (int n = num2; n <= num; n++)
			{
				Paragraph paragraph2 = doc.Paragraphs[n];
				if (paragraph2.Range.Tables.Count != 0)
				{
					continue;
				}
				string text12 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "");
				string text13 = text12.Trim();
				if (string.IsNullOrEmpty(text13))
				{
					continue;
				}
				if (text13.Length > 120)
				{
					if (flag2)
					{
						break;
					}
					continue;
				}
				string text14 = ConvertToUnsigned(text13).ToUpperInvariant();
				if (text14.StartsWith("DIEU ") || text14.Contains("CHIU TRACH NHIEM") || text14.Contains("THI HANH"))
				{
					if (flag2)
					{
						break;
					}
				}
				else
				{
					if (!SignatureRebuilder.HasStrongSignatureAnchor(text12) && !flag2)
					{
						continue;
					}
					flag2 = true;
					bool flag3 = text12.StartsWith("\t") || text12.StartsWith("   ") || paragraph2.Format.Alignment == WdParagraphAlignment.wdAlignParagraphRight;
					string text15 = "";
					string text16 = "";
					string input = text12.Trim();
					string[] array4 = Regex.Split(input, "\\t+|\\s{2,}");
					List<string> list3 = new List<string>();
					string[] array2 = array4;
					foreach (string text17 in array2)
					{
						if (!string.IsNullOrEmpty(text17) && !string.IsNullOrEmpty(text17.Trim()))
						{
							list3.Add(text17.Trim());
						}
					}
					if (list3.Count >= 2 && (list3[0] == "-" || list3[0] == "+" || list3[0] == "*" || list3[0] == "–" || list3[0] == "—"))
					{
						list3[1] = list3[0] + " " + list3[1];
						list3.RemoveAt(0);
					}
					string[] array5 = list3.ToArray();
					if (array5.Length >= 2)
					{
						text15 = array5[0];
						text16 = array5[array5.Length - 1];
					}
					else if (array5.Length == 1)
					{
						string text18 = array5[0];
						string text19 = ConvertToUnsigned(text18).ToUpperInvariant();
						if (text19.Contains("NOI NHAN") || text18.StartsWith("-") || text18.StartsWith("+") || text19.Contains("LUU") || text19.Contains("NHU TREN") || text19.Contains("NHU THANH PHAN"))
						{
							text15 = text18;
						}
						else if (text19.Contains("TM.") || text19.Contains("KT.") || text19.Contains("TL.") || text19.Contains("TUQ.") || text19.Contains("CHU TICH") || text19.Contains("BI THU") || text19.Contains("GIAM DOC") || flag3)
						{
							text16 = text18;
						}
						else if (text18.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length <= 5 && !text18.StartsWith("-"))
						{
							text16 = text18;
						}
						else
						{
							text15 = text18;
						}
					}
					if (!string.IsNullOrEmpty(text16))
					{
						string text20 = ConvertToUnsigned(text16).ToUpperInvariant();
						if (text20.StartsWith("LUU") || text20.Contains("LUU:") || text20.Contains("LUU ") || text20.Contains("LUU.") || text20.Contains("NHU TREN") || text20.Contains("NHU THANH PHAN") || text20.Contains("NOI NHAN") || text16.StartsWith("-") || text16.StartsWith("+"))
						{
							text15 = ((!string.IsNullOrEmpty(text15)) ? (text15 + " " + text16) : text16);
							text16 = "";
						}
					}
					if (!string.IsNullOrEmpty(text16) && (text16.ToUpperInvariant().Contains("NOI NHAN") || text16.ToLowerInvariant().Contains("nơi nhận")))
					{
						text16 = Regex.Replace(text16, "(?i)nơi\\s*nhận\\s*:?\\s*", "").Trim();
						bool flag4 = false;
						foreach (string item2 in list)
						{
							if (item2.ToLowerInvariant().Contains("nơi nhận"))
							{
								flag4 = true;
								break;
							}
						}
						if (!flag4 && !text15.ToLowerInvariant().Contains("nơi nhận"))
						{
							text15 = ((!string.IsNullOrEmpty(text15)) ? ("Nơi nhận:\r\n" + text15) : "Nơi nhận:");
						}
					}
					if (!string.IsNullOrEmpty(text15) && !list.Contains(text15))
					{
						list.Add(text15);
					}
					if (!string.IsNullOrEmpty(text16))
					{
						list2.Add(text16);
					}
				}
			}
		}
		foreach (string item3 in list)
		{
			string text21 = item3.Trim();
			string text22 = ConvertToUnsigned(text21).ToUpperInvariant();
			if (text22.StartsWith("NOI NHAN"))
			{
				continue;
			}
			switch (text21)
			{
			case "-":
			case "+":
			case "–":
			case "—":
				continue;
			}
			text21 = ((!text21.StartsWith("-") && !text21.StartsWith("+") && !text21.StartsWith("–") && !text21.StartsWith("—") && !text21.StartsWith("*")) ? ("- " + text21) : Regex.Replace(text21, "^[-+*–—]\\s*", "- "));
			if (!signatureData.Recipients.Contains(text21))
			{
				signatureData.Recipients.Add(text21);
			}
		}
		if (signatureData.Recipients.Count == 0)
		{
			signatureData.Recipients.Add("- Như trên;");
			signatureData.Recipients.Add("- Lưu: VT.");
		}
		List<string> list4 = new List<string>();
		string text23 = string.Empty;
		for (int num3 = 0; num3 < list2.Count; num3++)
		{
			string text24 = list2[num3].Trim();
			string text25 = ConvertToUnsigned(text24).ToUpperInvariant();
			if (RegexPatterns.SignAuthority.IsMatch(text24) || text25.StartsWith("TM.") || text25.StartsWith("KT.") || text25.StartsWith("TL.") || text25.StartsWith("TUQ.") || text25.StartsWith("T/M") || text25.StartsWith("K/T"))
			{
				if (string.IsNullOrEmpty(signatureData.Authority))
				{
					signatureData.Authority = text24;
				}
			}
			else if (RegexPatterns.MilitaryRank.IsMatch(text24))
			{
				if (string.IsNullOrEmpty(signatureData.MilitaryRank))
				{
					Match match = RegexPatterns.MilitaryRank.Match(text24);
					if (match.Success)
					{
						signatureData.MilitaryRank = match.Value.Trim();
					}
				}
				if (text24.Length > 10 && string.IsNullOrEmpty(text23))
				{
					text23 = text24;
				}
			}
			else if (RegexPatterns.SignPosition.IsMatch(text24) || text25.Contains("CHU TICH") || text25.Contains("GIAM DOC") || text25.Contains("BI THU") || text25.Contains("TRUONG PHONG") || text25.Contains("TRUONG BAN") || text25.Contains("HIEU TRUONG") || text25.Contains("CHI HUY TRUONG") || text25.Contains("CHINH UY") || text25.Contains("CUC TRUONG") || text25.Contains("BO TRUONG") || text25.Contains("THU TRUONG") || text25.Contains("CHU NHIEM") || text25.Contains("TONG GIAM DOC") || text25.Contains("DAI DIEN"))
			{
				if (string.IsNullOrEmpty(signatureData.Title))
				{
					signatureData.Title = text24;
				}
			}
			else if (num3 == list2.Count - 1 && text24.Length >= 4 && text24.Length <= 40 && !text25.Contains("KY") && !text25.Contains("DONG DAU"))
			{
				text23 = text24;
			}
			else
			{
				list4.Add(text24);
			}
		}
		if (string.IsNullOrEmpty(signatureData.Title) && list4.Count > 0)
		{
			if (list4.Count == 1)
			{
				signatureData.Title = list4[0];
			}
			else
			{
				if (string.IsNullOrEmpty(signatureData.Authority))
				{
					signatureData.Authority = list4[0];
				}
				signatureData.Title = list4[list4.Count - 1];
			}
		}
		SignatureData signatureData2 = signatureData;
		string signerName;
		if (string.IsNullOrEmpty(text23))
		{
			signerName = ((rulePack == RulePackType.HD05_VPTW || RulePackTraits.IsYouthUnion(rulePack)) ? "Nguyễn Văn A" : "Trần Văn B");
		}
		else
		{
			signerName = text23;
		}
		signatureData2.SignerName = signerName;
		if (string.IsNullOrEmpty(signatureData.Title))
		{
			if (rulePack == RulePackType.HD05_VPTW)
			{
				if (string.IsNullOrEmpty(signatureData.Authority))
				{
					signatureData.Authority = "T/M BAN CHẤP HÀNH";
				}
				signatureData.Title = "BÍ THƯ";
			}
			else if (RulePackTraits.IsYouthUnion(rulePack))
			{
				if (string.IsNullOrEmpty(signatureData.Authority))
				{
					signatureData.Authority = "TM. BAN THƯỜNG VỤ";
				}
				signatureData.Title = "BÍ THƯ";
			}
			else
			{
				if (string.IsNullOrEmpty(signatureData.Authority))
				{
					signatureData.Authority = "TM. ỦY BAN NHÂN DÂN";
				}
				signatureData.Title = "CHỦ TỊCH";
			}
		}
		return signatureData;
	}

	public static TitleAbstractData ExtractTitleAbstractData(Document doc)
	{
		TitleAbstractData titleAbstractData = new TitleAbstractData();
		if (doc == null)
		{
			return titleAbstractData;
		}
		int num = Math.Min(doc.Paragraphs.Count, 25);
		int num2 = -1;
		for (int i = 1; i <= num; i++)
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
				string text2 = ConvertToUnsigned(text).ToUpperInvariant();
				if (HeaderRebuilder.IsDocumentTitle(text))
				{
					titleAbstractData.DocumentType = text.ToUpperInvariant();
					titleAbstractData.IsDispatch = false;
					num2 = i;
					break;
				}
				if (RegexPatterns.DispatchSubject.IsMatch(text) || text2.StartsWith("V/V ") || text2.StartsWith("V/V:"))
				{
					titleAbstractData.IsDispatch = true;
					titleAbstractData.SubjectText = text;
					return titleAbstractData;
				}
			}
		}
		if (num2 > 0)
		{
			List<string> list = new List<string>();
			int num3 = Math.Min(doc.Paragraphs.Count, num2 + 5);
			for (int j = num2 + 1; j <= num3; j++)
			{
				Paragraph paragraph2 = doc.Paragraphs[j];
				if (paragraph2.Range.Tables.Count > 0)
				{
					break;
				}
				string text3 = (paragraph2.Range.Text ?? "").Replace("\r", "").Replace("\n", "").Replace("\a", "")
					.Trim();
				if (!string.IsNullOrEmpty(text3) && !HeaderRebuilder.IsSeparatorLine(text3))
				{
					if (RegexPatterns.LegalGround.IsMatch(text3) || text3.StartsWith("Căn cứ ") || text3.StartsWith("Điều ") || text3.StartsWith("I.") || text3.StartsWith("1."))
					{
						break;
					}
					list.Add(text3);
				}
			}
			titleAbstractData.SubjectText = string.Join("\r\n", list.ToArray());
		}
		return titleAbstractData;
	}
}
