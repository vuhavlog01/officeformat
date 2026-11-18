using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Interfaces;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Serialization;

namespace VFormat.Core.Rules;

public class RuleEngine : IRuleEngine
{
	private ProfileModel _activeProfile;

	private UserPersonalization _personalization;

	public ProfileModel ActiveProfile => _activeProfile;

	public UserPersonalization Personalization => _personalization;

	public RuleEngine()
	{
		try
		{
			JsonProfileRepository jsonProfileRepository = new JsonProfileRepository();
			_personalization = jsonProfileRepository.LoadPersonalization() ?? UserPersonalization.CreateDefault();
		}
		catch
		{
			_personalization = UserPersonalization.CreateDefault();
		}
		_activeProfile = ND30Rules.CreateProfile();
		ApplyPersonalizationToProfile();
	}

	public RuleEngine(ProfileModel profile, UserPersonalization personalization)
	{
		_activeProfile = profile ?? ND30Rules.CreateProfile();
		_personalization = personalization ?? UserPersonalization.CreateDefault();
		ApplyPersonalizationToProfile();
	}

	public void SetActiveProfile(string profileId)
	{
		if (!string.IsNullOrEmpty(profileId))
		{
			string text = profileId.ToUpperInvariant();
			if (text.Contains("DOAN") || text.Contains("HD29") || text.Contains("TWDTN") || text.Contains("YOUTH"))
			{
				_activeProfile = HD29Rules.CreateProfile();
			}
			else if (text.Contains("DANG") || text.Contains("HD05") || text.Contains("05HD"))
			{
				_activeProfile = HD05Rules.CreateProfile();
			}
			else if (text.Contains("TT68") || text.Contains("BQP") || text.Contains("MILITARY"))
			{
				_activeProfile = TT68Rules.CreateProfile();
			}
			else
			{
				_activeProfile = ND30Rules.CreateProfile();
			}
			_personalization.ActiveRulePack = _activeProfile.RulePack;
			ApplyPersonalizationToProfile();
		}
	}

	public void SetActiveProfile(RulePackType rulePack)
	{
		switch (rulePack)
		{
		case RulePackType.HD05_VPTW:
			_activeProfile = HD05Rules.CreateProfile();
			break;
		case RulePackType.TT68_2023_BQP:
			_activeProfile = TT68Rules.CreateProfile();
			break;
		case RulePackType.HD29_TWDTN:
			_activeProfile = HD29Rules.CreateProfile();
			break;
		default:
			_activeProfile = ND30Rules.CreateProfile();
			break;
		}
		_personalization.ActiveRulePack = rulePack;
		ApplyPersonalizationToProfile();
	}

	public void UpdatePersonalization(UserPersonalization personalization)
	{
		if (personalization != null)
		{
			_personalization = personalization;
			if (_activeProfile != null)
			{
				SetActiveProfile(_activeProfile.RulePack);
			}
		}
	}

	private void ApplyPersonalizationToProfile()
	{
		if (_activeProfile == null || _personalization == null)
		{
			return;
		}
		_activeProfile.HeaderDateInTopCell = _personalization.HeaderDateInTopCell;
		MarginProfile activeMargins = _personalization.GetActiveMargins(_activeProfile.RulePack);
		TypographyProfile activeTypography = _personalization.GetActiveTypography(_activeProfile.RulePack);
		if (activeTypography != null && activeTypography.SizeThanBai > 0f)
		{
			_activeProfile.Typography.BodySizePt = activeTypography.SizeThanBai;
			if (_activeProfile.TablePreset != null)
			{
				_activeProfile.TablePreset.HeaderFontSizePt = activeTypography.SizeThanBai;
				_activeProfile.TablePreset.BodyFontSizePt = activeTypography.SizeThanBai;
			}
		}
		if (activeMargins != null)
		{
			if (TryParseLineSpacing(activeMargins.LineSpacingRule, activeMargins.LineSpacingAtPt, out var rule, out var amount))
			{
				_activeProfile.Typography.LineSpacingRule = rule;
				_activeProfile.Typography.LineSpacingValue = amount;
			}
			_activeProfile.Typography.SpacingBeforePt = activeMargins.SpacingBeforePt;
			_activeProfile.Typography.SpacingAfterPt = activeMargins.SpacingAfterPt;
			_activeProfile.Typography.NoSpaceBetweenParagraphsOfSameStyle = activeMargins.NoSpaceBetweenParagraphsOfSameStyle;
		}
		if (activeMargins != null && _activeProfile.PageSetup != null)
		{
			if (activeMargins.MarginTopMm > 0f)
			{
				_activeProfile.PageSetup.MarginTopMm = activeMargins.MarginTopMm;
			}
			if (activeMargins.MarginBottomMm > 0f)
			{
				_activeProfile.PageSetup.MarginBottomMm = activeMargins.MarginBottomMm;
			}
			if (activeMargins.MarginLeftMm > 0f)
			{
				_activeProfile.PageSetup.MarginLeftMm = activeMargins.MarginLeftMm;
			}
			if (activeMargins.MarginRightMm > 0f)
			{
				_activeProfile.PageSetup.MarginRightMm = activeMargins.MarginRightMm;
			}
		}
		SemanticRole[] array = new SemanticRole[4]
		{
			SemanticRole.BodyText,
			SemanticRole.ClauseText,
			SemanticRole.PointText,
			SemanticRole.LegalGround
		};
		SemanticRole[] array2 = array;
		foreach (SemanticRole semanticRole in array2)
		{
			ComponentStyle componentStyle = _activeProfile.GetComponentStyle(semanticRole);
			if (componentStyle == null)
			{
				continue;
			}
			componentStyle.LineSpacingRule = _activeProfile.Typography.LineSpacingRule;
			componentStyle.LineSpacingValue = _activeProfile.Typography.LineSpacingValue;
			componentStyle.SpacingBeforePt = _activeProfile.Typography.SpacingBeforePt;
			componentStyle.SpacingAfterPt = _activeProfile.Typography.SpacingAfterPt;
			componentStyle.NoSpaceBetweenParagraphsOfSameStyle = _activeProfile.Typography.NoSpaceBetweenParagraphsOfSameStyle;
			if (activeMargins != null && _personalization.UseCustomParagraph)
			{
				float num = activeMargins.FirstLineIndentCm;
				if (num >= 50f)
				{
					num /= 100f;
				}
				if (num >= 0f)
				{
					componentStyle.FirstLineIndentMm = num * 10f;
				}
				if (activeMargins.SpacingBeforePt >= 0f)
				{
					componentStyle.SpacingBeforePt = activeMargins.SpacingBeforePt;
				}
				if (activeMargins.SpacingAfterPt >= 0f)
				{
					componentStyle.SpacingAfterPt = activeMargins.SpacingAfterPt;
				}
				if (semanticRole == SemanticRole.BodyText)
				{
					if (componentStyle.FirstLineIndentMm.HasValue)
					{
						_activeProfile.Typography.FirstLineIndentMm = componentStyle.FirstLineIndentMm.Value;
					}
					if (componentStyle.SpacingBeforePt.HasValue)
					{
						_activeProfile.Typography.SpacingBeforePt = componentStyle.SpacingBeforePt.Value;
					}
					if (componentStyle.SpacingAfterPt.HasValue)
					{
						_activeProfile.Typography.SpacingAfterPt = componentStyle.SpacingAfterPt.Value;
					}
				}
			}
			if (activeTypography != null && activeTypography.SizeThanBai > 0f)
			{
				componentStyle.SizePt = activeTypography.SizeThanBai;
				if (semanticRole != SemanticRole.LegalGround)
				{
					componentStyle.Bold = activeTypography.BoldThanBai;
					componentStyle.Italic = activeTypography.ItalicThanBai;
				}
			}
		}
		if (activeTypography != null && activeTypography.SizeThanBai > 0f)
		{
			SemanticRole[] array3 = new SemanticRole[4]
			{
				SemanticRole.PartHeader,
				SemanticRole.ChapterHeader,
				SemanticRole.SectionHeader,
				SemanticRole.ArticleTitle
			};
			foreach (SemanticRole semanticRole2 in array3)
			{
				if (semanticRole2 == SemanticRole.PartHeader && _activeProfile.RulePack == RulePackType.HD29_TWDTN)
				{
					ComponentStyle componentStyle2 = _activeProfile.GetComponentStyle(semanticRole2);
					componentStyle2.SizePt = LegalFontSizeStandards.EnsureLegalInteger(RulePackType.HD29_TWDTN, semanticRole2, Math.Max(activeTypography.SizeThanBai + 1f, componentStyle2.SizePt));
				}
				else
				{
					_activeProfile.GetComponentStyle(semanticRole2).SizePt = activeTypography.SizeThanBai;
				}
			}
		}
		if (activeTypography == null || _activeProfile.ComponentStyles == null)
		{
			return;
		}
		if (activeTypography.SizeQuocHieu > 0f)
		{
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.NationalHeader))
			{
				_activeProfile.ComponentStyles[SemanticRole.NationalHeader].SizePt = activeTypography.SizeQuocHieu;
				_activeProfile.ComponentStyles[SemanticRole.NationalHeader].Bold = activeTypography.BoldQuocHieu;
				_activeProfile.ComponentStyles[SemanticRole.NationalHeader].Italic = activeTypography.ItalicQuocHieu;
			}
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.PartyHeader))
			{
				_activeProfile.ComponentStyles[SemanticRole.PartyHeader].SizePt = activeTypography.SizeQuocHieu;
				_activeProfile.ComponentStyles[SemanticRole.PartyHeader].Bold = activeTypography.BoldQuocHieu;
				_activeProfile.ComponentStyles[SemanticRole.PartyHeader].Italic = activeTypography.ItalicQuocHieu;
			}
		}
		if (activeTypography.SizeTieuNgu > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.Motto))
		{
			_activeProfile.ComponentStyles[SemanticRole.Motto].SizePt = activeTypography.SizeTieuNgu;
			_activeProfile.ComponentStyles[SemanticRole.Motto].Bold = activeTypography.BoldTieuNgu;
			_activeProfile.ComponentStyles[SemanticRole.Motto].Italic = activeTypography.ItalicTieuNgu;
		}
		if (activeTypography.SizeCoQuan > 0f)
		{
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.IssuingAgency))
			{
				_activeProfile.ComponentStyles[SemanticRole.IssuingAgency].SizePt = activeTypography.SizeCoQuan;
				_activeProfile.ComponentStyles[SemanticRole.IssuingAgency].Bold = activeTypography.BoldCoQuan;
				_activeProfile.ComponentStyles[SemanticRole.IssuingAgency].Italic = activeTypography.ItalicCoQuan;
			}
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.SuperiorAgency))
			{
				_activeProfile.ComponentStyles[SemanticRole.SuperiorAgency].SizePt = activeTypography.SizeCoQuan;
			}
		}
		if (activeTypography.SizeSoKyHieu > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.DocumentNumber))
		{
			_activeProfile.ComponentStyles[SemanticRole.DocumentNumber].SizePt = activeTypography.SizeSoKyHieu;
			_activeProfile.ComponentStyles[SemanticRole.DocumentNumber].Bold = activeTypography.BoldSoKyHieu;
			_activeProfile.ComponentStyles[SemanticRole.DocumentNumber].Italic = activeTypography.ItalicSoKyHieu;
		}
		if (activeTypography.SizeDiaDanh > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.DocumentDate))
		{
			_activeProfile.ComponentStyles[SemanticRole.DocumentDate].SizePt = activeTypography.SizeDiaDanh;
			_activeProfile.ComponentStyles[SemanticRole.DocumentDate].Bold = activeTypography.BoldDiaDanh;
			_activeProfile.ComponentStyles[SemanticRole.DocumentDate].Italic = activeTypography.ItalicDiaDanh;
		}
		if (activeTypography.SizeTenLoai > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.DocumentType))
		{
			_activeProfile.ComponentStyles[SemanticRole.DocumentType].SizePt = activeTypography.SizeTenLoai;
			_activeProfile.ComponentStyles[SemanticRole.DocumentType].Bold = activeTypography.BoldTenLoai;
			_activeProfile.ComponentStyles[SemanticRole.DocumentType].Italic = activeTypography.ItalicTenLoai;
		}
		if (activeTypography.SizeTrichYeu > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.SubjectExcerpt))
		{
			_activeProfile.ComponentStyles[SemanticRole.SubjectExcerpt].SizePt = activeTypography.SizeTrichYeu;
			_activeProfile.ComponentStyles[SemanticRole.SubjectExcerpt].Bold = activeTypography.BoldTrichYeu;
			_activeProfile.ComponentStyles[SemanticRole.SubjectExcerpt].Italic = activeTypography.ItalicTrichYeu;
		}
		if (activeTypography.SizeTrichYeuCongVan > 0f)
		{
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.DispatchSubject))
			{
				_activeProfile.ComponentStyles[SemanticRole.DispatchSubject].SizePt = activeTypography.SizeTrichYeuCongVan;
				_activeProfile.ComponentStyles[SemanticRole.DispatchSubject].Bold = activeTypography.BoldTrichYeuCongVan;
				_activeProfile.ComponentStyles[SemanticRole.DispatchSubject].Italic = activeTypography.ItalicTrichYeuCongVan;
			}
			else
			{
				_activeProfile.ComponentStyles[SemanticRole.DispatchSubject] = new ComponentStyle(SemanticRole.DispatchSubject, activeTypography.SizeTrichYeuCongVan, activeTypography.BoldTrichYeuCongVan, activeTypography.ItalicTrichYeuCongVan, AlignmentType.Center);
			}
		}
		if (activeTypography.SizeThanBai > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.BodyText))
		{
			_activeProfile.ComponentStyles[SemanticRole.BodyText].SizePt = activeTypography.SizeThanBai;
			_activeProfile.ComponentStyles[SemanticRole.BodyText].Bold = activeTypography.BoldThanBai;
			_activeProfile.ComponentStyles[SemanticRole.BodyText].Italic = activeTypography.ItalicThanBai;
		}
		if (activeTypography.SizeNoiNhan > 0f && _activeProfile.ComponentStyles.ContainsKey(SemanticRole.RecipientItem))
		{
			_activeProfile.ComponentStyles[SemanticRole.RecipientItem].SizePt = activeTypography.SizeNoiNhan;
			_activeProfile.ComponentStyles[SemanticRole.RecipientItem].Bold = activeTypography.BoldNoiNhan;
			_activeProfile.ComponentStyles[SemanticRole.RecipientItem].Italic = activeTypography.ItalicNoiNhan;
		}
		if (activeTypography.SizeChuKy > 0f)
		{
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.SignerName))
			{
				_activeProfile.ComponentStyles[SemanticRole.SignerName].SizePt = activeTypography.SizeChuKy;
				_activeProfile.ComponentStyles[SemanticRole.SignerName].Bold = activeTypography.BoldChuKy;
				_activeProfile.ComponentStyles[SemanticRole.SignerName].Italic = activeTypography.ItalicChuKy;
			}
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.SignRankName))
			{
				_activeProfile.ComponentStyles[SemanticRole.SignRankName].SizePt = activeTypography.SizeChuKy;
				_activeProfile.ComponentStyles[SemanticRole.SignRankName].Bold = activeTypography.BoldChuKy;
				_activeProfile.ComponentStyles[SemanticRole.SignRankName].Italic = activeTypography.ItalicChuKy;
			}
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.SignAuthority))
			{
				_activeProfile.ComponentStyles[SemanticRole.SignAuthority].SizePt = activeTypography.SizeChuKy;
			}
			if (_activeProfile.ComponentStyles.ContainsKey(SemanticRole.SignTitle))
			{
				_activeProfile.ComponentStyles[SemanticRole.SignTitle].SizePt = activeTypography.SizeChuKy;
			}
		}
	}

	private static bool TryParseLineSpacing(string value, out LineSpacingType rule, out float amount)
	{
		return TryParseLineSpacing(value, 0f, out rule, out amount);
	}

	private static bool TryParseLineSpacing(string value, float atPt, out LineSpacingType rule, out float amount)
	{
		rule = LineSpacingType.Single;
		amount = 1f;
		if (string.IsNullOrWhiteSpace(value))
		{
			return false;
		}
		string text = value.Trim();
		if (text.Equals("Exactly", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.Exactly;
			amount = (atPt > 0f) ? atPt : 18f;
			return true;
		}
		if (text.Equals("At least", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.AtLeast;
			amount = (atPt > 0f) ? atPt : 18f;
			return true;
		}
		if (text.Equals("Single", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.Single;
			amount = 1f;
			return true;
		}
		if (text.Equals("1.5 lines", StringComparison.OrdinalIgnoreCase) || text.Equals("1.5", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.OneAndHalf;
			amount = 1.5f;
			return true;
		}
		if (text.Equals("Double", StringComparison.OrdinalIgnoreCase) || text.Equals("2", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.Double;
			amount = 2f;
			return true;
		}
		if (text.Equals("Multiple", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.Multiple;
			amount = (atPt > 0f) ? atPt : 1.15f;
			return true;
		}
		Match match = Regex.Match(text, "^(?:(Exactly)\\s+)?(\\d+(?:[.,]\\d+)?)\\s*(pt|lines?)?$", RegexOptions.IgnoreCase);
		if (!match.Success || !float.TryParse(match.Groups[2].Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out amount) || amount <= 0f || float.IsInfinity(amount))
		{
			return false;
		}
		if (match.Groups[1].Success || match.Groups[3].Value.Equals("pt", StringComparison.OrdinalIgnoreCase))
		{
			rule = LineSpacingType.Exactly;
		}
		else
		{
			rule = ((amount != 1f) ? LineSpacingType.Multiple : LineSpacingType.Single);
		}
		return true;
	}

	public RulePackType DetectRulePack(string sampleText)
	{
		return DetectRulePackDetailed(sampleText).RulePack;
	}

	public RulePackDetectionResult DetectRulePackDetailed(string sampleText)
	{
		RulePackDetectionResult rulePackDetectionResult = new RulePackDetectionResult();
		if (string.IsNullOrEmpty(sampleText))
		{
			rulePackDetectionResult.Warnings.Add("Văn bản không có nội dung để nhận diện.");
			return rulePackDetectionResult;
		}
		string[] array = sampleText.Split(new char[4] { '\r', '\n', '\a', '\v' }, StringSplitOptions.RemoveEmptyEntries);
		if (array.Length == 0)
		{
			rulePackDetectionResult.Warnings.Add("Không tìm thấy dòng văn bản hợp lệ.");
			return rulePackDetectionResult;
		}
		List<string> list = new List<string>();
		List<string> list2 = new List<string>();
		bool flag = false;
		for (int i = 0; i < array.Length; i++)
		{
			string text = array[i].Trim();
			if (string.IsNullOrEmpty(text))
			{
				continue;
			}
			if (!flag)
			{
				if (RegexPatterns.DocumentType.IsMatch(text) || IsDocumentTitleLine(text) || list.Count >= 20)
				{
					flag = true;
					list2.Add(text);
				}
				else
				{
					list.Add(text);
				}
			}
			else
			{
				list2.Add(text);
			}
		}
		string text2 = string.Join("\n", list);
		string text3 = ConvertToUnsigned(text2).ToUpperInvariant();
		string text4 = ConvertToUnsigned(sampleText).ToUpperInvariant();
		bool flag2 = RegexPatterns.NationalHeader.IsMatch(text2) || RegexPatterns.Motto.IsMatch(text2) || text3.Contains("CONG HOA XA HOI CHU NGHIA VIET NAM") || text3.Contains("DOC LAP - TU DO");
		bool flag3 = RegexPatterns.PartyHeader.IsMatch(text2) || text3.Contains("DANG CONG SAN VIET NAM");
		bool flag4 = RegexPatterns.MilitaryAgency.IsMatch(text2) || text3.Contains("BO QUOC PHONG") || text3.Contains("BO CHQS") || text3.Contains("BAN CHQS") || text3.Contains("BAN CHI HUY QUAN SU") || text3.Contains("BO CHI HUY QUAN SU") || text3.Contains("BCHQS") || text3.Contains("BCH QUAN SU") || text3.Contains("CHQS") || text3.Contains("QUAN KHU") || text3.Contains("QUAN DOAN") || text3.Contains("QUAN CHUNG") || text3.Contains("BO TU LENH") || text3.Contains("BO DOI BIEN PHONG") || text3.Contains("BCHBDBP") || text3.Contains("BDBP") || text3.Contains("DON BIEN PHONG") || text3.Contains("DAN QUAN TU VE") || text3.Contains("DQTV") || text3.Contains("LLVT") || text3.Contains("LUC LUONG VU TRANG");
		bool flag5 = RegexPatterns.MilitaryNumberSymbol.IsMatch(text2) || text3.Contains("-BCH") || text3.Contains("/KH-BCH") || text3.Contains("/QD-BCH") || text3.Contains("/CV-BCH") || text3.Contains("-BQP");
		bool flag6 = RegexPatterns.YouthUnionHeader.IsMatch(text2) || text3.Contains("DOAN TNCS HO CHI MINH") || text3.Contains("DOAN THANH NIEN CONG SAN HO CHI MINH");
		bool flag7 = RegexPatterns.YouthUnionAgency.IsMatch(text2);
		bool flag8 = RegexPatterns.YouthUnionNumberSymbol.IsMatch(text2);
		bool flag9 = false;
		foreach (string item in list)
		{
			string text5 = item.Replace(" ", "").Replace("\t", "");
			if (text5 == "***")
			{
				flag9 = true;
				break;
			}
		}
		bool flag10 = flag6 || flag7 || flag8;
		bool flag11 = RegexPatterns.PartyAgency.IsMatch(text2) || text3.Contains("DANG BO") || text3.Contains("CHI BO") || text3.Contains("CHI UY") || text3.Contains("DANG UY") || text3.Contains("TINH UY") || text3.Contains("HUYEN UY") || text3.Contains("THANH UY") || text3.Contains("QUAN UY") || text3.Contains("THI UY") || text3.Contains("BAN BI THU") || text3.Contains("BO CHINH TRI") || text3.Contains("BAN TO CHUC") || text3.Contains("BAN TUYEN GIAO") || text3.Contains("UY BAN KIEM TRA") || text3.Contains("VAN PHONG TRUNG UONG DANG");
		if (flag10 && !flag3 && !RegexPatterns.PartyHeader.IsMatch(text2) && !text3.Contains("DANG UY") && !text3.Contains("DANG BO") && !text3.Contains("CHI BO") && !text3.Contains("TINH UY") && !text3.Contains("HUYEN UY") && !text3.Contains("THANH UY") && !text3.Contains("QUAN UY") && !text3.Contains("THI UY"))
		{
			flag11 = false;
		}
		bool flag12 = (!flag10 || flag2) && (RegexPatterns.AdminAgency.IsMatch(text2) || text3.Contains("UY BAN NHAN DAN") || text3.Contains("UBND") || text3.Contains("HOI DONG NHAN DAN") || text3.Contains("HDND") || text3.Contains("CHINH PHU") || text3.Contains("THU TUONG") || (text3.Contains("BO ") && !text3.Contains("BO QUOC PHONG")) || text3.Contains("SO ") || text3.Contains("CUC ")) && !flag4;
		if (flag12 && !text3.Contains("BO QUOC PHONG") && !text3.Contains("BCHQS") && !text3.Contains("BO CHQS") && !text3.Contains("QUAN KHU") && !text3.Contains("QUAN DOAN"))
		{
			flag4 = false;
			flag5 = false;
		}
		bool flag13 = flag4 || flag5 || RegexPatterns.MilitaryRank.IsMatch(sampleText) || RegexPatterns.MilitaryAgency.IsMatch(sampleText) || text4.Contains("CHI HUY TRUONG") || text4.Contains("CHINH TRI VIEN") || text4.Contains("THAM MUU TRUONG") || text4.Contains("BO CHQS") || text4.Contains("BAN CHQS") || text4.Contains("BAN CHI HUY QUAN SU");
		int num = 0;
		int num2 = 0;
		int num3 = 0;
		int num4 = 0;
		if (flag2 && flag10 && !flag3)
		{
			rulePackDetectionResult.Warnings.Add("Khối đầu có Quốc hiệu và tên cơ quan Đoàn: theo HD 29 (Phần II.10) văn bản dùng dấu Nhà nước hoặc liên tịch trình bày theo cơ quan chủ trì; giữ thể thức Nhà nước, hãy chọn lại quy chuẩn nếu cần.");
		}
		if (flag3 && flag6)
		{
			rulePackDetectionResult.RulePack = RulePackType.HD05_VPTW;
			rulePackDetectionResult.Confidence = 45;
			rulePackDetectionResult.IsConfident = false;
			rulePackDetectionResult.Evidence.Add("có đồng thời tiêu đề Đảng và tiêu đề Đoàn ở khối đầu");
			rulePackDetectionResult.Warnings.Add("Văn bản liên tịch Đảng - Đoàn: trình bày theo cơ quan chủ trì. Đang chọn HD 05; nếu Đoàn chủ trì, hãy chọn \"VB Đoàn (HD 29)\".");
			return rulePackDetectionResult;
		}
		if (flag2 && (flag3 || flag11))
		{
			rulePackDetectionResult.RulePack = RulePackType.ND30_2020;
			rulePackDetectionResult.Confidence = 25;
			rulePackDetectionResult.IsConfident = false;
			rulePackDetectionResult.Evidence.Add("có đồng thời Quốc hiệu Nhà nước và dấu hiệu Đảng ở khối đầu");
			rulePackDetectionResult.Warnings.Add("Khối đầu có dấu hiệu thể thức mâu thuẫn; chạy chế độ thận trọng, không áp dụng quy chuẩn Đảng tự động.");
			return rulePackDetectionResult;
		}
		if (flag2)
		{
			num += 5;
			rulePackDetectionResult.Evidence.Add("có Quốc hiệu/Tiêu ngữ Nhà nước ở đầu văn bản");
		}
		if (flag12)
		{
			num += 8;
			rulePackDetectionResult.Evidence.Add("có cơ quan hành chính Nhà nước trong khối đầu");
		}
		if (flag3)
		{
			num2 += 10;
			rulePackDetectionResult.Evidence.Add("có tiêu đề Đảng ở đầu văn bản");
		}
		if (flag11)
		{
			num2 += 6;
			rulePackDetectionResult.Evidence.Add("có cơ quan Đảng trong khối đầu");
		}
		if (flag4)
		{
			num3 += 10;
			rulePackDetectionResult.Evidence.Add("có cơ quan quân sự trong khối đầu");
		}
		if (flag5)
		{
			num3 += 5;
			rulePackDetectionResult.Evidence.Add("có số/ký hiệu quân sự trong khối đầu");
		}
		if (flag6)
		{
			num4 += 10;
			rulePackDetectionResult.Evidence.Add("có tiêu đề \"ĐOÀN TNCS HỒ CHÍ MINH\" ở đầu văn bản");
		}
		if (flag7)
		{
			num4 += 6;
			rulePackDetectionResult.Evidence.Add("có cơ quan Đoàn trong khối đầu");
		}
		if (flag8)
		{
			num4 += 5;
			rulePackDetectionResult.Evidence.Add("có ký hiệu số văn bản của Đoàn (TWĐTN, TĐTN, ĐTN...)");
		}
		if (flag9 && flag10)
		{
			num4 += 2;
			rulePackDetectionResult.Evidence.Add("có dòng \"***\" dưới tên cơ quan");
		}
		if (flag2)
		{
			num4 = Math.Min(num4, 4);
		}
		if (flag13 && !flag4 && !flag5)
		{
			num3++;
			rulePackDetectionResult.Warnings.Add("Dấu hiệu quân sự chỉ xuất hiện ngoài khối đầu; không dùng làm căn cứ mạnh.");
		}
		if (flag2 && flag12 && !text3.Contains("BO QUOC PHONG") && !text3.Contains("BCHQS") && !text3.Contains("BO CHQS"))
		{
			num3 = 0;
		}
		int num5 = Math.Max(Math.Max(num, num4), Math.Max(num2, num3));
		int num6 = num + num2 + num3 + num4 - num5;
		if (num4 >= num5 && num4 > num && num4 > num2 && num4 > num3)
		{
			rulePackDetectionResult.RulePack = RulePackType.HD29_TWDTN;
		}
		else if (num2 >= num5 && num2 > num && num2 > num3)
		{
			rulePackDetectionResult.RulePack = RulePackType.HD05_VPTW;
		}
		else if (num3 >= num5 && num3 > num && num3 > num2)
		{
			rulePackDetectionResult.RulePack = RulePackType.TT68_2023_BQP;
		}
		else
		{
			rulePackDetectionResult.RulePack = RulePackType.ND30_2020;
		}
		rulePackDetectionResult.Confidence = ((num5 != 0) ? Math.Min(100, 55 + num5 * 5 - num6 * 4) : 0);
		rulePackDetectionResult.IsConfident = num5 >= 6 && num5 - num6 >= 3;
		if (!rulePackDetectionResult.IsConfident)
		{
			rulePackDetectionResult.Warnings.Add("Các dấu hiệu thể thức chưa đủ rõ hoặc đang mâu thuẫn; cần người dùng chọn quy chuẩn.");
		}
		return rulePackDetectionResult;
	}

	private static bool IsDocumentTitleLine(string line)
	{
		if (string.IsNullOrEmpty(line))
		{
			return false;
		}
		string text = ConvertToUnsigned(line).Trim().ToUpperInvariant();
		if (!(text == "QUYET DINH") && !text.StartsWith("QUYET DINH ") && !(text == "NGHI QUYET") && !text.StartsWith("NGHI QUYET ") && !(text == "TO TRINH") && !text.StartsWith("TO TRINH ") && !(text == "KE HOACH") && !text.StartsWith("KE HOACH ") && !(text == "BAO CAO") && !text.StartsWith("BAO CAO ") && !(text == "THONG BAO") && !text.StartsWith("THONG BAO ") && !(text == "CHI THI") && !text.StartsWith("CHI THI ") && !(text == "THONG TRI") && !text.StartsWith("THONG TRI ") && !(text == "QUY DINH") && !text.StartsWith("QUY DINH ") && !(text == "HUONG DAN") && !text.StartsWith("HUONG DAN ") && !(text == "CONG VAN") && !text.StartsWith("CONG VAN ") && !(text == "CONG DIEN") && !text.StartsWith("CONG DIEN ") && !(text == "GIAY MOI") && !text.StartsWith("GIAY MOI ") && !(text == "BIEN BAN") && !text.StartsWith("BIEN BAN ") && !(text == "GIAY GIOI THIEU") && !text.StartsWith("GIAY GIOI THIEU ") && !(text == "BAN DANG KY") && !text.StartsWith("BAN DANG KY ") && !(text == "BAN CAM KET") && !text.StartsWith("BAN CAM KET ") && !(text == "BAN KIEM DIEM") && !text.StartsWith("BAN KIEM DIEM ") && !(text == "BAN TU KIEM DIEM") && !text.StartsWith("BAN TU KIEM DIEM ") && !(text == "CHUONG TRINH") && !text.StartsWith("CHUONG TRINH ") && !(text == "DE AN"))
		{
			return text.StartsWith("DE AN ");
		}
		return true;
	}

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

	public SemanticRole ClassifyParagraph(string text, int paragraphIndex, int totalParagraphs, bool isInTable)
	{
		if (string.IsNullOrEmpty(text))
		{
			return SemanticRole.Unknown;
		}
		string text2 = text.Trim();
		if (text2.Length == 0)
		{
			return SemanticRole.Unknown;
		}
		if (isInTable)
		{
			if (RegexPatterns.NationalHeader.IsMatch(text2))
			{
				return SemanticRole.NationalHeader;
			}
			if (RegexPatterns.Motto.IsMatch(text2))
			{
				return SemanticRole.Motto;
			}
			if (RegexPatterns.PartyHeader.IsMatch(text2))
			{
				return SemanticRole.PartyHeader;
			}
			if (RegexPatterns.YouthUnionHeader.IsMatch(text2))
			{
				return SemanticRole.PartyHeader;
			}
			if (RegexPatterns.AdminNumber.IsMatch(text2) || RegexPatterns.DocumentNumberGeneral.IsMatch(text2))
			{
				return SemanticRole.DocumentNumber;
			}
			if (RegexPatterns.DocumentDate.IsMatch(text2))
			{
				return SemanticRole.DocumentDate;
			}
			if (RegexPatterns.RecipientHeader.IsMatch(text2))
			{
				return SemanticRole.RecipientHeader;
			}
			if (RegexPatterns.RecipientItem.IsMatch(text2))
			{
				return SemanticRole.RecipientItem;
			}
			if (RegexPatterns.SignAuthority.IsMatch(text2))
			{
				return SemanticRole.SignAuthority;
			}
			if (RegexPatterns.SignPosition.IsMatch(text2))
			{
				return SemanticRole.SignTitle;
			}
			if (_activeProfile.RulePack == RulePackType.TT68_2023_BQP && RegexPatterns.MilitaryRank.IsMatch(text2))
			{
				return SemanticRole.SignRankName;
			}
			return SemanticRole.TableBody;
		}
		if (paragraphIndex <= 15)
		{
			if (RegexPatterns.NationalHeader.IsMatch(text2))
			{
				return SemanticRole.NationalHeader;
			}
			if (RegexPatterns.Motto.IsMatch(text2))
			{
				return SemanticRole.Motto;
			}
			if (RegexPatterns.PartyHeader.IsMatch(text2))
			{
				return SemanticRole.PartyHeader;
			}
			if (RegexPatterns.YouthUnionHeader.IsMatch(text2) && text2.Length <= 60)
			{
				return SemanticRole.PartyHeader;
			}
			if (RegexPatterns.DocumentNumberGeneral.IsMatch(text2))
			{
				return SemanticRole.DocumentNumber;
			}
			if (RegexPatterns.DocumentDate.IsMatch(text2))
			{
				return SemanticRole.DocumentDate;
			}
			if (RegexPatterns.UrgencyIndicator.IsMatch(text2))
			{
				return SemanticRole.UrgencyIndicator;
			}
			if (RegexPatterns.SecurityIndicator.IsMatch(text2))
			{
				return SemanticRole.SecurityIndicator;
			}
			if (RegexPatterns.CirculationScope.IsMatch(text2))
			{
				return SemanticRole.CirculationScope;
			}
			if (RegexPatterns.DraftIndicator.IsMatch(text2))
			{
				return SemanticRole.DraftIndicator;
			}
			if (RegexPatterns.DocumentType.IsMatch(text2))
			{
				return SemanticRole.DocumentType;
			}
			if (RegexPatterns.DispatchSubject.IsMatch(text2))
			{
				return SemanticRole.DispatchSubject;
			}
			if (RegexPatterns.DirectRecipient.IsMatch(text2))
			{
				return SemanticRole.DirectRecipient;
			}
		}
		if (paragraphIndex <= 25 && RegexPatterns.LegalGround.IsMatch(text2))
		{
			return SemanticRole.LegalGround;
		}
		if (RegexPatterns.PartHeader.IsMatch(text2))
		{
			return SemanticRole.PartHeader;
		}
		if (RegexPatterns.ChapterHeader.IsMatch(text2))
		{
			return SemanticRole.ChapterHeader;
		}
		if (RegexPatterns.SectionHeader.IsMatch(text2))
		{
			return SemanticRole.SectionHeader;
		}
		if (RegexPatterns.ArticleTitle.IsMatch(text2))
		{
			return SemanticRole.ArticleTitle;
		}
		if (RegexPatterns.ClauseNumber.IsMatch(text2))
		{
			return SemanticRole.ClauseText;
		}
		if (RegexPatterns.PointLetter.IsMatch(text2))
		{
			return SemanticRole.PointText;
		}
		if (paragraphIndex >= totalParagraphs - 15)
		{
			if (RegexPatterns.RecipientHeader.IsMatch(text2))
			{
				return SemanticRole.RecipientHeader;
			}
			if (RegexPatterns.RecipientItem.IsMatch(text2))
			{
				return SemanticRole.RecipientItem;
			}
			if (RegexPatterns.SignAuthority.IsMatch(text2))
			{
				return SemanticRole.SignAuthority;
			}
			if (RegexPatterns.SignPosition.IsMatch(text2))
			{
				return SemanticRole.SignTitle;
			}
			if (_activeProfile.RulePack == RulePackType.TT68_2023_BQP && RegexPatterns.MilitaryRank.IsMatch(text2))
			{
				return SemanticRole.SignRankName;
			}
		}
		return SemanticRole.BodyText;
	}

	public ComponentStyle GetComponentStyle(SemanticRole role)
	{
		return _activeProfile.GetComponentStyle(role);
	}
}
