using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Interfaces;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;

namespace VFormat.Core.Serialization;

public class JsonProfileRepository : IProfileRepository
{
	private readonly string _configFilePath;

	private readonly JavaScriptSerializer _serializer;

	public JsonProfileRepository(string customConfigPath = null)
	{
		_configFilePath = ((!string.IsNullOrEmpty(customConfigPath)) ? customConfigPath : Constants.ConfigFilePath);
		_serializer = new JavaScriptSerializer
		{
			MaxJsonLength = 10485760
		};
	}

	public List<ProfileModel> GetAllProfiles()
	{
		List<ProfileModel> list = new List<ProfileModel>();
		list.Add(ND30Rules.CreateProfile());
		list.Add(HD05Rules.CreateProfile());
		list.Add(TT68Rules.CreateProfile());
		list.Add(HD29Rules.CreateProfile());
		return list;
	}

	public ProfileModel GetProfile(string profileId)
	{
		if (string.IsNullOrEmpty(profileId))
		{
			return ND30Rules.CreateProfile();
		}
		string text = profileId.ToUpperInvariant();
		if (text.Contains("DOAN") || text.Contains("HD29") || text.Contains("TWDTN") || text.Contains("YOUTH"))
		{
			return HD29Rules.CreateProfile();
		}
		if (text.Contains("DANG") || text.Contains("HD05") || text.Contains("05HD"))
		{
			return HD05Rules.CreateProfile();
		}
		if (text.Contains("TT68") || text.Contains("BQP") || text.Contains("MILITARY"))
		{
			return TT68Rules.CreateProfile();
		}
		return ND30Rules.CreateProfile();
	}

	public ProfileModel GetDefaultProfile(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => HD05Rules.CreateProfile(), 
			RulePackType.TT68_2023_BQP => TT68Rules.CreateProfile(), 
			RulePackType.HD29_TWDTN => HD29Rules.CreateProfile(), 
			_ => ND30Rules.CreateProfile(), 
		};
	}

	public void SaveProfile(ProfileModel profile)
	{
	}

	public UserPersonalization LoadPersonalization()
	{
		try
		{
			if (File.Exists(_configFilePath))
			{
				try
				{
					string json = File.ReadAllText(_configFilePath, Encoding.UTF8);
					UserPersonalization userPersonalization = ParsePersonalizationJson(json);
					if (userPersonalization != null)
					{
						return userPersonalization;
					}
				}
				catch
				{
				}
			}
			string path = _configFilePath + ".bak";
			if (File.Exists(path))
			{
				try
				{
					string json2 = File.ReadAllText(path, Encoding.UTF8);
					UserPersonalization userPersonalization2 = ParsePersonalizationJson(json2);
					if (userPersonalization2 != null)
					{
						try
						{
							SavePersonalization(userPersonalization2);
						}
						catch
						{
						}
						return userPersonalization2;
					}
				}
				catch
				{
				}
			}
			string directoryName = Path.GetDirectoryName(_configFilePath);
			if (!string.IsNullOrEmpty(directoryName))
			{
				string path2 = Path.Combine(directoryName, "user_presets.json");
				if (File.Exists(path2))
				{
					try
					{
						string json3 = File.ReadAllText(path2, Encoding.UTF8);
						UserPersonalization userPersonalization3 = ParsePersonalizationJson(json3);
						if (userPersonalization3 != null)
						{
							try
							{
								SavePersonalization(userPersonalization3);
							}
							catch
							{
							}
							return userPersonalization3;
						}
					}
					catch
					{
					}
				}
			}
			string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
			string path3 = Path.Combine(baseDirectory, "user_presets.json");
			if (File.Exists(path3))
			{
				try
				{
					string json4 = File.ReadAllText(path3, Encoding.UTF8);
					UserPersonalization userPersonalization4 = ParsePersonalizationJson(json4);
					if (userPersonalization4 != null)
					{
						try
						{
							SavePersonalization(userPersonalization4);
						}
						catch
						{
						}
						return userPersonalization4;
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
		UserPersonalization userPersonalization5 = UserPersonalization.CreateDefault();
		try
		{
			SavePersonalization(userPersonalization5);
		}
		catch
		{
		}
		return userPersonalization5;
	}

	public void SavePersonalization(UserPersonalization personalization)
	{
		if (personalization == null)
		{
			return;
		}
		try
		{
			string directoryName = Path.GetDirectoryName(_configFilePath);
			if (!string.IsNullOrEmpty(directoryName) && !Directory.Exists(directoryName))
			{
				Directory.CreateDirectory(directoryName);
			}
			Dictionary<string, object> dictionary = new Dictionary<string, object>();
			List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
			foreach (AgencyProfile item in personalization.AgenciesAdmin)
			{
				list.Add(new Dictionary<string, object>
				{
					{ "id", item.Id },
					{ "alias", item.Alias },
					{ "parent_agency", item.ParentAgency },
					{ "agency_name", item.AgencyName },
					{ "location", item.Location },
					{ "code_prefix", item.CodePrefix }
				});
			}
			dictionary["agencies_admin"] = list;
			List<Dictionary<string, object>> list2 = new List<Dictionary<string, object>>();
			foreach (AgencyProfile item2 in personalization.AgenciesParty)
			{
				list2.Add(new Dictionary<string, object>
				{
					{ "id", item2.Id },
					{ "alias", item2.Alias },
					{ "parent_party", item2.ParentAgency },
					{ "party_name", item2.AgencyName },
					{ "location", item2.Location },
					{ "code_prefix", item2.CodePrefix }
				});
			}
			dictionary["agencies_party"] = list2;
			List<Dictionary<string, object>> list3 = new List<Dictionary<string, object>>();
			foreach (AgencyProfile item3 in personalization.AgenciesMilitary)
			{
				list3.Add(new Dictionary<string, object>
				{
					{ "id", item3.Id },
					{ "alias", item3.Alias },
					{ "parent_agency", item3.ParentAgency },
					{ "agency_name", item3.AgencyName },
					{ "location", item3.Location },
					{ "code_prefix", item3.CodePrefix }
				});
			}
			dictionary["agencies_military"] = list3;
			List<Dictionary<string, object>> list4 = new List<Dictionary<string, object>>();
			if (personalization.AgenciesCustom != null)
			{
				foreach (AgencyProfile item4 in personalization.AgenciesCustom)
				{
					list4.Add(new Dictionary<string, object>
					{
						{ "id", item4.Id },
						{ "alias", item4.Alias },
						{ "parent_agency", item4.ParentAgency },
						{ "agency_name", item4.AgencyName },
						{ "location", item4.Location },
						{ "code_prefix", item4.CodePrefix }
					});
				}
			}
			dictionary["agencies_custom"] = list4;
			List<Dictionary<string, object>> list5 = new List<Dictionary<string, object>>();
			if (personalization.AgenciesYouth != null)
			{
				foreach (AgencyProfile item5 in personalization.AgenciesYouth)
				{
					list5.Add(new Dictionary<string, object>
					{
						{ "id", item5.Id },
						{ "alias", item5.Alias },
						{ "parent_agency", item5.ParentAgency },
						{ "agency_name", item5.AgencyName },
						{ "location", item5.Location },
						{ "code_prefix", item5.CodePrefix }
					});
				}
			}
			dictionary["agencies_youth"] = list5;
			dictionary["signers_admin"] = SerializeSigners(personalization.SignersAdmin);
			dictionary["signers_party"] = SerializeSigners(personalization.SignersParty);
			dictionary["signers_military"] = SerializeSigners(personalization.SignersMilitary);
			dictionary["signers_custom"] = SerializeSigners(personalization.SignersCustom);
			dictionary["signers_youth"] = SerializeSigners(personalization.SignersYouth);
			dictionary["signers_list"] = SerializeSigners((personalization.SignersCustom.Count > 0) ? personalization.SignersCustom : personalization.Signers);
			List<Dictionary<string, object>> list6 = new List<Dictionary<string, object>>();
			foreach (RecipientProfile recipient in personalization.Recipients)
			{
				list6.Add(new Dictionary<string, object>
				{
					{ "id", recipient.Id },
					{ "title", recipient.Title },
					{ "content", recipient.Content },
					{
						"rule_pack",
						(int)recipient.RulePack
					}
				});
			}
			dictionary["recipients_list"] = list6;
			dictionary["selected_agency_id"] = personalization.SelectedAgencyId;
			dictionary["selected_signer_id"] = personalization.SelectedSignerId;
			dictionary["selected_recipient_id"] = personalization.SelectedRecipientId;
			dictionary["active_rule_pack"] = (int)personalization.ActiveRulePack;
			dictionary["agency_standard"] = personalization.AgencyStandard;
			dictionary["signer_standard"] = personalization.SignerStandard;
			dictionary["margin_standard"] = personalization.MarginStandard;
			dictionary["typography_standard"] = personalization.TypographyStandard;
			dictionary["use_custom_agency"] = personalization.UseCustomAgency;
			dictionary["use_custom_signer"] = personalization.UseCustomSigner;
			dictionary["use_custom_margins"] = personalization.UseCustomMargins;
			dictionary["use_custom_typography"] = personalization.UseCustomTypography;
			dictionary["use_custom_paragraph"] = personalization.UseCustomParagraph;
			dictionary["amount_four_as_tu"] = personalization.AmountFourAsTu;
			dictionary["amount_capitalize_first"] = personalization.AmountCapitalizeFirst;
			dictionary["amount_zero_tens_as_le"] = personalization.AmountZeroTensAsLe;
			dictionary["amount_thousand_as_ngan"] = personalization.AmountThousandAsNgan;
			dictionary["header_date_in_top_cell"] = personalization.HeaderDateInTopCell;
			dictionary["margins_admin"] = SerializeMarginProfile(personalization.MarginsAdmin);
			dictionary["margins_party"] = SerializeMarginProfile(personalization.MarginsParty);
			dictionary["margins_military"] = SerializeMarginProfile(personalization.MarginsMilitary);
			dictionary["margins_custom"] = SerializeMarginProfile(personalization.MarginsCustom);
			dictionary["margins_youth"] = SerializeMarginProfile(personalization.MarginsYouth);
			dictionary["typography_admin"] = SerializeTypographyProfile(personalization.TypographyAdmin);
			dictionary["typography_party"] = SerializeTypographyProfile(personalization.TypographyParty);
			dictionary["typography_military"] = SerializeTypographyProfile(personalization.TypographyMilitary);
			dictionary["typography_custom"] = SerializeTypographyProfile(personalization.TypographyCustom);
			dictionary["typography_youth"] = SerializeTypographyProfile(personalization.TypographyYouth);
			dictionary["margin_top_mm"] = personalization.MarginTopMm;
			dictionary["margin_bottom_mm"] = personalization.MarginBottomMm;
			dictionary["margin_left_mm"] = personalization.MarginLeftMm;
			dictionary["margin_right_mm"] = personalization.MarginRightMm;
			dictionary["first_line_indent_cm"] = personalization.FirstLineIndentCm;
			dictionary["spacing_before_pt"] = personalization.SpacingBeforePt;
			dictionary["spacing_after_pt"] = personalization.SpacingAfterPt;
			dictionary["line_spacing_rule"] = personalization.LineSpacingRule ?? "1.15 lines";
			dictionary["line_spacing_at_pt"] = personalization.LineSpacingAtPt;
			dictionary["no_space_between_paragraphs_of_same_style"] = personalization.NoSpaceBetweenParagraphsOfSameStyle;
			dictionary["size_quoc_hieu"] = personalization.SizeQuocHieu;
			dictionary["size_tieu_ngu"] = personalization.SizeTieuNgu;
			dictionary["size_co_quan"] = personalization.SizeCoQuan;
			dictionary["size_so_ky_hieu"] = personalization.SizeSoKyHieu;
			dictionary["size_dia_danh"] = personalization.SizeDiaDanh;
			dictionary["size_ten_loai"] = personalization.SizeTenLoai;
			dictionary["size_trich_yeu"] = personalization.SizeTrichYeu;
			dictionary["size_than_bai"] = personalization.SizeThanBai;
			dictionary["size_noi_nhan"] = personalization.SizeNoiNhan;
			dictionary["size_chu_ky"] = personalization.SizeChuKy;
			string text = _serializer.Serialize(dictionary);
			string text2 = _configFilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
			try
			{
				using (FileStream fileStream = new FileStream(text2, FileMode.CreateNew, FileAccess.Write, FileShare.None))
				{
					byte[] bytes = Encoding.UTF8.GetBytes(text);
					fileStream.Write(bytes, 0, bytes.Length);
					fileStream.Flush(flushToDisk: true);
				}
				if (File.Exists(_configFilePath))
				{
					bool flag = false;
					try
					{
						flag = _serializer.DeserializeObject(File.ReadAllText(_configFilePath)) is Dictionary<string, object>;
					}
					catch
					{
					}
					File.Replace(text2, _configFilePath, flag ? (_configFilePath + ".bak") : null);
				}
				else
				{
					File.Move(text2, _configFilePath);
				}
			}
			finally
			{
				if (File.Exists(text2))
				{
					File.Delete(text2);
				}
			}
			string path = Path.Combine(directoryName ?? "", "user_presets.json");
			if (!Path.GetFullPath(path).Equals(Path.GetFullPath(_configFilePath), StringComparison.OrdinalIgnoreCase))
			{
				try
				{
					File.WriteAllText(path, text, Encoding.UTF8);
					return;
				}
				catch
				{
					return;
				}
			}
		}
		catch
		{
			throw;
		}
	}

	public void ResetToDefaults()
	{
		UserPersonalization personalization = UserPersonalization.CreateDefault();
		SavePersonalization(personalization);
	}

	private UserPersonalization ParsePersonalizationJson(string json)
	{
		if (string.IsNullOrEmpty(json))
		{
			return null;
		}
		UserPersonalization userPersonalization = UserPersonalization.CreateDefault();
		Dictionary<string, object> dictionary = _serializer.Deserialize<Dictionary<string, object>>(json);
		if (dictionary == null)
		{
			return userPersonalization;
		}
		if (dictionary.ContainsKey("agencies_admin") && dictionary["agencies_admin"] is ArrayList { Count: >0 } arrayList)
		{
			userPersonalization.AgenciesAdmin.Clear();
			foreach (Dictionary<string, object> item in arrayList)
			{
				userPersonalization.AgenciesAdmin.Add(new AgencyProfile(GetString(item, "id"), GetString(item, "alias"), GetString(item, "parent_agency"), GetString(item, "agency_name"), GetString(item, "location"), GetString(item, "code_prefix")));
			}
		}
		if (dictionary.ContainsKey("agencies_party") && dictionary["agencies_party"] is ArrayList { Count: >0 } arrayList2)
		{
			userPersonalization.AgenciesParty.Clear();
			foreach (Dictionary<string, object> item2 in arrayList2)
			{
				string parentAgency = (item2.ContainsKey("parent_party") ? GetString(item2, "parent_party") : GetString(item2, "parent_agency"));
				string agencyName = (item2.ContainsKey("party_name") ? GetString(item2, "party_name") : GetString(item2, "agency_name"));
				userPersonalization.AgenciesParty.Add(new AgencyProfile(GetString(item2, "id"), GetString(item2, "alias"), parentAgency, agencyName, GetString(item2, "location"), GetString(item2, "code_prefix"), RulePackType.HD05_VPTW));
			}
		}
		if (dictionary.ContainsKey("agencies_military") && dictionary["agencies_military"] is ArrayList { Count: >0 } arrayList3)
		{
			userPersonalization.AgenciesMilitary.Clear();
			foreach (Dictionary<string, object> item3 in arrayList3)
			{
				userPersonalization.AgenciesMilitary.Add(new AgencyProfile(GetString(item3, "id"), GetString(item3, "alias"), GetString(item3, "parent_agency"), GetString(item3, "agency_name"), GetString(item3, "location"), GetString(item3, "code_prefix"), RulePackType.TT68_2023_BQP));
			}
		}
		if (dictionary.ContainsKey("agencies_custom") && dictionary["agencies_custom"] is ArrayList { Count: >0 } arrayList4)
		{
			userPersonalization.AgenciesCustom.Clear();
			foreach (Dictionary<string, object> item4 in arrayList4)
			{
				userPersonalization.AgenciesCustom.Add(new AgencyProfile(GetString(item4, "id"), GetString(item4, "alias"), GetString(item4, "parent_agency"), GetString(item4, "agency_name"), GetString(item4, "location"), GetString(item4, "code_prefix")));
			}
		}
		if (dictionary.ContainsKey("agencies_youth") && dictionary["agencies_youth"] is ArrayList { Count: >0 } arrayList5)
		{
			userPersonalization.AgenciesYouth.Clear();
			foreach (Dictionary<string, object> item5 in arrayList5)
			{
				userPersonalization.AgenciesYouth.Add(new AgencyProfile(GetString(item5, "id"), GetString(item5, "alias"), GetString(item5, "parent_agency"), GetString(item5, "agency_name"), GetString(item5, "location"), GetString(item5, "code_prefix"), RulePackType.HD29_TWDTN));
			}
		}
		if (dictionary.ContainsKey("signers_admin"))
		{
			List<SignerProfile> list = DeserializeSigners(dictionary["signers_admin"] as ArrayList);
			if (list != null && list.Count > 0)
			{
				userPersonalization.SignersAdmin = list;
			}
		}
		if (dictionary.ContainsKey("signers_party"))
		{
			List<SignerProfile> list2 = DeserializeSigners(dictionary["signers_party"] as ArrayList);
			if (list2 != null && list2.Count > 0)
			{
				userPersonalization.SignersParty = list2;
			}
		}
		if (dictionary.ContainsKey("signers_military"))
		{
			List<SignerProfile> list3 = DeserializeSigners(dictionary["signers_military"] as ArrayList);
			if (list3 != null && list3.Count > 0)
			{
				userPersonalization.SignersMilitary = list3;
			}
		}
		if (dictionary.ContainsKey("signers_youth"))
		{
			List<SignerProfile> list4 = DeserializeSigners(dictionary["signers_youth"] as ArrayList);
			if (list4 != null && list4.Count > 0)
			{
				foreach (SignerProfile item6 in list4)
				{
					item6.RulePack = RulePackType.HD29_TWDTN;
				}
				userPersonalization.SignersYouth = list4;
			}
		}
		if (dictionary.ContainsKey("signers_custom"))
		{
			List<SignerProfile> list5 = DeserializeSigners(dictionary["signers_custom"] as ArrayList);
			if (list5 != null && list5.Count > 0)
			{
				userPersonalization.SignersCustom = list5;
			}
		}
		if (dictionary.ContainsKey("signers_list"))
		{
			List<SignerProfile> list6 = DeserializeSigners(dictionary["signers_list"] as ArrayList);
			if (list6 != null && list6.Count > 0)
			{
				userPersonalization.Signers = list6;
				if (!dictionary.ContainsKey("signers_custom"))
				{
					userPersonalization.SignersCustom.Clear();
					foreach (SignerProfile item7 in list6)
					{
						userPersonalization.SignersCustom.Add(item7.Clone());
					}
				}
			}
		}
		if (dictionary.ContainsKey("recipients_list") && dictionary["recipients_list"] is ArrayList arrayList6)
		{
			foreach (Dictionary<string, object> item8 in arrayList6)
			{
				int rulePack = ((!item8.ContainsKey("rule_pack")) ? 1 : Convert.ToInt32(item8["rule_pack"]));
				userPersonalization.Recipients.Add(new RecipientProfile(GetString(item8, "id"), GetString(item8, "title"), GetString(item8, "content"), (RulePackType)rulePack));
			}
		}
		if (dictionary.ContainsKey("selected_agency_id"))
		{
			userPersonalization.SelectedAgencyId = GetString(dictionary, "selected_agency_id");
		}
		if (dictionary.ContainsKey("selected_signer_id"))
		{
			userPersonalization.SelectedSignerId = GetString(dictionary, "selected_signer_id");
		}
		if (dictionary.ContainsKey("selected_recipient_id"))
		{
			userPersonalization.SelectedRecipientId = GetString(dictionary, "selected_recipient_id");
		}
		if (dictionary.ContainsKey("active_rule_pack"))
		{
			userPersonalization.ActiveRulePack = (RulePackType)Convert.ToInt32(dictionary["active_rule_pack"]);
		}
		if (dictionary.ContainsKey("agency_standard"))
		{
			userPersonalization.AgencyStandard = Convert.ToInt32(dictionary["agency_standard"]);
		}
		if (dictionary.ContainsKey("signer_standard"))
		{
			userPersonalization.SignerStandard = Convert.ToInt32(dictionary["signer_standard"]);
		}
		if (dictionary.ContainsKey("margin_standard"))
		{
			userPersonalization.MarginStandard = Convert.ToInt32(dictionary["margin_standard"]);
		}
		if (dictionary.ContainsKey("typography_standard"))
		{
			userPersonalization.TypographyStandard = Convert.ToInt32(dictionary["typography_standard"]);
		}
		userPersonalization.UseCustomAgency = GetBool(dictionary, "use_custom_agency", defaultValue: false);
		userPersonalization.UseCustomSigner = GetBool(dictionary, "use_custom_signer", defaultValue: false);
		userPersonalization.UseCustomMargins = GetBool(dictionary, "use_custom_margins", defaultValue: false);
		userPersonalization.UseCustomTypography = GetBool(dictionary, "use_custom_typography", defaultValue: false);
		userPersonalization.UseCustomParagraph = GetBool(dictionary, "use_custom_paragraph", defaultValue: false);
		userPersonalization.AmountFourAsTu = GetBool(dictionary, "amount_four_as_tu", defaultValue: false);
		userPersonalization.AmountCapitalizeFirst = GetBool(dictionary, "amount_capitalize_first", defaultValue: true);
		userPersonalization.AmountZeroTensAsLe = GetBool(dictionary, "amount_zero_tens_as_le", defaultValue: false);
		userPersonalization.AmountThousandAsNgan = GetBool(dictionary, "amount_thousand_as_ngan", defaultValue: false);
		userPersonalization.HeaderDateInTopCell = GetBool(dictionary, "header_date_in_top_cell", defaultValue: false);
		if (dictionary.ContainsKey("margins_admin"))
		{
			userPersonalization.MarginsAdmin = DeserializeMarginProfile(dictionary["margins_admin"], userPersonalization.MarginsAdmin);
		}
		if (dictionary.ContainsKey("margins_party"))
		{
			userPersonalization.MarginsParty = DeserializeMarginProfile(dictionary["margins_party"], userPersonalization.MarginsParty);
		}
		if (dictionary.ContainsKey("margins_military"))
		{
			userPersonalization.MarginsMilitary = DeserializeMarginProfile(dictionary["margins_military"], userPersonalization.MarginsMilitary);
		}
		if (dictionary.ContainsKey("margins_custom"))
		{
			userPersonalization.MarginsCustom = DeserializeMarginProfile(dictionary["margins_custom"], userPersonalization.MarginsCustom);
		}
		if (dictionary.ContainsKey("margins_youth"))
		{
			userPersonalization.MarginsYouth = DeserializeMarginProfile(dictionary["margins_youth"], userPersonalization.MarginsYouth);
		}
		if (dictionary.ContainsKey("typography_admin"))
		{
			userPersonalization.TypographyAdmin = DeserializeTypographyProfile(dictionary["typography_admin"], userPersonalization.TypographyAdmin);
		}
		if (dictionary.ContainsKey("typography_party"))
		{
			userPersonalization.TypographyParty = DeserializeTypographyProfile(dictionary["typography_party"], userPersonalization.TypographyParty);
		}
		if (dictionary.ContainsKey("typography_military"))
		{
			userPersonalization.TypographyMilitary = DeserializeTypographyProfile(dictionary["typography_military"], userPersonalization.TypographyMilitary);
		}
		if (dictionary.ContainsKey("typography_custom"))
		{
			userPersonalization.TypographyCustom = DeserializeTypographyProfile(dictionary["typography_custom"], userPersonalization.TypographyCustom);
		}
		if (dictionary.ContainsKey("typography_youth"))
		{
			userPersonalization.TypographyYouth = DeserializeTypographyProfile(dictionary["typography_youth"], userPersonalization.TypographyYouth);
		}
		if (userPersonalization.TypographyAdmin != null)
		{
			userPersonalization.TypographyAdmin.ValidateAndClamp(RulePackType.ND30_2020);
		}
		if (userPersonalization.TypographyParty != null)
		{
			userPersonalization.TypographyParty.ValidateAndClamp(RulePackType.HD05_VPTW);
		}
		if (userPersonalization.TypographyMilitary != null)
		{
			userPersonalization.TypographyMilitary.ValidateAndClamp(RulePackType.TT68_2023_BQP);
		}
		if (userPersonalization.TypographyCustom != null)
		{
			userPersonalization.TypographyCustom.ValidateAndClamp(RulePackType.ND30_2020);
		}
		if (userPersonalization.TypographyYouth != null)
		{
			userPersonalization.TypographyYouth.ValidateAndClamp(RulePackType.HD29_TWDTN);
		}
		userPersonalization.MarginTopMm = GetFloat(dictionary, "margin_top_mm", userPersonalization.MarginTopMm);
		userPersonalization.MarginBottomMm = GetFloat(dictionary, "margin_bottom_mm", userPersonalization.MarginBottomMm);
		userPersonalization.MarginLeftMm = GetFloat(dictionary, "margin_left_mm", userPersonalization.MarginLeftMm);
		userPersonalization.MarginRightMm = GetFloat(dictionary, "margin_right_mm", userPersonalization.MarginRightMm);
		userPersonalization.FirstLineIndentCm = GetFloat(dictionary, "first_line_indent_cm", userPersonalization.FirstLineIndentCm);
		userPersonalization.SpacingBeforePt = GetFloat(dictionary, "spacing_before_pt", userPersonalization.SpacingBeforePt);
		userPersonalization.SpacingAfterPt = GetFloat(dictionary, "spacing_after_pt", userPersonalization.SpacingAfterPt);
		if (dictionary.ContainsKey("line_spacing_rule"))
		{
			userPersonalization.LineSpacingRule = GetString(dictionary, "line_spacing_rule");
		}
		userPersonalization.LineSpacingAtPt = GetFloat(dictionary, "line_spacing_at_pt", userPersonalization.LineSpacingAtPt);
		userPersonalization.NoSpaceBetweenParagraphsOfSameStyle = GetBool(dictionary, "no_space_between_paragraphs_of_same_style", userPersonalization.NoSpaceBetweenParagraphsOfSameStyle);
		userPersonalization.SizeQuocHieu = (float)Math.Round(GetFloat(dictionary, "size_quoc_hieu", userPersonalization.SizeQuocHieu));
		userPersonalization.SizeTieuNgu = (float)Math.Round(GetFloat(dictionary, "size_tieu_ngu", userPersonalization.SizeTieuNgu));
		userPersonalization.SizeCoQuan = (float)Math.Round(GetFloat(dictionary, "size_co_quan", userPersonalization.SizeCoQuan));
		userPersonalization.SizeSoKyHieu = (float)Math.Round(GetFloat(dictionary, "size_so_ky_hieu", userPersonalization.SizeSoKyHieu));
		userPersonalization.SizeDiaDanh = (float)Math.Round(GetFloat(dictionary, "size_dia_danh", userPersonalization.SizeDiaDanh));
		userPersonalization.SizeTenLoai = (float)Math.Round(GetFloat(dictionary, "size_ten_loai", userPersonalization.SizeTenLoai));
		userPersonalization.SizeTrichYeu = (float)Math.Round(GetFloat(dictionary, "size_trich_yeu", userPersonalization.SizeTrichYeu));
		userPersonalization.SizeThanBai = (float)Math.Round(GetFloat(dictionary, "size_than_bai", userPersonalization.SizeThanBai));
		userPersonalization.SizeNoiNhan = (float)Math.Round(GetFloat(dictionary, "size_noi_nhan", userPersonalization.SizeNoiNhan));
		userPersonalization.SizeChuKy = (float)Math.Round(GetFloat(dictionary, "size_chu_ky", userPersonalization.SizeChuKy));
		if (userPersonalization.TypographyCustom != null)
		{
			userPersonalization.SizeQuocHieu = userPersonalization.TypographyCustom.SizeQuocHieu;
			userPersonalization.SizeTieuNgu = userPersonalization.TypographyCustom.SizeTieuNgu;
			userPersonalization.SizeCoQuan = userPersonalization.TypographyCustom.SizeCoQuan;
			userPersonalization.SizeSoKyHieu = userPersonalization.TypographyCustom.SizeSoKyHieu;
			userPersonalization.SizeDiaDanh = userPersonalization.TypographyCustom.SizeDiaDanh;
			userPersonalization.SizeTenLoai = userPersonalization.TypographyCustom.SizeTenLoai;
			userPersonalization.SizeTrichYeu = userPersonalization.TypographyCustom.SizeTrichYeu;
			userPersonalization.SizeThanBai = userPersonalization.TypographyCustom.SizeThanBai;
			userPersonalization.SizeNoiNhan = userPersonalization.TypographyCustom.SizeNoiNhan;
			userPersonalization.SizeChuKy = userPersonalization.TypographyCustom.SizeChuKy;
		}
		return userPersonalization;
	}

	private static string GetString(Dictionary<string, object> dict, string key)
	{
		if (dict != null && dict.ContainsKey(key) && dict[key] != null)
		{
			return dict[key].ToString();
		}
		return string.Empty;
	}

	public static bool TryParseFlexibleFloat(string input, out float result)
	{
		result = 0f;
		if (string.IsNullOrWhiteSpace(input))
		{
			return false;
		}
		string text = input.Trim();
		int num = text.LastIndexOf('.');
		int num2 = text.LastIndexOf(',');
		if (num < 0 || num2 < 0)
		{
			text = text.Replace(',', '.');
		}
		else
		{
			text = ((num2 <= num) ? text.Replace(",", "") : text.Replace(".", "").Replace(',', '.'));
		}
		return float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
	}

	private static float GetFloat(Dictionary<string, object> dict, string key, float defaultValue)
	{
		if (dict != null && dict.ContainsKey(key) && dict[key] != null)
		{
			object obj = dict[key];
			float num = defaultValue;
			bool flag = false;
			if (obj is float)
			{
				num = (float)obj;
				flag = true;
			}
			else if (obj is double)
			{
				num = (float)(double)obj;
				flag = true;
			}
			else if (obj is int)
			{
				num = (int)obj;
				flag = true;
			}
			else if (obj is long)
			{
				num = (long)obj;
				flag = true;
			}
			else if (obj is decimal)
			{
				num = (float)(decimal)obj;
				flag = true;
			}
			else
			{
				string input = obj.ToString();
				if (TryParseFlexibleFloat(input, out var result))
				{
					num = result;
					flag = true;
				}
			}
			if (flag)
			{
				if (key == "first_line_indent_cm" && num >= 50f)
				{
					num /= 100f;
				}
				return num;
			}
		}
		return defaultValue;
	}

	private static bool GetBool(Dictionary<string, object> dict, string key, bool defaultValue)
	{
		if (dict != null && dict.ContainsKey(key) && dict[key] != null)
		{
			try
			{
				return Convert.ToBoolean(dict[key]);
			}
			catch
			{
			}
		}
		return defaultValue;
	}

	private static Dictionary<string, object> SerializeMarginProfile(MarginProfile m)
	{
		if (m == null)
		{
			return null;
		}
		Dictionary<string, object> dictionary = new Dictionary<string, object>();
		dictionary.Add("margin_top_mm", m.MarginTopMm);
		dictionary.Add("margin_bottom_mm", m.MarginBottomMm);
		dictionary.Add("margin_left_mm", m.MarginLeftMm);
		dictionary.Add("margin_right_mm", m.MarginRightMm);
		dictionary.Add("first_line_indent_cm", m.FirstLineIndentCm);
		dictionary.Add("spacing_before_pt", m.SpacingBeforePt);
		dictionary.Add("spacing_after_pt", m.SpacingAfterPt);
		dictionary.Add("line_spacing_rule", m.LineSpacingRule ?? "1.15 lines");
		dictionary.Add("line_spacing_at_pt", m.LineSpacingAtPt);
		dictionary.Add("no_space_between_paragraphs_of_same_style", m.NoSpaceBetweenParagraphsOfSameStyle);
		return dictionary;
	}

	private static MarginProfile DeserializeMarginProfile(object obj, MarginProfile defaultProfile)
	{
		if (!(obj is Dictionary<string, object> dictionary))
		{
			return defaultProfile;
		}
		MarginProfile marginProfile = ((defaultProfile != null) ? defaultProfile.Clone() : new MarginProfile());
		marginProfile.MarginTopMm = GetFloat(dictionary, "margin_top_mm", marginProfile.MarginTopMm);
		marginProfile.MarginBottomMm = GetFloat(dictionary, "margin_bottom_mm", marginProfile.MarginBottomMm);
		marginProfile.MarginLeftMm = GetFloat(dictionary, "margin_left_mm", marginProfile.MarginLeftMm);
		marginProfile.MarginRightMm = GetFloat(dictionary, "margin_right_mm", marginProfile.MarginRightMm);
		marginProfile.FirstLineIndentCm = GetFloat(dictionary, "first_line_indent_cm", marginProfile.FirstLineIndentCm);
		marginProfile.SpacingBeforePt = GetFloat(dictionary, "spacing_before_pt", marginProfile.SpacingBeforePt);
		marginProfile.SpacingAfterPt = GetFloat(dictionary, "spacing_after_pt", marginProfile.SpacingAfterPt);
		if (dictionary.ContainsKey("line_spacing_rule"))
		{
			marginProfile.LineSpacingRule = GetString(dictionary, "line_spacing_rule");
		}
		marginProfile.LineSpacingAtPt = GetFloat(dictionary, "line_spacing_at_pt", marginProfile.LineSpacingAtPt);
		marginProfile.NoSpaceBetweenParagraphsOfSameStyle = GetBool(dictionary, "no_space_between_paragraphs_of_same_style", marginProfile.NoSpaceBetweenParagraphsOfSameStyle);
		if (!dictionary.ContainsKey("line_spacing_at_pt") && string.Equals(marginProfile.LineSpacingRule, "Exactly 20 pt", StringComparison.OrdinalIgnoreCase))
		{
			marginProfile.LineSpacingRule = "Exactly";
			marginProfile.LineSpacingAtPt = 18f;
			marginProfile.SpacingBeforePt = 6f;
			marginProfile.SpacingAfterPt = 6f;
		}
		return marginProfile;
	}

	private static Dictionary<string, object> SerializeTypographyProfile(TypographyProfile t)
	{
		if (t == null)
		{
			return null;
		}
		Dictionary<string, object> dictionary = new Dictionary<string, object>();
		dictionary.Add("size_quoc_hieu", t.SizeQuocHieu);
		dictionary.Add("size_tieu_ngu", t.SizeTieuNgu);
		dictionary.Add("size_co_quan", t.SizeCoQuan);
		dictionary.Add("size_so_ky_hieu", t.SizeSoKyHieu);
		dictionary.Add("size_dia_danh", t.SizeDiaDanh);
		dictionary.Add("size_ten_loai", t.SizeTenLoai);
		dictionary.Add("size_trich_yeu", t.SizeTrichYeu);
		dictionary.Add("size_trich_yeu_cong_van", t.SizeTrichYeuCongVan);
		dictionary.Add("size_than_bai", t.SizeThanBai);
		dictionary.Add("size_noi_nhan", t.SizeNoiNhan);
		dictionary.Add("size_chu_ky", t.SizeChuKy);
		dictionary.Add("bold_quoc_hieu", t.BoldQuocHieu);
		dictionary.Add("italic_quoc_hieu", t.ItalicQuocHieu);
		dictionary.Add("bold_tieu_ngu", t.BoldTieuNgu);
		dictionary.Add("italic_tieu_ngu", t.ItalicTieuNgu);
		dictionary.Add("bold_co_quan", t.BoldCoQuan);
		dictionary.Add("italic_co_quan", t.ItalicCoQuan);
		dictionary.Add("bold_so_ky_hieu", t.BoldSoKyHieu);
		dictionary.Add("italic_so_ky_hieu", t.ItalicSoKyHieu);
		dictionary.Add("bold_dia_danh", t.BoldDiaDanh);
		dictionary.Add("italic_dia_danh", t.ItalicDiaDanh);
		dictionary.Add("bold_ten_loai", t.BoldTenLoai);
		dictionary.Add("italic_ten_loai", t.ItalicTenLoai);
		dictionary.Add("bold_trich_yeu", t.BoldTrichYeu);
		dictionary.Add("italic_trich_yeu", t.ItalicTrichYeu);
		dictionary.Add("bold_trich_yeu_cong_van", t.BoldTrichYeuCongVan);
		dictionary.Add("italic_trich_yeu_cong_van", t.ItalicTrichYeuCongVan);
		dictionary.Add("bold_than_bai", t.BoldThanBai);
		dictionary.Add("italic_than_bai", t.ItalicThanBai);
		dictionary.Add("bold_noi_nhan", t.BoldNoiNhan);
		dictionary.Add("italic_noi_nhan", t.ItalicNoiNhan);
		dictionary.Add("bold_chu_ky", t.BoldChuKy);
		dictionary.Add("italic_chu_ky", t.ItalicChuKy);
		return dictionary;
	}

	private static TypographyProfile DeserializeTypographyProfile(object obj, TypographyProfile defaultProfile)
	{
		if (!(obj is Dictionary<string, object> dict))
		{
			return defaultProfile;
		}
		TypographyProfile typographyProfile = ((defaultProfile != null) ? defaultProfile.Clone() : new TypographyProfile());
		typographyProfile.SizeQuocHieu = GetFloat(dict, "size_quoc_hieu", typographyProfile.SizeQuocHieu);
		typographyProfile.SizeTieuNgu = GetFloat(dict, "size_tieu_ngu", typographyProfile.SizeTieuNgu);
		typographyProfile.SizeCoQuan = GetFloat(dict, "size_co_quan", typographyProfile.SizeCoQuan);
		typographyProfile.SizeSoKyHieu = GetFloat(dict, "size_so_ky_hieu", typographyProfile.SizeSoKyHieu);
		typographyProfile.SizeDiaDanh = GetFloat(dict, "size_dia_danh", typographyProfile.SizeDiaDanh);
		typographyProfile.SizeTenLoai = GetFloat(dict, "size_ten_loai", typographyProfile.SizeTenLoai);
		typographyProfile.SizeTrichYeu = GetFloat(dict, "size_trich_yeu", typographyProfile.SizeTrichYeu);
		typographyProfile.SizeTrichYeuCongVan = GetFloat(dict, "size_trich_yeu_cong_van", typographyProfile.SizeTrichYeuCongVan);
		typographyProfile.SizeThanBai = GetFloat(dict, "size_than_bai", typographyProfile.SizeThanBai);
		typographyProfile.SizeNoiNhan = GetFloat(dict, "size_noi_nhan", typographyProfile.SizeNoiNhan);
		typographyProfile.SizeChuKy = GetFloat(dict, "size_chu_ky", typographyProfile.SizeChuKy);
		typographyProfile.BoldQuocHieu = GetBool(dict, "bold_quoc_hieu", typographyProfile.BoldQuocHieu);
		typographyProfile.ItalicQuocHieu = GetBool(dict, "italic_quoc_hieu", typographyProfile.ItalicQuocHieu);
		typographyProfile.BoldTieuNgu = GetBool(dict, "bold_tieu_ngu", typographyProfile.BoldTieuNgu);
		typographyProfile.ItalicTieuNgu = GetBool(dict, "italic_tieu_ngu", typographyProfile.ItalicTieuNgu);
		typographyProfile.BoldCoQuan = GetBool(dict, "bold_co_quan", typographyProfile.BoldCoQuan);
		typographyProfile.ItalicCoQuan = GetBool(dict, "italic_co_quan", typographyProfile.ItalicCoQuan);
		typographyProfile.BoldSoKyHieu = GetBool(dict, "bold_so_ky_hieu", typographyProfile.BoldSoKyHieu);
		typographyProfile.ItalicSoKyHieu = GetBool(dict, "italic_so_ky_hieu", typographyProfile.ItalicSoKyHieu);
		typographyProfile.BoldDiaDanh = GetBool(dict, "bold_dia_danh", typographyProfile.BoldDiaDanh);
		typographyProfile.ItalicDiaDanh = GetBool(dict, "italic_dia_danh", typographyProfile.ItalicDiaDanh);
		typographyProfile.BoldTenLoai = GetBool(dict, "bold_ten_loai", typographyProfile.BoldTenLoai);
		typographyProfile.ItalicTenLoai = GetBool(dict, "italic_ten_loai", typographyProfile.ItalicTenLoai);
		typographyProfile.BoldTrichYeu = GetBool(dict, "bold_trich_yeu", typographyProfile.BoldTrichYeu);
		typographyProfile.ItalicTrichYeu = GetBool(dict, "italic_trich_yeu", typographyProfile.ItalicTrichYeu);
		typographyProfile.BoldTrichYeuCongVan = GetBool(dict, "bold_trich_yeu_cong_van", typographyProfile.BoldTrichYeuCongVan);
		typographyProfile.ItalicTrichYeuCongVan = GetBool(dict, "italic_trich_yeu_cong_van", typographyProfile.ItalicTrichYeuCongVan);
		typographyProfile.BoldThanBai = GetBool(dict, "bold_than_bai", typographyProfile.BoldThanBai);
		typographyProfile.ItalicThanBai = GetBool(dict, "italic_than_bai", typographyProfile.ItalicThanBai);
		typographyProfile.BoldNoiNhan = GetBool(dict, "bold_noi_nhan", typographyProfile.BoldNoiNhan);
		typographyProfile.ItalicNoiNhan = GetBool(dict, "italic_noi_nhan", typographyProfile.ItalicNoiNhan);
		typographyProfile.BoldChuKy = GetBool(dict, "bold_chu_ky", typographyProfile.BoldChuKy);
		typographyProfile.ItalicChuKy = GetBool(dict, "italic_chu_ky", typographyProfile.ItalicChuKy);
		return typographyProfile;
	}

	private static List<Dictionary<string, object>> SerializeSigners(List<SignerProfile> signers)
	{
		List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
		if (signers == null)
		{
			return list;
		}
		foreach (SignerProfile signer in signers)
		{
			list.Add(new Dictionary<string, object>
			{
				{ "id", signer.Id },
				{ "alias", signer.Alias },
				{ "title", signer.AuthorityTitle },
				{ "position", signer.Position },
				{ "fullname", signer.FullName },
				{
					"military_rank",
					signer.MilitaryRank ?? ""
				},
				{ "signer_type", signer.SignerType },
				{
					"rule_pack",
					(int)signer.RulePack
				},
				{
					"recipient_list",
					signer.RecipientList ?? ""
				},
				{
					"storage",
					signer.Storage ?? ""
				}
			});
		}
		return list;
	}

	private static List<SignerProfile> DeserializeSigners(ArrayList arr)
	{
		List<SignerProfile> list = new List<SignerProfile>();
		if (arr == null)
		{
			return list;
		}
		foreach (Dictionary<string, object> item in arr)
		{
			int rulePack = ((!item.ContainsKey("rule_pack")) ? 1 : Convert.ToInt32(item["rule_pack"]));
			string militaryRank = (item.ContainsKey("military_rank") ? GetString(item, "military_rank") : "");
			string recipientList = (item.ContainsKey("recipient_list") ? GetString(item, "recipient_list") : "");
			string storage = (item.ContainsKey("storage") ? GetString(item, "storage") : "");
			list.Add(new SignerProfile(GetString(item, "id"), GetString(item, "alias"), GetString(item, "title"), GetString(item, "position"), GetString(item, "fullname"), GetString(item, "signer_type"), militaryRank, (RulePackType)rulePack, recipientList, storage));
		}
		return list;
	}
}
