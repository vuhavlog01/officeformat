using System;
using System.Collections.Generic;
using System.Linq;
using VFormat.Core.Common;
using VFormat.Core.Enums;

namespace VFormat.Core.Personalization;

public class UserPersonalization
{
	public List<AgencyProfile> AgenciesAdmin { get; set; }

	public List<AgencyProfile> AgenciesParty { get; set; }

	public List<AgencyProfile> AgenciesMilitary { get; set; }

	public List<AgencyProfile> AgenciesCustom { get; set; }

	public List<AgencyProfile> AgenciesYouth { get; set; }

	public List<SignerProfile> SignersAdmin { get; set; }

	public List<SignerProfile> SignersParty { get; set; }

	public List<SignerProfile> SignersMilitary { get; set; }

	public List<SignerProfile> SignersCustom { get; set; }

	public List<SignerProfile> SignersYouth { get; set; }

	public List<SignerProfile> Signers { get; set; }

	public List<RecipientProfile> Recipients { get; set; }

	public string SelectedAgencyId { get; set; }

	public string SelectedSignerId { get; set; }

	public string SelectedRecipientId { get; set; }

	public RulePackType ActiveRulePack { get; set; }

	public int AgencyStandard { get; set; }

	public int SignerStandard { get; set; }

	public int MarginStandard { get; set; }

	public int TypographyStandard { get; set; }

	public bool UseCustomAgency { get; set; }

	public bool UseCustomSigner { get; set; }

	public bool UseCustomMargins { get; set; }

	public bool UseCustomTypography { get; set; }

	public bool UseCustomParagraph { get; set; }

	public bool AmountFourAsTu { get; set; }

	public bool AmountCapitalizeFirst { get; set; }

	public bool AmountZeroTensAsLe { get; set; }

	public bool AmountThousandAsNgan { get; set; }

	public bool HeaderDateInTopCell { get; set; }

	public MarginProfile MarginsAdmin { get; set; }

	public MarginProfile MarginsParty { get; set; }

	public MarginProfile MarginsMilitary { get; set; }

	public MarginProfile MarginsCustom { get; set; }

	public MarginProfile MarginsYouth { get; set; }

	public TypographyProfile TypographyAdmin { get; set; }

	public TypographyProfile TypographyParty { get; set; }

	public TypographyProfile TypographyMilitary { get; set; }

	public TypographyProfile TypographyCustom { get; set; }

	public TypographyProfile TypographyYouth { get; set; }

	public float MarginTopMm
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 20f;
			}
			return MarginsCustom.MarginTopMm;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.MarginTopMm = value;
			}
		}
	}

	public float MarginBottomMm
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 20f;
			}
			return MarginsCustom.MarginBottomMm;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.MarginBottomMm = value;
			}
		}
	}

	public float MarginLeftMm
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 30f;
			}
			return MarginsCustom.MarginLeftMm;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.MarginLeftMm = value;
			}
		}
	}

	public float MarginRightMm
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 15f;
			}
			return MarginsCustom.MarginRightMm;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.MarginRightMm = value;
			}
		}
	}

	public float FirstLineIndentCm
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 1f;
			}
			return MarginsCustom.FirstLineIndentCm;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.FirstLineIndentCm = value;
			}
		}
	}

	public float SpacingAfterPt
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 6f;
			}
			return MarginsCustom.SpacingAfterPt;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.SpacingAfterPt = value;
			}
		}
	}

	public string LineSpacingRule
	{
		get
		{
			if (MarginsCustom == null)
			{
				return "1.15 lines";
			}
			return MarginsCustom.LineSpacingRule;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.LineSpacingRule = value;
			}
		}
	}

	public float SpacingBeforePt
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 0f;
			}
			return MarginsCustom.SpacingBeforePt;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.SpacingBeforePt = value;
			}
		}
	}

	public float LineSpacingAtPt
	{
		get
		{
			if (MarginsCustom == null)
			{
				return 0f;
			}
			return MarginsCustom.LineSpacingAtPt;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.LineSpacingAtPt = value;
			}
		}
	}

	public bool NoSpaceBetweenParagraphsOfSameStyle
	{
		get
		{
			if (MarginsCustom == null)
			{
				return false;
			}
			return MarginsCustom.NoSpaceBetweenParagraphsOfSameStyle;
		}
		set
		{
			if (MarginsCustom != null)
			{
				MarginsCustom.NoSpaceBetweenParagraphsOfSameStyle = value;
			}
		}
	}

	public float SizeQuocHieu
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 12f;
			}
			return TypographyCustom.SizeQuocHieu;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeQuocHieu = (float)Math.Round(value);
			}
		}
	}

	public float SizeTieuNgu
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 13f;
			}
			return TypographyCustom.SizeTieuNgu;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeTieuNgu = (float)Math.Round(value);
			}
		}
	}

	public float SizeCoQuan
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 13f;
			}
			return TypographyCustom.SizeCoQuan;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeCoQuan = (float)Math.Round(value);
			}
		}
	}

	public float SizeSoKyHieu
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 13f;
			}
			return TypographyCustom.SizeSoKyHieu;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeSoKyHieu = (float)Math.Round(value);
			}
		}
	}

	public float SizeDiaDanh
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 14f;
			}
			return TypographyCustom.SizeDiaDanh;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeDiaDanh = (float)Math.Round(value);
			}
		}
	}

	public float SizeTenLoai
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 14f;
			}
			return TypographyCustom.SizeTenLoai;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeTenLoai = (float)Math.Round(value);
			}
		}
	}

	public float SizeTrichYeu
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 14f;
			}
			return TypographyCustom.SizeTrichYeu;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeTrichYeu = (float)Math.Round(value);
			}
		}
	}

	public float SizeTrichYeuCongVan
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 12f;
			}
			return TypographyCustom.SizeTrichYeuCongVan;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeTrichYeuCongVan = (float)Math.Round(value);
			}
		}
	}

	public float SizeThanBai
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 14f;
			}
			return TypographyCustom.SizeThanBai;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeThanBai = (float)Math.Round(value);
			}
		}
	}

	public float SizeNoiNhan
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 11f;
			}
			return TypographyCustom.SizeNoiNhan;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeNoiNhan = (float)Math.Round(value);
			}
		}
	}

	public float SizeChuKy
	{
		get
		{
			if (TypographyCustom == null)
			{
				return 14f;
			}
			return TypographyCustom.SizeChuKy;
		}
		set
		{
			if (TypographyCustom != null)
			{
				TypographyCustom.SizeChuKy = (float)Math.Round(value);
			}
		}
	}

	public VndReadingOptions GetVndReadingOptions()
	{
		VndReadingOptions vndReadingOptions = new VndReadingOptions();
		vndReadingOptions.FourAsTu = AmountFourAsTu;
		vndReadingOptions.CapitalizeFirst = AmountCapitalizeFirst;
		vndReadingOptions.ZeroTensAsLe = AmountZeroTensAsLe;
		vndReadingOptions.ThousandAsNgan = AmountThousandAsNgan;
		return vndReadingOptions;
	}

	public UserPersonalization()
	{
		AgenciesAdmin = new List<AgencyProfile>();
		AgenciesParty = new List<AgencyProfile>();
		AgenciesMilitary = new List<AgencyProfile>();
		AgenciesCustom = new List<AgencyProfile>();
		AgenciesYouth = new List<AgencyProfile>();
		SignersAdmin = new List<SignerProfile>();
		SignersParty = new List<SignerProfile>();
		SignersMilitary = new List<SignerProfile>();
		SignersCustom = new List<SignerProfile>();
		SignersYouth = new List<SignerProfile>();
		Signers = new List<SignerProfile>();
		Recipients = new List<RecipientProfile>();
		ActiveRulePack = RulePackType.ND30_2020;
		AgencyStandard = 0;
		SignerStandard = 0;
		MarginStandard = 0;
		TypographyStandard = 0;
		UseCustomAgency = false;
		UseCustomSigner = false;
		UseCustomMargins = false;
		UseCustomTypography = false;
		UseCustomParagraph = false;
		AmountFourAsTu = false;
		AmountCapitalizeFirst = true;
		AmountZeroTensAsLe = false;
		AmountThousandAsNgan = false;
		MarginsAdmin = MarginProfile.CreateND30Default();
		MarginsParty = MarginProfile.CreateHD05Default();
		MarginsMilitary = MarginProfile.CreateTT68Default();
		MarginsCustom = MarginProfile.CreateND30Default();
		MarginsYouth = MarginProfile.CreateHD29Default();
		TypographyAdmin = TypographyProfile.CreateND30Default();
		TypographyParty = TypographyProfile.CreateHD05Default();
		TypographyMilitary = TypographyProfile.CreateTT68Default();
		TypographyCustom = TypographyProfile.CreateND30Default();
		TypographyYouth = TypographyProfile.CreateHD29Default();
	}

	public MarginProfile GetMarginsByStandard(int standard)
	{
		return standard switch
		{
			1 => MarginsParty ?? (MarginsParty = MarginProfile.CreateHD05Default()), 
			2 => MarginsMilitary ?? (MarginsMilitary = MarginProfile.CreateTT68Default()), 
			3 => MarginsCustom ?? (MarginsCustom = MarginProfile.CreateND30Default()), 
			4 => MarginsYouth ?? (MarginsYouth = MarginProfile.CreateHD29Default()), 
			_ => MarginsAdmin ?? (MarginsAdmin = MarginProfile.CreateND30Default()), 
		};
	}

	public MarginProfile GetActiveMargins(RulePackType activeRulePack)
	{
		if (UseCustomMargins)
		{
			return MarginsCustom ?? (MarginsCustom = MarginProfile.CreateND30Default());
		}
		return activeRulePack switch
		{
			RulePackType.HD05_VPTW => MarginsParty ?? (MarginsParty = MarginProfile.CreateHD05Default()), 
			RulePackType.TT68_2023_BQP => MarginsMilitary ?? (MarginsMilitary = MarginProfile.CreateTT68Default()), 
			RulePackType.HD29_TWDTN => MarginsYouth ?? (MarginsYouth = MarginProfile.CreateHD29Default()), 
			_ => MarginsAdmin ?? (MarginsAdmin = MarginProfile.CreateND30Default()), 
		};
	}

	public TypographyProfile GetTypographyByStandard(int standard)
	{
		return standard switch
		{
			1 => TypographyParty ?? (TypographyParty = TypographyProfile.CreateHD05Default()), 
			2 => TypographyMilitary ?? (TypographyMilitary = TypographyProfile.CreateTT68Default()), 
			3 => TypographyCustom ?? (TypographyCustom = TypographyProfile.CreateND30Default()), 
			4 => TypographyYouth ?? (TypographyYouth = TypographyProfile.CreateHD29Default()), 
			_ => TypographyAdmin ?? (TypographyAdmin = TypographyProfile.CreateND30Default()), 
		};
	}

	public TypographyProfile GetActiveTypography(RulePackType activeRulePack)
	{
		if (UseCustomTypography)
		{
			return TypographyCustom ?? (TypographyCustom = TypographyProfile.CreateND30Default());
		}
		return activeRulePack switch
		{
			RulePackType.HD05_VPTW => TypographyParty ?? (TypographyParty = TypographyProfile.CreateHD05Default()), 
			RulePackType.TT68_2023_BQP => TypographyMilitary ?? (TypographyMilitary = TypographyProfile.CreateTT68Default()), 
			RulePackType.HD29_TWDTN => TypographyYouth ?? (TypographyYouth = TypographyProfile.CreateHD29Default()), 
			_ => TypographyAdmin ?? (TypographyAdmin = TypographyProfile.CreateND30Default()), 
		};
	}

	public List<AgencyProfile> GetAgenciesByRulePack(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => AgenciesParty, 
			RulePackType.TT68_2023_BQP => AgenciesMilitary, 
			RulePackType.HD29_TWDTN => AgenciesYouth, 
			_ => AgenciesAdmin, 
		};
	}

	public List<AgencyProfile> GetAgenciesByStandard(int standard)
	{
		return standard switch
		{
			1 => AgenciesParty, 
			2 => AgenciesMilitary, 
			3 => AgenciesCustom, 
			4 => AgenciesYouth, 
			_ => AgenciesAdmin, 
		};
	}

	public List<SignerProfile> GetSignersByRulePack(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => SignersParty, 
			RulePackType.TT68_2023_BQP => SignersMilitary, 
			RulePackType.HD29_TWDTN => SignersYouth, 
			_ => SignersAdmin, 
		};
	}

	public List<SignerProfile> GetSignersByStandard(int standard)
	{
		return standard switch
		{
			1 => SignersParty, 
			2 => SignersMilitary, 
			3 => SignersCustom, 
			4 => SignersYouth, 
			_ => SignersAdmin, 
		};
	}

	public AgencyProfile GetSelectedAgency()
	{
		if (UseCustomAgency && AgenciesCustom != null && AgenciesCustom.Count > 0)
		{
			if (!string.IsNullOrEmpty(SelectedAgencyId))
			{
				AgencyProfile agencyProfile = AgenciesCustom.FirstOrDefault((AgencyProfile a) => a.Id == SelectedAgencyId);
				if (agencyProfile != null)
				{
					return agencyProfile;
				}
			}
			return AgenciesCustom.FirstOrDefault();
		}
		List<AgencyProfile> agenciesByRulePack = GetAgenciesByRulePack(ActiveRulePack);
		if (!string.IsNullOrEmpty(SelectedAgencyId))
		{
			AgencyProfile agencyProfile2 = agenciesByRulePack.FirstOrDefault((AgencyProfile a) => a.Id == SelectedAgencyId);
			if (agencyProfile2 != null)
			{
				return agencyProfile2;
			}
		}
		return agenciesByRulePack.FirstOrDefault();
	}

	public SignerProfile GetSelectedSigner()
	{
		if (UseCustomSigner && SignersCustom != null && SignersCustom.Count > 0)
		{
			if (!string.IsNullOrEmpty(SelectedSignerId))
			{
				SignerProfile signerProfile = SignersCustom.FirstOrDefault((SignerProfile s) => s.Id == SelectedSignerId);
				if (signerProfile != null)
				{
					return signerProfile;
				}
			}
			return SignersCustom.FirstOrDefault();
		}
		List<SignerProfile> signersByRulePack = GetSignersByRulePack(ActiveRulePack);
		if (signersByRulePack != null && signersByRulePack.Count > 0)
		{
			if (!string.IsNullOrEmpty(SelectedSignerId))
			{
				SignerProfile signerProfile2 = signersByRulePack.FirstOrDefault((SignerProfile s) => s.Id == SelectedSignerId);
				if (signerProfile2 != null)
				{
					return signerProfile2;
				}
			}
			return signersByRulePack.FirstOrDefault();
		}
		if (!string.IsNullOrEmpty(SelectedSignerId))
		{
			SignerProfile signerProfile3 = Signers.FirstOrDefault((SignerProfile s) => s.Id == SelectedSignerId);
			if (signerProfile3 != null)
			{
				return signerProfile3;
			}
		}
		return Signers.FirstOrDefault((SignerProfile s) => s.RulePack == ActiveRulePack) ?? Signers.FirstOrDefault();
	}

	public RecipientProfile GetSelectedRecipient()
	{
		SignerProfile selectedSigner = GetSelectedSigner();
		if (selectedSigner != null && !string.IsNullOrEmpty(selectedSigner.RecipientList))
		{
			return new RecipientProfile(selectedSigner.Id + "_rec", selectedSigner.Alias ?? "Nơi nhận", selectedSigner.RecipientList, selectedSigner.RulePack);
		}
		if (!string.IsNullOrEmpty(SelectedRecipientId))
		{
			RecipientProfile recipientProfile = Recipients.FirstOrDefault((RecipientProfile r) => r.Id == SelectedRecipientId);
			if (recipientProfile != null)
			{
				return recipientProfile;
			}
		}
		return Recipients.FirstOrDefault((RecipientProfile r) => r.RulePack == ActiveRulePack) ?? Recipients.FirstOrDefault();
	}

	public static UserPersonalization CreateDefault()
	{
		UserPersonalization userPersonalization = new UserPersonalization();
		userPersonalization.AgenciesAdmin.Add(new AgencyProfile("admin_bo", "Cấp Bộ (Bộ Nội vụ)", "", "BỘ NỘI VỤ", "Hà Nội", "BNV"));
		userPersonalization.AgenciesAdmin.Add(new AgencyProfile("admin_cuc", "Cục trực thuộc Bộ (Cục VTLT)", "BỘ NỘI VỤ", "CỤC VĂN THƯ VÀ LƯU TRỮ NHÀ NƯỚC", "Hà Nội", "VTLT"));
		userPersonalization.AgenciesAdmin.Add(new AgencyProfile("admin_so", "Sở Khoa học và Công nghệ", "UBND TỈNH...", "SỞ KHOA HỌC VÀ CÔNG NGHỆ", "...", "SKHCN"));
		userPersonalization.AgenciesAdmin.Add(new AgencyProfile("admin_ubnd", "UBND Xã/Phường", "UBND TỈNH...", "UBND XÃ...", "...", "UBND"));
		userPersonalization.AgenciesAdmin.Add(new AgencyProfile("admin_vp", "Văn phòng UBND tỉnh", "UBND TỈNH...", "VĂN PHÒNG ỦY BAN NHÂN DÂN", "...", "VP"));
		userPersonalization.AgenciesParty.Add(new AgencyProfile("party_1", "Đảng ủy Các cơ quan Đảng", "TỈNH ỦY...", "ĐẢNG ỦY CÁC CƠ QUAN ĐẢNG TỈNH", "...", "NQ/DU", RulePackType.HD05_VPTW));
		userPersonalization.AgenciesParty.Add(new AgencyProfile("party_2", "Ban Thường vụ Tỉnh ủy", "TỈNH ỦY...", "BAN THƯỜNG VỤ", "...", "NQ/TU", RulePackType.HD05_VPTW));
		userPersonalization.AgenciesParty.Add(new AgencyProfile("party_3", "Đảng ủy Xã/Phường", "TỈNH ỦY...", "ĐẢNG ỦY XÃ...", "...", "NQ/DU", RulePackType.HD05_VPTW));
		userPersonalization.AgenciesParty.Add(new AgencyProfile("party_4", "Chi bộ Cơ sở", "ĐẢNG ỦY XÃ...", "CHI BỘ...", "...", "NQ/CB", RulePackType.HD05_VPTW));
		userPersonalization.AgenciesMilitary.Add(new AgencyProfile("mil_bo", "Cấp Bộ (Bộ Quốc phòng)", "", "BỘ QUỐC PHÒNG", "Hà Nội", "BQP", RulePackType.TT68_2023_BQP));
		userPersonalization.AgenciesMilitary.Add(new AgencyProfile("mil_bttm", "Bộ Tổng Tham mưu", "BỘ QUỐC PHÒNG", "BỘ TỔNG THAM MƯU", "Hà Nội", "TM", RulePackType.TT68_2023_BQP));
		userPersonalization.AgenciesMilitary.Add(new AgencyProfile("mil_tcct", "Tổng cục Chính trị", "BỘ QUỐC PHÒNG", "TỔNG CỤC CHÍNH TRỊ", "Hà Nội", "CT", RulePackType.TT68_2023_BQP));
		userPersonalization.AgenciesMilitary.Add(new AgencyProfile("mil_qk", "Quân khu 7", "BỘ QUỐC PHÒNG", "BỘ TƯ LỆNH QUÂN KHU 7", "TP. Hồ Chí Minh", "QK", RulePackType.TT68_2023_BQP));
		userPersonalization.AgenciesMilitary.Add(new AgencyProfile("mil_bchqs", "Bộ CHQS tỉnh", "QUÂN KHU 7", "BỘ CHỈ HUY QUÂN SỰ TỈNH...", "...", "CHQS", RulePackType.TT68_2023_BQP));
		userPersonalization.AgenciesMilitary.Add(new AgencyProfile("mil_bchqs_xa", "Ban CHQS Xã/Phường", "BỘ CHQS TỈNH...", "BAN CHỈ HUY QUÂN SỰ\nXÃ...", "...", "KH-BCH", RulePackType.TT68_2023_BQP));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_tw", "BCH Trung ương Đoàn", "", "BAN CHẤP HÀNH TRUNG ƯƠNG", "Hà Nội", "TWĐTN", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_tinh", "BCH Đoàn tỉnh / thành phố", "", "BCH ĐOÀN TỈNH...", "...", "TĐTN", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_ban", "Ban tham mưu Tỉnh đoàn", "BCH ĐOÀN TỈNH...", "BAN TỔ CHỨC", "...", "BTC", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_phuong", "BCH Đoàn phường / đặc khu", "TỈNH ĐOÀN...", "BCH ĐOÀN PHƯỜNG...", "...", "ĐTN", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_coso", "BCH Đoàn xã (trực thuộc Tỉnh đoàn)", "TỈNH ĐOÀN...", "BCH ĐOÀN XÃ...", "...", "ĐTN", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_lcd", "BCH Liên chi đoàn (ký hiệu do Đoàn trường quy định)", "ĐOÀN TRƯỜNG...", "BCH LIÊN CHI ĐOÀN KHOA...", "...", "LCĐ", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_cd", "Chi đoàn", "ĐOÀN XÃ...", "CHI ĐOÀN...", "...", "CĐ", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesYouth.Add(new AgencyProfile("youth_dh", "Đại hội Đoàn (đại biểu)", "", "ĐẠI HỘI ĐẠI BIỂU\nĐOÀN TỈNH...\nNHIỆM KỲ...", "...", "ĐH", RulePackType.HD29_TWDTN));
		userPersonalization.AgenciesCustom.Add(new AgencyProfile("custom_so", "Mẫu 1: Sở Ban Ngành", "UBND TỈNH...", "SỞ KHOA HỌC VÀ CÔNG NGHỆ", "Hà Nội", "SKHCN"));
		userPersonalization.AgenciesCustom.Add(new AgencyProfile("custom_ubnd", "Mẫu 2: UBND Xã/Phường", "UBND TỈNH...", "UBND XÃ/PHƯỜNG...", "...", "UBND"));
		userPersonalization.AgenciesCustom.Add(new AgencyProfile("custom_cuc", "Mẫu 3: Cục trực thuộc", "BỘ NỘI VỤ", "CỤC VĂN THƯ VÀ LƯU TRỮ NHÀ NƯỚC", "Hà Nội", "VTLT"));
		userPersonalization.AgenciesCustom.Add(new AgencyProfile("custom_dn", "Mẫu 4: Doanh nghiệp / Tổ chức", "", "TẬP ĐOÀN / CÔNG TY...", "Hà Nội", "VP"));
		userPersonalization.SignersAdmin.Add(new SignerProfile("admin_sig_1", "Mẫu 1: Cấp Bộ (Bộ trưởng)", "BỘ TRƯỞNG", "BỘ TRƯỞNG", "Nguyễn Văn A", "Thủ trưởng cấp Bộ", "", RulePackType.ND30_2020, "- Như Điều ...;\n- Thủ tướng, các Phó Thủ tướng Chính phủ;\n- Văn phòng Trung ương Đảng và các Ban của Đảng;\n- Văn phòng Quốc hội, Văn phòng Chủ tịch nước;\n- Các Bộ, cơ quan ngang Bộ, cơ quan thuộc CP;\n- HĐND, UBND các tỉnh, thành phố trực thuộc TW;\n- Lưu: VT, PC.", "- Lưu: VT, PC."));
		userPersonalization.SignersAdmin.Add(new SignerProfile("admin_sig_2", "Mẫu 2: Ký thay Bộ trưởng (Thứ trưởng)", "KT. BỘ TRƯỞNG\nTHỨ TRƯỞNG", "THỨ TRƯỞNG", "Trần Văn B", "Ký thay Bộ trưởng", "", RulePackType.ND30_2020, "- Như Điều ...;\n- Bộ trưởng (để b/c);\n- Các Thứ trưởng;\n- Các đơn vị trực thuộc;\n- Lưu: VT, TH.", "- Lưu: VT, TH."));
		userPersonalization.SignersAdmin.Add(new SignerProfile("admin_sig_3", "Mẫu 3: Cấp Tỉnh (Chủ tịch UBND)", "TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH", "CHỦ TỊCH", "Trần Văn B", "Ký thay mặt UBND", "", RulePackType.ND30_2020, "- Như Điều ...;\n- Thường trực Tỉnh ủy (để b/c);\n- Thường trực HĐND tỉnh;\n- Chủ tịch, các PCT UBND tỉnh;\n- Các Sở, ban, ngành cấp tỉnh;\n- UBND các xã, phường;\n- Lưu: VT, TH.", "- Lưu: VT, TH."));
		userPersonalization.SignersAdmin.Add(new SignerProfile("admin_sig_4", "Mẫu 4: Cấp Sở (Giám đốc Sở)", "GIÁM ĐỐC", "GIÁM ĐỐC", "Lê Thị C", "Thủ trưởng cấp Sở", "", RulePackType.ND30_2020, "- Như Điều ...;\n- UBND tỉnh (để b/c);\n- Giám đốc, các Phó Giám đốc Sở;\n- Các phòng, ban, đơn vị trực thuộc;\n- UBND các xã, phường;\n- Lưu: VT, TH.", "- Lưu: VT, TH."));
		userPersonalization.SignersParty.Add(new SignerProfile("party_sig_1", "Mẫu 1: Bí thư Tỉnh ủy", "T/M BAN CHẤP HÀNH\nBÍ THƯ", "BÍ THƯ", "Lê Văn C", "Ký thay mặt Cấp ủy", "", RulePackType.HD05_VPTW, "- Văn phòng Trung ương Đảng (để b/c);\n- Thường trực Tỉnh ủy;\n- Các đồng chí Tỉnh ủy viên;\n- Các ban đảng, ban cán sự đảng, đảng đoàn;\n- Các đảng ủy xã, phường;\n- Lưu: Văn phòng Tỉnh ủy.", "- Lưu: Văn phòng Tỉnh ủy."));
		userPersonalization.SignersParty.Add(new SignerProfile("party_sig_2", "Mẫu 2: Ban Thường vụ (Phó Bí thư)", "T/M BAN THƯỜNG VỤ\nPHÓ BÍ THƯ", "PHÓ BÍ THƯ", "Phạm Văn D", "Ký thay Ban Thường vụ", "", RulePackType.HD05_VPTW, "- Thường trực Tỉnh ủy;\n- Các đồng chí Ủy viên Ban Thường vụ Tỉnh ủy;\n- Các ban đảng thuộc Tỉnh ủy;\n- Lưu: Văn phòng Tỉnh ủy.", "- Lưu: Văn phòng Tỉnh ủy."));
		userPersonalization.SignersParty.Add(new SignerProfile("party_sig_3", "Mẫu 3: Đảng ủy Xã/Phường", "T/M BAN CHẤP HÀNH\nBÍ THƯ", "BÍ THƯ", "Hoàng Văn E", "Bí thư Đảng ủy", "", RulePackType.HD05_VPTW, "- Thường trực Tỉnh ủy (để b/c);\n- Thường trực HĐND, UBND xã;\n- Các chi bộ trực thuộc;\n- Lưu: Văn phòng Đảng ủy.", "- Lưu: Văn phòng Đảng ủy."));
		userPersonalization.SignersParty.Add(new SignerProfile("party_sig_4", "Mẫu 4: Chi bộ Cơ sở (Bí thư)", "T/M CHI BỘ\nBÍ THƯ", "BÍ THƯ", "Vũ Văn F", "Bí thư Chi bộ", "", RulePackType.HD05_VPTW, "- Đảng ủy cấp trên (để b/c);\n- Các đồng chí đảng viên trong Chi bộ;\n- Lưu: Chi bộ.", "- Lưu: Chi bộ."));
		userPersonalization.SignersParty.Add(new SignerProfile("party_sig_5", "Mẫu 5: Văn phòng TW / Cấp ủy (K/T Chánh VP)", "K/T CHÁNH VĂN PHÒNG\nPHÓ CHÁNH VĂN PHÒNG", "PHÓ CHÁNH VĂN PHÒNG", "Lê Khánh Toàn", "Ký thay Chánh Văn phòng", "", RulePackType.HD05_VPTW, "- Các đồng chí Ủy viên Trung ương Đảng;\n- Các tỉnh ủy, thành ủy;\n- Văn phòng Trung ương Đảng và các ban đảng TW;\n- Lưu: Văn phòng Trung ương Đảng.", "- Lưu: Văn phòng Trung ương Đảng."));
		userPersonalization.SignersMilitary.Add(new SignerProfile("mil_sig_1", "Mẫu 1: Bộ Quốc phòng (Bộ trưởng)", "BỘ TRƯỞNG", "BỘ TRƯỞNG", "Phan Văn Giang", "Thủ trưởng Bộ Quốc phòng", "Đại tướng", RulePackType.TT68_2023_BQP, "- Ban Bí thư Trung ương Đảng (để b/c);\n- Thủ tướng Chính phủ (để b/c);\n- Thủ trưởng Bộ Quốc phòng;\n- Bộ Tổng Tham mưu, Tổng cục Chính trị;\n- Các quân khu, quân đoàn, quân chủng, binh chủng;\n- Lưu: VT, TH.", "- Lưu: VT, TH."));
		userPersonalization.SignersMilitary.Add(new SignerProfile("mil_sig_2", "Mẫu 2: Quân khu (Tư lệnh)", "TƯ LỆNH", "TƯ LỆNH", "Nguyễn Văn A", "Tư lệnh Quân khu", "Trung tướng", RulePackType.TT68_2023_BQP, "- Thủ trưởng Bộ Quốc phòng (để b/c);\n- Bộ Tổng Tham mưu, Tổng cục Chính trị;\n- Bộ Tư lệnh Quân khu 7;\n- Bộ CHQS các tỉnh, thành phố trực thuộc;\n- Các sư đoàn, lữ đoàn trực thuộc;\n- Lưu: VT, Tác chiến.", "- Lưu: VT, Tác chiến."));
		userPersonalization.SignersMilitary.Add(new SignerProfile("mil_sig_3", "Mẫu 3: Bộ CHQS tỉnh (Chỉ huy trưởng)", "CHỈ HUY TRƯỞNG", "CHỈ HUY TRƯỞNG", "Trần Văn B", "Chỉ huy trưởng", "Đại tá", RulePackType.TT68_2023_BQP, "- Bộ Tư lệnh Quân khu (để b/c);\n- Tỉnh ủy, UBND tỉnh (để b/c);\n- Thủ trưởng Bộ CHQS tỉnh;\n- Ban CHQS các xã, phường;\n- Các cơ quan, đơn vị trực thuộc;\n- Lưu: VT, Tác chiến.", "- Lưu: VT, Tác chiến."));
		userPersonalization.SignersMilitary.Add(new SignerProfile("mil_sig_4", "Mẫu 4: Cấp Trung đoàn (Trung đoàn trưởng)", "TRUNG ĐOÀN TRƯỞNG", "TRUNG ĐOÀN TRƯỞNG", "Lê Văn C", "Trung đoàn trưởng", "Thượng tá", RulePackType.TT68_2023_BQP, "- Thủ trưởng Sư đoàn / Bộ CHQS (để b/c);\n- Ban Chỉ huy Trung đoàn;\n- Các tiểu đoàn, đại đội trực thuộc;\n- Lưu: VT, Ban Tham mưu.", "- Lưu: VT, Ban Tham mưu."));
		userPersonalization.SignersMilitary.Add(new SignerProfile("mil_sig_5", "Mẫu 5: Ban CHQS Xã/Phường (Chỉ huy trưởng)", "CHỈ HUY TRƯỞNG", "CHỈ HUY TRƯỞNG", "Nguyễn Văn A", "Chỉ huy trưởng", "", RulePackType.TT68_2023_BQP, "- Bộ CHQS tỉnh (để b/c);\n- Đảng ủy, UBND xã (để b/c);\n- Các ban, ngành xã;\n- Các thôn, ấp, khu phố;\n- Lưu: VT.", "- Lưu: VT."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_1", "Mẫu 1: TM. Ban Thường vụ (Bí thư)", "TM. BAN THƯỜNG VỤ\nBÍ THƯ", "BÍ THƯ", "Nguyễn Văn A", "Thay mặt Ban Thường vụ", "", RulePackType.HD29_TWDTN, "- Trung ương Đoàn (để báo cáo);\n- Thường trực Tỉnh ủy (để báo cáo);\n- Các cơ sở Đoàn trực thuộc;\n- Lưu VP.", "- Lưu VP."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_2", "Mẫu 2: TM. Ban Thường vụ (Phó Bí thư thường trực)", "TM. BAN THƯỜNG VỤ\nPHÓ BÍ THƯ THƯỜNG TRỰC", "PHÓ BÍ THƯ THƯỜNG TRỰC", "Trần Thị B", "Phó Bí thư ký thay mặt", "", RulePackType.HD29_TWDTN, "- Như trên;\n- Thường trực Tỉnh đoàn;\n- Lưu VP.", "- Lưu VP."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_3", "Mẫu 3: TM. Ban Chấp hành Đoàn cơ sở", "TM. BAN CHẤP HÀNH\nBÍ THƯ", "BÍ THƯ", "Lê Văn C", "Bí thư Đoàn cơ sở", "", RulePackType.HD29_TWDTN, "- Tỉnh đoàn (để báo cáo);\n- Đảng ủy xã (để báo cáo);\n- Các chi đoàn trực thuộc;\n- Lưu VP.", "- Lưu VP."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_4", "Mẫu 4: TM. Ban Chấp hành Chi đoàn", "TM. BAN CHẤP HÀNH CHI ĐOÀN\nBÍ THƯ", "BÍ THƯ", "Phạm Thị D", "Bí thư Chi đoàn (có BCH)", "", RulePackType.HD29_TWDTN, "- Đoàn cấp trên (để báo cáo);\n- Chi bộ (để báo cáo);\n- Lưu Chi đoàn.", "- Lưu Chi đoàn."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_5", "Mẫu 5: TL. Ban Thường vụ (Chánh Văn phòng)", "TL. BAN THƯỜNG VỤ TỈNH ĐOÀN\nCHÁNH VĂN PHÒNG", "CHÁNH VĂN PHÒNG", "Nguyễn Văn T", "Chánh Văn phòng được ủy quyền", "", RulePackType.HD29_TWDTN, "- Như trên;\n- Lưu VP.", "- Lưu VP."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_6", "Mẫu 6: KT. Trưởng ban (Phó Trưởng ban)", "KT. TRƯỞNG BAN\nPHÓ TRƯỞNG BAN", "PHÓ TRƯỞNG BAN", "Nguyễn Văn K", "Phó Trưởng ban ký thay", "", RulePackType.HD29_TWDTN, "- Như trên;\n- Thường trực Tỉnh đoàn (để báo cáo);\n- Lưu VP, BTC.", "- Lưu VP, BTC."));
		userPersonalization.SignersYouth.Add(new SignerProfile("youth_sig_7", "Mẫu 7: TM. Đoàn Chủ tịch Đại hội", "TM. ĐOÀN CHỦ TỊCH", "", "Nguyễn Văn A", "Đoàn Chủ tịch Đại hội", "", RulePackType.HD29_TWDTN, "- Các đại biểu dự Đại hội;\n- Lưu Đại hội.", "- Lưu Đại hội."));
		foreach (SignerProfile item in userPersonalization.SignersAdmin)
		{
			userPersonalization.SignersCustom.Add(item.Clone());
		}
		foreach (SignerProfile item2 in userPersonalization.SignersCustom)
		{
			userPersonalization.Signers.Add(item2);
		}
		userPersonalization.Recipients.Add(new RecipientProfile("rec_bo", "Mẫu Nơi nhận Cấp Bộ / Ngang Bộ", "- Như Điều ...;\n- Thủ tướng, các Phó Thủ tướng Chính phủ;\n- Văn phòng Trung ương Đảng và các Ban của Đảng;\n- Văn phòng Quốc hội, Văn phòng Chủ tịch nước;\n- Các Bộ, cơ quan ngang Bộ, cơ quan thuộc CP;\n- HĐND, UBND các tỉnh, thành phố trực thuộc TW;\n- Lưu: VT, PC."));
		userPersonalization.Recipients.Add(new RecipientProfile("rec_cv", "Mẫu Nơi nhận Công văn thông dụng", "- Như kính gửi;\n- Giám đốc, các Phó Giám đốc;\n- Các phòng, ban, đơn vị trực thuộc;\n- UBND các xã, phường;\n- Lưu: VT, TH."));
		userPersonalization.Recipients.Add(new RecipientProfile("rec_dang", "Mẫu Nơi nhận Chi bộ Đảng", "- Đảng ủy cấp trên (để b/c);\n- Các đồng chí đảng viên trong Chi bộ;\n- Lưu: Chi bộ.", RulePackType.HD05_VPTW));
		userPersonalization.Recipients.Add(new RecipientProfile("rec_doan", "Mẫu Nơi nhận Văn bản Đoàn (HD 29)", "- Trung ương Đoàn (để báo cáo);\n- Thường trực Tỉnh ủy (để báo cáo);\n- Các cơ sở Đoàn trực thuộc;\n- Lưu VP.", RulePackType.HD29_TWDTN));
		userPersonalization.Recipients.Add(new RecipientProfile("rec_mil", "Mẫu Nơi nhận Đơn vị Quân đội (TT 68)", "- Như Điều ...;\n- Thủ trưởng Bộ Quốc phòng (để b/c);\n- Bộ Tổng Tham mưu, Tổng cục Chính trị (để b/c);\n- Các cơ quan, đơn vị trực thuộc;\n- Lưu: VT, ...;", RulePackType.TT68_2023_BQP));
		userPersonalization.SelectedAgencyId = "admin_bo";
		userPersonalization.SelectedSignerId = "admin_sig_1";
		userPersonalization.SelectedRecipientId = "rec_bo";
		return userPersonalization;
	}
}
