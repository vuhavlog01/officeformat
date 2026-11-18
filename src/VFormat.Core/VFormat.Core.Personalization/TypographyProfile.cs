using System;
using VFormat.Core.Enums;
using VFormat.Core.Rules;

namespace VFormat.Core.Personalization;

public class TypographyProfile
{
	private float _sizeQuocHieu = 12f;

	private float _sizeTieuNgu = 13f;

	private float _sizeCoQuan = 13f;

	private float _sizeSoKyHieu = 13f;

	private float _sizeDiaDanh = 14f;

	private float _sizeTenLoai = 14f;

	private float _sizeTrichYeu = 14f;

	private float _sizeTrichYeuCongVan = 12f;

	private float _sizeThanBai = 14f;

	private float _sizeNoiNhan = 11f;

	private float _sizeChuKy = 14f;

	public float SizeQuocHieu
	{
		get
		{
			return _sizeQuocHieu;
		}
		set
		{
			_sizeQuocHieu = (float)Math.Round(value);
		}
	}

	public float SizeTieuNgu
	{
		get
		{
			return _sizeTieuNgu;
		}
		set
		{
			_sizeTieuNgu = (float)Math.Round(value);
		}
	}

	public float SizeCoQuan
	{
		get
		{
			return _sizeCoQuan;
		}
		set
		{
			_sizeCoQuan = (float)Math.Round(value);
		}
	}

	public float SizeSoKyHieu
	{
		get
		{
			return _sizeSoKyHieu;
		}
		set
		{
			_sizeSoKyHieu = (float)Math.Round(value);
		}
	}

	public float SizeDiaDanh
	{
		get
		{
			return _sizeDiaDanh;
		}
		set
		{
			_sizeDiaDanh = (float)Math.Round(value);
		}
	}

	public float SizeTenLoai
	{
		get
		{
			return _sizeTenLoai;
		}
		set
		{
			_sizeTenLoai = (float)Math.Round(value);
		}
	}

	public float SizeTrichYeu
	{
		get
		{
			return _sizeTrichYeu;
		}
		set
		{
			_sizeTrichYeu = (float)Math.Round(value);
		}
	}

	public float SizeTrichYeuCongVan
	{
		get
		{
			return _sizeTrichYeuCongVan;
		}
		set
		{
			_sizeTrichYeuCongVan = (float)Math.Round(value);
		}
	}

	public float SizeThanBai
	{
		get
		{
			return _sizeThanBai;
		}
		set
		{
			_sizeThanBai = (float)Math.Round(value);
		}
	}

	public float SizeNoiNhan
	{
		get
		{
			return _sizeNoiNhan;
		}
		set
		{
			_sizeNoiNhan = (float)Math.Round(value);
		}
	}

	public float SizeChuKy
	{
		get
		{
			return _sizeChuKy;
		}
		set
		{
			_sizeChuKy = (float)Math.Round(value);
		}
	}

	public bool BoldQuocHieu { get; set; }

	public bool ItalicQuocHieu { get; set; }

	public bool BoldTieuNgu { get; set; }

	public bool ItalicTieuNgu { get; set; }

	public bool BoldCoQuan { get; set; }

	public bool ItalicCoQuan { get; set; }

	public bool BoldSoKyHieu { get; set; }

	public bool ItalicSoKyHieu { get; set; }

	public bool BoldDiaDanh { get; set; }

	public bool ItalicDiaDanh { get; set; }

	public bool BoldTenLoai { get; set; }

	public bool ItalicTenLoai { get; set; }

	public bool BoldTrichYeu { get; set; }

	public bool ItalicTrichYeu { get; set; }

	public bool BoldTrichYeuCongVan { get; set; }

	public bool ItalicTrichYeuCongVan { get; set; }

	public bool BoldThanBai { get; set; }

	public bool ItalicThanBai { get; set; }

	public bool BoldNoiNhan { get; set; }

	public bool ItalicNoiNhan { get; set; }

	public bool BoldChuKy { get; set; }

	public bool ItalicChuKy { get; set; }

	public TypographyProfile()
	{
		SizeQuocHieu = 12f;
		SizeTieuNgu = 13f;
		SizeCoQuan = 13f;
		SizeSoKyHieu = 13f;
		SizeDiaDanh = 14f;
		SizeTenLoai = 14f;
		SizeTrichYeu = 14f;
		SizeTrichYeuCongVan = 12f;
		SizeThanBai = 14f;
		SizeNoiNhan = 11f;
		SizeChuKy = 14f;
		BoldQuocHieu = true;
		ItalicQuocHieu = false;
		BoldTieuNgu = true;
		ItalicTieuNgu = false;
		BoldCoQuan = true;
		ItalicCoQuan = false;
		BoldSoKyHieu = false;
		ItalicSoKyHieu = false;
		BoldDiaDanh = false;
		ItalicDiaDanh = true;
		BoldTenLoai = true;
		ItalicTenLoai = false;
		BoldTrichYeu = true;
		ItalicTrichYeu = false;
		BoldTrichYeuCongVan = false;
		ItalicTrichYeuCongVan = false;
		BoldThanBai = false;
		ItalicThanBai = false;
		BoldNoiNhan = false;
		ItalicNoiNhan = false;
		BoldChuKy = true;
		ItalicChuKy = false;
	}

	public TypographyProfile Clone()
	{
		TypographyProfile typographyProfile = new TypographyProfile();
		typographyProfile.SizeQuocHieu = SizeQuocHieu;
		typographyProfile.SizeTieuNgu = SizeTieuNgu;
		typographyProfile.SizeCoQuan = SizeCoQuan;
		typographyProfile.SizeSoKyHieu = SizeSoKyHieu;
		typographyProfile.SizeDiaDanh = SizeDiaDanh;
		typographyProfile.SizeTenLoai = SizeTenLoai;
		typographyProfile.SizeTrichYeu = SizeTrichYeu;
		typographyProfile.SizeTrichYeuCongVan = SizeTrichYeuCongVan;
		typographyProfile.SizeThanBai = SizeThanBai;
		typographyProfile.SizeNoiNhan = SizeNoiNhan;
		typographyProfile.SizeChuKy = SizeChuKy;
		typographyProfile.BoldQuocHieu = BoldQuocHieu;
		typographyProfile.ItalicQuocHieu = ItalicQuocHieu;
		typographyProfile.BoldTieuNgu = BoldTieuNgu;
		typographyProfile.ItalicTieuNgu = ItalicTieuNgu;
		typographyProfile.BoldCoQuan = BoldCoQuan;
		typographyProfile.ItalicCoQuan = ItalicCoQuan;
		typographyProfile.BoldSoKyHieu = BoldSoKyHieu;
		typographyProfile.ItalicSoKyHieu = ItalicSoKyHieu;
		typographyProfile.BoldDiaDanh = BoldDiaDanh;
		typographyProfile.ItalicDiaDanh = ItalicDiaDanh;
		typographyProfile.BoldTenLoai = BoldTenLoai;
		typographyProfile.ItalicTenLoai = ItalicTenLoai;
		typographyProfile.BoldTrichYeu = BoldTrichYeu;
		typographyProfile.ItalicTrichYeu = ItalicTrichYeu;
		typographyProfile.BoldTrichYeuCongVan = BoldTrichYeuCongVan;
		typographyProfile.ItalicTrichYeuCongVan = ItalicTrichYeuCongVan;
		typographyProfile.BoldThanBai = BoldThanBai;
		typographyProfile.ItalicThanBai = ItalicThanBai;
		typographyProfile.BoldNoiNhan = BoldNoiNhan;
		typographyProfile.ItalicNoiNhan = ItalicNoiNhan;
		typographyProfile.BoldChuKy = BoldChuKy;
		typographyProfile.ItalicChuKy = ItalicChuKy;
		return typographyProfile;
	}

	public void ValidateAndClamp(RulePackType rulePack)
	{
		if (rulePack == RulePackType.HD05_VPTW || rulePack == RulePackType.HD29_TWDTN)
		{
			SizeQuocHieu = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.PartyHeader, SizeQuocHieu);
		}
		else
		{
			SizeQuocHieu = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.NationalHeader, SizeQuocHieu);
			SizeTieuNgu = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.Motto, SizeTieuNgu);
		}
		SizeCoQuan = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.IssuingAgency, SizeCoQuan);
		SizeSoKyHieu = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.DocumentNumber, SizeSoKyHieu);
		SizeDiaDanh = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.DocumentDate, SizeDiaDanh);
		SizeTenLoai = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.DocumentType, SizeTenLoai);
		SizeTrichYeu = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.SubjectExcerpt, SizeTrichYeu);
		SizeTrichYeuCongVan = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.DispatchSubject, SizeTrichYeuCongVan);
		SizeThanBai = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.BodyText, SizeThanBai);
		SizeNoiNhan = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.RecipientItem, SizeNoiNhan);
		SizeChuKy = LegalFontSizeStandards.EnsureLegalInteger(rulePack, SemanticRole.SignerName, SizeChuKy);
	}

	public static TypographyProfile CreateND30Default()
	{
		TypographyProfile typographyProfile = new TypographyProfile();
		typographyProfile.SizeQuocHieu = 12f;
		typographyProfile.SizeTieuNgu = 13f;
		typographyProfile.SizeCoQuan = 13f;
		typographyProfile.SizeSoKyHieu = 13f;
		typographyProfile.SizeDiaDanh = 14f;
		typographyProfile.SizeTenLoai = 14f;
		typographyProfile.SizeTrichYeu = 14f;
		typographyProfile.SizeTrichYeuCongVan = 12f;
		typographyProfile.SizeThanBai = 14f;
		typographyProfile.SizeNoiNhan = 11f;
		typographyProfile.SizeChuKy = 14f;
		typographyProfile.BoldQuocHieu = true;
		typographyProfile.ItalicQuocHieu = false;
		typographyProfile.BoldTieuNgu = true;
		typographyProfile.ItalicTieuNgu = false;
		typographyProfile.BoldCoQuan = true;
		typographyProfile.ItalicCoQuan = false;
		typographyProfile.BoldSoKyHieu = false;
		typographyProfile.ItalicSoKyHieu = false;
		typographyProfile.BoldDiaDanh = false;
		typographyProfile.ItalicDiaDanh = true;
		typographyProfile.BoldTenLoai = true;
		typographyProfile.ItalicTenLoai = false;
		typographyProfile.BoldTrichYeu = true;
		typographyProfile.ItalicTrichYeu = false;
		typographyProfile.BoldTrichYeuCongVan = false;
		typographyProfile.ItalicTrichYeuCongVan = false;
		typographyProfile.BoldThanBai = false;
		typographyProfile.ItalicThanBai = false;
		typographyProfile.BoldNoiNhan = false;
		typographyProfile.ItalicNoiNhan = false;
		typographyProfile.BoldChuKy = true;
		typographyProfile.ItalicChuKy = false;
		return typographyProfile;
	}

	public static TypographyProfile CreateHD05Default()
	{
		TypographyProfile typographyProfile = new TypographyProfile();
		typographyProfile.SizeQuocHieu = 15f;
		typographyProfile.SizeTieuNgu = 14f;
		typographyProfile.SizeCoQuan = 14f;
		typographyProfile.SizeSoKyHieu = 14f;
		typographyProfile.SizeDiaDanh = 14f;
		typographyProfile.SizeTenLoai = 15f;
		typographyProfile.SizeTrichYeu = 14f;
		typographyProfile.SizeTrichYeuCongVan = 12f;
		typographyProfile.SizeThanBai = 14f;
		typographyProfile.SizeNoiNhan = 12f;
		typographyProfile.SizeChuKy = 14f;
		typographyProfile.BoldQuocHieu = true;
		typographyProfile.ItalicQuocHieu = false;
		typographyProfile.BoldTieuNgu = true;
		typographyProfile.ItalicTieuNgu = false;
		typographyProfile.BoldCoQuan = true;
		typographyProfile.ItalicCoQuan = false;
		typographyProfile.BoldSoKyHieu = false;
		typographyProfile.ItalicSoKyHieu = false;
		typographyProfile.BoldDiaDanh = false;
		typographyProfile.ItalicDiaDanh = true;
		typographyProfile.BoldTenLoai = true;
		typographyProfile.ItalicTenLoai = false;
		typographyProfile.BoldTrichYeu = true;
		typographyProfile.ItalicTrichYeu = false;
		typographyProfile.BoldTrichYeuCongVan = false;
		typographyProfile.ItalicTrichYeuCongVan = true;
		typographyProfile.BoldThanBai = false;
		typographyProfile.ItalicThanBai = false;
		typographyProfile.BoldNoiNhan = false;
		typographyProfile.ItalicNoiNhan = false;
		typographyProfile.BoldChuKy = true;
		typographyProfile.ItalicChuKy = false;
		return typographyProfile;
	}

	public static TypographyProfile CreateTT68Default()
	{
		TypographyProfile typographyProfile = new TypographyProfile();
		typographyProfile.SizeQuocHieu = 13f;
		typographyProfile.SizeTieuNgu = 13f;
		typographyProfile.SizeCoQuan = 13f;
		typographyProfile.SizeSoKyHieu = 13f;
		typographyProfile.SizeDiaDanh = 13f;
		typographyProfile.SizeTenLoai = 14f;
		typographyProfile.SizeTrichYeu = 14f;
		typographyProfile.SizeTrichYeuCongVan = 12f;
		typographyProfile.SizeThanBai = 14f;
		typographyProfile.SizeNoiNhan = 11f;
		typographyProfile.SizeChuKy = 14f;
		typographyProfile.BoldQuocHieu = true;
		typographyProfile.ItalicQuocHieu = false;
		typographyProfile.BoldTieuNgu = true;
		typographyProfile.ItalicTieuNgu = false;
		typographyProfile.BoldCoQuan = true;
		typographyProfile.ItalicCoQuan = false;
		typographyProfile.BoldSoKyHieu = false;
		typographyProfile.ItalicSoKyHieu = false;
		typographyProfile.BoldDiaDanh = false;
		typographyProfile.ItalicDiaDanh = true;
		typographyProfile.BoldTenLoai = true;
		typographyProfile.ItalicTenLoai = false;
		typographyProfile.BoldTrichYeu = true;
		typographyProfile.ItalicTrichYeu = false;
		typographyProfile.BoldTrichYeuCongVan = false;
		typographyProfile.ItalicTrichYeuCongVan = false;
		typographyProfile.BoldThanBai = false;
		typographyProfile.ItalicThanBai = false;
		typographyProfile.BoldNoiNhan = false;
		typographyProfile.ItalicNoiNhan = false;
		typographyProfile.BoldChuKy = true;
		typographyProfile.ItalicChuKy = false;
		return typographyProfile;
	}

	public static TypographyProfile CreateHD29Default()
	{
		TypographyProfile typographyProfile = new TypographyProfile();
		typographyProfile.SizeQuocHieu = 15f;
		typographyProfile.SizeTieuNgu = 14f;
		typographyProfile.SizeCoQuan = 14f;
		typographyProfile.SizeSoKyHieu = 14f;
		typographyProfile.SizeDiaDanh = 13f;
		typographyProfile.SizeTenLoai = 16f;
		typographyProfile.SizeTrichYeu = 14f;
		typographyProfile.SizeTrichYeuCongVan = 12f;
		typographyProfile.SizeThanBai = 14f;
		typographyProfile.SizeNoiNhan = 11f;
		typographyProfile.SizeChuKy = 14f;
		typographyProfile.BoldQuocHieu = true;
		typographyProfile.ItalicQuocHieu = false;
		typographyProfile.BoldTieuNgu = true;
		typographyProfile.ItalicTieuNgu = false;
		typographyProfile.BoldCoQuan = true;
		typographyProfile.ItalicCoQuan = false;
		typographyProfile.BoldSoKyHieu = false;
		typographyProfile.ItalicSoKyHieu = false;
		typographyProfile.BoldDiaDanh = false;
		typographyProfile.ItalicDiaDanh = true;
		typographyProfile.BoldTenLoai = true;
		typographyProfile.ItalicTenLoai = false;
		typographyProfile.BoldTrichYeu = true;
		typographyProfile.ItalicTrichYeu = false;
		typographyProfile.BoldTrichYeuCongVan = false;
		typographyProfile.ItalicTrichYeuCongVan = true;
		typographyProfile.BoldThanBai = false;
		typographyProfile.ItalicThanBai = false;
		typographyProfile.BoldNoiNhan = false;
		typographyProfile.ItalicNoiNhan = false;
		typographyProfile.BoldChuKy = true;
		typographyProfile.ItalicChuKy = false;
		return typographyProfile;
	}
}
