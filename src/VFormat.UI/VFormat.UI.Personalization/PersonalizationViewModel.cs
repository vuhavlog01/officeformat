using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Input;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;
using VFormat.Core.Serialization;
using VFormat.UI.Common;

namespace VFormat.UI.Personalization;

public class PersonalizationViewModel : ViewModelBase
{
	private readonly JsonProfileRepository _repo;

	private UserPersonalization _personalization;

	private int _selectedSectionIndex;

	private int _selectedAgencyStandard;

	private bool _useCustomAgency;

	private int _agencySlotIndex;

	private string _agencyAlias;

	private string _agencyParent;

	private string _agencyName;

	private string _agencyLoc;

	private string _agencyPrefix;

	private int _selectedSignerStandard;

	private bool _useCustomSigner;

	private int _signerSlotIndex;

	private string _authorityTitle = "BỘ TRƯỞNG";

	private string _position = "BỘ TRƯỞNG";

	private string _militaryRank = "";

	private string _fullName = "Nguyễn Văn A";

	private string _recipientList = "- Như Điều ...;\n- Lưu: VT, PC.";

	private string _storage = "- Lưu: VT, PC.";

	private int _selectedMarginStandard;

	private bool _useCustomMargins;

	private bool _useCustomParagraph;

	private float _marginTopMm = 20f;

	private float _marginBottomMm = 20f;

	private float _marginLeftMm = 30f;

	private float _marginRightMm = 15f;

	private float _firstLineIndentCm = 1f;

	private string _firstLineIndentCmText = "1";

	private string _lineSpacingRule = "1.15 lines";

	private float _spacingBeforePt = 0f;

	private string _spacingBeforePtText = "0";

	private float _spacingAfterPt = 6f;

	private string _spacingAfterPtText = "6";

	private float _lineSpacingAtPt = 0f;

	private string _lineSpacingAtPtText = "0";

	private bool _noSpaceBetweenParagraphsOfSameStyle;

	public static readonly string[] LineSpacingOptions = new string[7]
	{
		"Single",
		"1.15 lines",
		"1.5 lines",
		"Double",
		"At least",
		"Exactly",
		"Multiple"
	};

	private int _selectedTypographyStandard;

	private bool _useCustomTypography;

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

	private bool _boldQuocHieu = true;

	private bool _italicQuocHieu;

	private bool _boldTieuNgu = true;

	private bool _italicTieuNgu;

	private bool _boldCoQuan = true;

	private bool _italicCoQuan;

	private bool _boldSoKyHieu;

	private bool _italicSoKyHieu;

	private bool _boldDiaDanh;

	private bool _italicDiaDanh = true;

	private bool _boldTenLoai = true;

	private bool _italicTenLoai;

	private bool _boldTrichYeu = true;

	private bool _italicTrichYeu;

	private bool _boldTrichYeuCongVan;

	private bool _italicTrichYeuCongVan;

	private bool _boldThanBai;

	private bool _italicThanBai;

	private bool _boldNoiNhan;

	private bool _italicNoiNhan;

	private bool _boldChuKy = true;

	private bool _italicChuKy;

	private string _statusMessage;

	private bool _isSaved;

	private bool _amountFourAsTu;

	private bool _amountCapitalizeFirst = true;

	private bool _amountZeroTensAsLe;

	private bool _amountThousandAsNgan;

	private bool _headerDateInTopCell;

	public bool HeaderDateInTopCell
	{
		get
		{
			return _headerDateInTopCell;
		}
		set
		{
			if (SetProperty(ref _headerDateInTopCell, value, "HeaderDateInTopCell"))
			{
				OnPropertyChanged("IsHeaderDateSeparateCell");
				OnPropertyChanged("IsHeaderDateInTopCell");
			}
		}
	}

	public bool IsHeaderDateSeparateCell
	{
		get
		{
			return !_headerDateInTopCell;
		}
		set
		{
			if (value)
			{
				HeaderDateInTopCell = false;
			}
		}
	}

	public bool IsHeaderDateInTopCell
	{
		get
		{
			return _headerDateInTopCell;
		}
		set
		{
			if (value)
			{
				HeaderDateInTopCell = true;
			}
		}
	}

	private string _amountPreviewInput = "21.105.024";

	public Action CloseAction { get; set; }

	public ObservableCollection<string> SlotOptions { get; set; }

	public int SelectedSectionIndex
	{
		get
		{
			return _selectedSectionIndex;
		}
		set
		{
			if (SetProperty(ref _selectedSectionIndex, value, "SelectedSectionIndex"))
			{
				OnPropertyChanged("IsSectionAgency");
				OnPropertyChanged("IsSectionSigner");
				OnPropertyChanged("IsSectionMargins");
				OnPropertyChanged("IsSectionTypography");
				OnPropertyChanged("IsSectionTextTools");
			}
		}
	}

	public bool IsSectionAgency
	{
		get
		{
			return _selectedSectionIndex == 0;
		}
		set
		{
			if (value)
			{
				SelectedSectionIndex = 0;
			}
		}
	}

	public bool IsSectionSigner
	{
		get
		{
			return _selectedSectionIndex == 1;
		}
		set
		{
			if (value)
			{
				SelectedSectionIndex = 1;
			}
		}
	}

	public bool IsSectionMargins
	{
		get
		{
			return _selectedSectionIndex == 2;
		}
		set
		{
			if (value)
			{
				SelectedSectionIndex = 2;
			}
		}
	}

	public bool IsSectionTypography
	{
		get
		{
			return _selectedSectionIndex == 3;
		}
		set
		{
			if (value)
			{
				SelectedSectionIndex = 3;
			}
		}
	}

	public bool IsSectionTextTools
	{
		get
		{
			return _selectedSectionIndex == 4;
		}
		set
		{
			if (value)
			{
				SelectedSectionIndex = 4;
			}
		}
	}

	public string AgencyBadgeText
	{
		get
		{
			if (!_useCustomAgency)
			{
				return "Theo quy chuẩn";
			}
			return "Mặc định: Bật";
		}
	}

	public string SignerBadgeText
	{
		get
		{
			if (!_useCustomSigner)
			{
				return "Theo quy chuẩn";
			}
			return "Mặc định: Bật";
		}
	}

	public string MarginBadgeText
	{
		get
		{
			if (!_useCustomMargins)
			{
				return "Theo quy chuẩn";
			}
			return "Mặc định: Bật";
		}
	}

	public string TypographyBadgeText
	{
		get
		{
			if (!_useCustomTypography)
			{
				return "Theo quy chuẩn";
			}
			return "Mặc định: Bật";
		}
	}

	public string AgencyStatusDetail
	{
		get
		{
			if (_useCustomAgency)
			{
				return "\ud83d\udfe2 ĐANG DÙNG LÀM MẶC ĐỊNH: Tất cả tính năng Chuẩn hóa tự động, Chèn tiêu đề sẽ áp dụng thông tin cơ quan cá nhân hóa đã thiết lập.";
			}
			return "⚪ BỎ MẶC ĐỊNH: Khi chuẩn hóa hoặc chèn tiêu đề sẽ áp dụng theo Quy chuẩn đang chọn trên thanh công cụ (NĐ 30 / HĐ 05 / TT 68 / HD 29).";
		}
	}

	public string SignerStatusDetail
	{
		get
		{
			if (_useCustomSigner)
			{
				return "\ud83d\udfe2 ĐANG DÙNG LÀM MẶC ĐỊNH: Tất cả tính năng Chuẩn hóa tự động, Chèn chữ ký & Nơi nhận sẽ áp dụng thông tin người ký cá nhân hóa đã thiết lập.";
			}
			return "⚪ BỎ MẶC ĐỊNH: Khi chuẩn hóa hoặc chèn chữ ký sẽ áp dụng theo Quy chuẩn đang chọn trên thanh công cụ (NĐ 30 / HĐ 05 / TT 68 / HD 29).";
		}
	}

	public string MarginStatusDetail
	{
		get
		{
			if (_useCustomMargins)
			{
				return $"\ud83d\udfe2 ĐANG DÙNG LÀM MẶC ĐỊNH: Tất cả tính năng Căn lề tự động, Chuẩn hóa 1-Click, Rà soát thể thức sẽ áp dụng thông số này (Trên {_marginTopMm:0.#}mm • Dưới {_marginBottomMm:0.#}mm • Trái {_marginLeftMm:0.#}mm • Phải {_marginRightMm:0.#}mm).";
			}
			return "⚪ BỎ MẶC ĐỊNH: Khi chuẩn hóa văn bản sẽ tự động áp dụng thông số lề chuẩn theo Quy chuẩn đang chọn trên thanh công cụ (NĐ 30 / HĐ 05 / TT 68 / HD 29).";
		}
	}

	public string TypographyStatusDetail
	{
		get
		{
			if (_useCustomTypography)
			{
				return "\ud83d\udfe2 ĐANG DÙNG LÀM MẶC ĐỊNH: Tất cả tính năng Chuẩn hóa tự động, Rà soát thể thức, Chuẩn hóa tiêu đề sẽ áp dụng các cỡ chữ thiết lập tại đây.";
			}
			return "⚪ BỎ MẶC ĐỊNH: Khi chuẩn hóa văn bản sẽ áp dụng cỡ chữ chuẩn theo Quy chuẩn đang chọn trên thanh công cụ (NĐ 30 / HĐ 05 / TT 68 / HD 29).";
		}
	}

	public Visibility AgencyDefaultVisibility
	{
		get
		{
			if (!IsAgencyCustom)
			{
				return Visibility.Collapsed;
			}
			return Visibility.Visible;
		}
	}

	public Visibility SignerDefaultVisibility
	{
		get
		{
			if (!IsSignerCustom)
			{
				return Visibility.Collapsed;
			}
			return Visibility.Visible;
		}
	}

	public Visibility MarginDefaultVisibility
	{
		get
		{
			if (!IsMarginCustom)
			{
				return Visibility.Collapsed;
			}
			return Visibility.Visible;
		}
	}

	public Visibility TypographyDefaultVisibility
	{
		get
		{
			if (!IsTypographyCustom)
			{
				return Visibility.Collapsed;
			}
			return Visibility.Visible;
		}
	}

	public int SelectedAgencyStandard
	{
		get
		{
			return _selectedAgencyStandard;
		}
		set
		{
			if (_selectedAgencyStandard != value)
			{
				SaveCurrentAgencySlot();
				_selectedAgencyStandard = value;
				OnPropertyChanged("SelectedAgencyStandard");
				OnPropertyChanged("IsAgencyND30");
				OnPropertyChanged("IsAgencyHD05");
				OnPropertyChanged("IsAgencyTT68");
				OnPropertyChanged("IsAgencyCustom");
				OnPropertyChanged("IsAgencyHD29");
				OnPropertyChanged("AgencyDefaultVisibility");
				LoadCurrentAgencySlot();
			}
		}
	}

	public bool IsAgencyND30
	{
		get
		{
			return _selectedAgencyStandard == 0;
		}
		set
		{
			if (value)
			{
				SelectedAgencyStandard = 0;
			}
		}
	}

	public bool IsAgencyHD05
	{
		get
		{
			return _selectedAgencyStandard == 1;
		}
		set
		{
			if (value)
			{
				SelectedAgencyStandard = 1;
			}
		}
	}

	public bool IsAgencyTT68
	{
		get
		{
			return _selectedAgencyStandard == 2;
		}
		set
		{
			if (value)
			{
				SelectedAgencyStandard = 2;
			}
		}
	}

	public bool IsAgencyCustom
	{
		get
		{
			return _selectedAgencyStandard == 3;
		}
		set
		{
			if (value)
			{
				SelectedAgencyStandard = 3;
			}
		}
	}

	public bool IsAgencyHD29
	{
		get
		{
			return _selectedAgencyStandard == 4;
		}
		set
		{
			if (value)
			{
				SelectedAgencyStandard = 4;
			}
		}
	}

	public bool UseCustomAgency
	{
		get
		{
			return _useCustomAgency;
		}
		set
		{
			if (SetProperty(ref _useCustomAgency, value, "UseCustomAgency"))
			{
				OnPropertyChanged("AgencyBadgeText");
				OnPropertyChanged("AgencyStatusDetail");
			}
		}
	}

	public int AgencySlotIndex
	{
		get
		{
			return _agencySlotIndex;
		}
		set
		{
			if (_agencySlotIndex != value)
			{
				SaveCurrentAgencySlot();
				_agencySlotIndex = value;
				OnPropertyChanged("AgencySlotIndex");
				LoadCurrentAgencySlot();
			}
		}
	}

	public int AdminSlotIndex
	{
		get
		{
			return AgencySlotIndex;
		}
		set
		{
			AgencySlotIndex = value;
		}
	}

	public int PartySlotIndex
	{
		get
		{
			return AgencySlotIndex;
		}
		set
		{
			AgencySlotIndex = value;
		}
	}

	public int MilitarySlotIndex
	{
		get
		{
			return AgencySlotIndex;
		}
		set
		{
			AgencySlotIndex = value;
		}
	}

	public int SelectedAgencyType
	{
		get
		{
			return SelectedAgencyStandard;
		}
		set
		{
			SelectedAgencyStandard = value;
		}
	}

	public bool IsAgencyAdmin
	{
		get
		{
			return IsAgencyND30;
		}
		set
		{
			IsAgencyND30 = value;
		}
	}

	public bool IsAgencyParty
	{
		get
		{
			return IsAgencyHD05;
		}
		set
		{
			IsAgencyHD05 = value;
		}
	}

	public bool IsAgencyMilitary
	{
		get
		{
			return IsAgencyTT68;
		}
		set
		{
			IsAgencyTT68 = value;
		}
	}

	public string AgencyAlias
	{
		get
		{
			return _agencyAlias;
		}
		set
		{
			SetProperty(ref _agencyAlias, value, "AgencyAlias");
		}
	}

	public string ParentAgency
	{
		get
		{
			return _agencyParent;
		}
		set
		{
			SetProperty(ref _agencyParent, value, "ParentAgency");
		}
	}

	public string AgencyName
	{
		get
		{
			return _agencyName;
		}
		set
		{
			SetProperty(ref _agencyName, value, "AgencyName");
		}
	}

	public string Location
	{
		get
		{
			return _agencyLoc;
		}
		set
		{
			SetProperty(ref _agencyLoc, value, "Location");
		}
	}

	public string CodePrefix
	{
		get
		{
			return _agencyPrefix;
		}
		set
		{
			SetProperty(ref _agencyPrefix, value, "CodePrefix");
		}
	}

	public string AdminAlias
	{
		get
		{
			return AgencyAlias;
		}
		set
		{
			AgencyAlias = value;
		}
	}

	public string AdminParent
	{
		get
		{
			return ParentAgency;
		}
		set
		{
			ParentAgency = value;
		}
	}

	public string AdminAgency
	{
		get
		{
			return AgencyName;
		}
		set
		{
			AgencyName = value;
		}
	}

	public string AdminLoc
	{
		get
		{
			return Location;
		}
		set
		{
			Location = value;
		}
	}

	public string AdminPrefix
	{
		get
		{
			return CodePrefix;
		}
		set
		{
			CodePrefix = value;
		}
	}

	public string PartyAlias
	{
		get
		{
			return AgencyAlias;
		}
		set
		{
			AgencyAlias = value;
		}
	}

	public string PartyParent
	{
		get
		{
			return ParentAgency;
		}
		set
		{
			ParentAgency = value;
		}
	}

	public string PartyAgency
	{
		get
		{
			return AgencyName;
		}
		set
		{
			AgencyName = value;
		}
	}

	public string PartyLoc
	{
		get
		{
			return Location;
		}
		set
		{
			Location = value;
		}
	}

	public string PartyPrefix
	{
		get
		{
			return CodePrefix;
		}
		set
		{
			CodePrefix = value;
		}
	}

	public string MilitaryAlias
	{
		get
		{
			return AgencyAlias;
		}
		set
		{
			AgencyAlias = value;
		}
	}

	public string MilitaryParent
	{
		get
		{
			return ParentAgency;
		}
		set
		{
			ParentAgency = value;
		}
	}

	public string MilitaryAgency
	{
		get
		{
			return AgencyName;
		}
		set
		{
			AgencyName = value;
		}
	}

	public string MilitaryLoc
	{
		get
		{
			return Location;
		}
		set
		{
			Location = value;
		}
	}

	public string MilitaryPrefix
	{
		get
		{
			return CodePrefix;
		}
		set
		{
			CodePrefix = value;
		}
	}

	public int SelectedSignerStandard
	{
		get
		{
			return _selectedSignerStandard;
		}
		set
		{
			if (_selectedSignerStandard != value)
			{
				SaveCurrentSignerSlot();
				_selectedSignerStandard = value;
				OnPropertyChanged("SelectedSignerStandard");
				OnPropertyChanged("IsSignerND30");
				OnPropertyChanged("IsSignerHD05");
				OnPropertyChanged("IsSignerTT68");
				OnPropertyChanged("IsSignerCustom");
				OnPropertyChanged("IsSignerHD29");
				OnPropertyChanged("SignerDefaultVisibility");
				LoadCurrentSignerSlot();
			}
		}
	}

	public bool IsSignerND30
	{
		get
		{
			return _selectedSignerStandard == 0;
		}
		set
		{
			if (value)
			{
				SelectedSignerStandard = 0;
			}
		}
	}

	public bool IsSignerHD05
	{
		get
		{
			return _selectedSignerStandard == 1;
		}
		set
		{
			if (value)
			{
				SelectedSignerStandard = 1;
			}
		}
	}

	public bool IsSignerTT68
	{
		get
		{
			return _selectedSignerStandard == 2;
		}
		set
		{
			if (value)
			{
				SelectedSignerStandard = 2;
			}
		}
	}

	public bool IsSignerCustom
	{
		get
		{
			return _selectedSignerStandard == 3;
		}
		set
		{
			if (value)
			{
				SelectedSignerStandard = 3;
			}
		}
	}

	public bool IsSignerHD29
	{
		get
		{
			return _selectedSignerStandard == 4;
		}
		set
		{
			if (value)
			{
				SelectedSignerStandard = 4;
			}
		}
	}

	public bool UseCustomSigner
	{
		get
		{
			return _useCustomSigner;
		}
		set
		{
			if (SetProperty(ref _useCustomSigner, value, "UseCustomSigner"))
			{
				OnPropertyChanged("SignerBadgeText");
				OnPropertyChanged("SignerStatusDetail");
			}
		}
	}

	public int SignerSlotIndex
	{
		get
		{
			return _signerSlotIndex;
		}
		set
		{
			if (_signerSlotIndex != value)
			{
				SaveCurrentSignerSlot();
				_signerSlotIndex = value;
				OnPropertyChanged("SignerSlotIndex");
				LoadCurrentSignerSlot();
			}
		}
	}

	public string AuthorityPrefix
	{
		get
		{
			return _authorityTitle;
		}
		set
		{
			SetProperty(ref _authorityTitle, value, "AuthorityPrefix");
		}
	}

	public string Position
	{
		get
		{
			return _position;
		}
		set
		{
			SetProperty(ref _position, value, "Position");
		}
	}

	public string MilitaryRank
	{
		get
		{
			return _militaryRank;
		}
		set
		{
			SetProperty(ref _militaryRank, value, "MilitaryRank");
		}
	}

	public string SignerName
	{
		get
		{
			return _fullName;
		}
		set
		{
			SetProperty(ref _fullName, value, "SignerName");
		}
	}

	public string RecipientList
	{
		get
		{
			return _recipientList;
		}
		set
		{
			SetProperty(ref _recipientList, value, "RecipientList");
		}
	}

	public string RecipientContent
	{
		get
		{
			return _recipientList;
		}
		set
		{
			RecipientList = value;
		}
	}

	public string Storage
	{
		get
		{
			return _storage;
		}
		set
		{
			SetProperty(ref _storage, value, "Storage");
		}
	}

	public int SelectedMarginStandard
	{
		get
		{
			return _selectedMarginStandard;
		}
		set
		{
			if (_selectedMarginStandard != value)
			{
				SaveCurrentMarginStandard();
				_selectedMarginStandard = value;
				OnPropertyChanged("SelectedMarginStandard");
				OnPropertyChanged("IsMarginND30");
				OnPropertyChanged("IsMarginHD05");
				OnPropertyChanged("IsMarginTT68");
				OnPropertyChanged("IsMarginCustom");
				OnPropertyChanged("IsMarginHD29");
				OnPropertyChanged("MarginDefaultVisibility");
				OnPropertyChanged("ResetMarginStandardText");
				LoadCurrentMarginStandard();
			}
		}
	}

	public bool IsMarginND30
	{
		get
		{
			return _selectedMarginStandard == 0;
		}
		set
		{
			if (value)
			{
				SelectedMarginStandard = 0;
			}
		}
	}

	public bool IsMarginHD05
	{
		get
		{
			return _selectedMarginStandard == 1;
		}
		set
		{
			if (value)
			{
				SelectedMarginStandard = 1;
			}
		}
	}

	public bool IsMarginTT68
	{
		get
		{
			return _selectedMarginStandard == 2;
		}
		set
		{
			if (value)
			{
				SelectedMarginStandard = 2;
			}
		}
	}

	public bool IsMarginCustom
	{
		get
		{
			return _selectedMarginStandard == 3;
		}
		set
		{
			if (value)
			{
				SelectedMarginStandard = 3;
			}
		}
	}

	public bool IsMarginHD29
	{
		get
		{
			return _selectedMarginStandard == 4;
		}
		set
		{
			if (value)
			{
				SelectedMarginStandard = 4;
			}
		}
	}

	public bool UseCustomMargins
	{
		get
		{
			return _useCustomMargins;
		}
		set
		{
			if (SetProperty(ref _useCustomMargins, value, "UseCustomMargins"))
			{
				OnPropertyChanged("MarginBadgeText");
				OnPropertyChanged("MarginStatusDetail");
			}
		}
	}

	public bool UseCustomParagraph
	{
		get
		{
			return _useCustomParagraph;
		}
		set
		{
			SetProperty(ref _useCustomParagraph, value, "UseCustomParagraph");
		}
	}

	public bool AmountFourAsBon
	{
		get
		{
			return !_amountFourAsTu;
		}
		set
		{
			if (value)
			{
				AmountFourAsTu = false;
			}
		}
	}

	public bool AmountFourAsTu
	{
		get
		{
			return _amountFourAsTu;
		}
		set
		{
			if (SetProperty(ref _amountFourAsTu, value, "AmountFourAsTu"))
			{
				OnPropertyChanged("AmountFourAsBon");
				RaiseAmountPreviewChanged();
			}
		}
	}

	public bool AmountCapitalizeFirst
	{
		get
		{
			return _amountCapitalizeFirst;
		}
		set
		{
			if (SetProperty(ref _amountCapitalizeFirst, value, "AmountCapitalizeFirst"))
			{
				RaiseAmountPreviewChanged();
			}
		}
	}

	public bool AmountZeroTensAsLinh
	{
		get
		{
			return !_amountZeroTensAsLe;
		}
		set
		{
			if (value)
			{
				AmountZeroTensAsLe = false;
			}
		}
	}

	public bool AmountZeroTensAsLe
	{
		get
		{
			return _amountZeroTensAsLe;
		}
		set
		{
			if (SetProperty(ref _amountZeroTensAsLe, value, "AmountZeroTensAsLe"))
			{
				OnPropertyChanged("AmountZeroTensAsLinh");
				RaiseAmountPreviewChanged();
			}
		}
	}

	public bool AmountThousandAsNghin
	{
		get
		{
			return !_amountThousandAsNgan;
		}
		set
		{
			if (value)
			{
				AmountThousandAsNgan = false;
			}
		}
	}

	public bool AmountThousandAsNgan
	{
		get
		{
			return _amountThousandAsNgan;
		}
		set
		{
			if (SetProperty(ref _amountThousandAsNgan, value, "AmountThousandAsNgan"))
			{
				OnPropertyChanged("AmountThousandAsNghin");
				RaiseAmountPreviewChanged();
			}
		}
	}

	public string AmountPreviewInput
	{
		get
		{
			return _amountPreviewInput;
		}
		set
		{
			if (SetProperty(ref _amountPreviewInput, value, "AmountPreviewInput"))
			{
				OnPropertyChanged("AmountPreviewText");
			}
		}
	}

	public string AmountPreviewText
	{
		get
		{
			if (string.IsNullOrWhiteSpace(_amountPreviewInput))
			{
				return "Nhập một số tiền để xem trước.";
			}
			if (!VndAmountReader.TryConvert(_amountPreviewInput, CurrentVndReadingOptions(), out var words, out var error))
			{
				return error;
			}
			return words;
		}
	}

	public string TextToolsBadgeText => string.Format("Đọc số: {0} · {1} · {2}", _amountFourAsTu ? "tư" : "bốn", _amountZeroTensAsLe ? "lẻ" : "linh", _amountThousandAsNgan ? "ngàn" : "nghìn");

	public float MarginTopMm
	{
		get
		{
			return _marginTopMm;
		}
		set
		{
			if (SetProperty(ref _marginTopMm, value, "MarginTopMm"))
			{
				SaveCurrentMarginStandard();
				OnPropertyChanged("MarginStatusDetail");
			}
		}
	}

	public float MarginBottomMm
	{
		get
		{
			return _marginBottomMm;
		}
		set
		{
			if (SetProperty(ref _marginBottomMm, value, "MarginBottomMm"))
			{
				SaveCurrentMarginStandard();
				OnPropertyChanged("MarginStatusDetail");
			}
		}
	}

	public float MarginLeftMm
	{
		get
		{
			return _marginLeftMm;
		}
		set
		{
			if (SetProperty(ref _marginLeftMm, value, "MarginLeftMm"))
			{
				SaveCurrentMarginStandard();
				OnPropertyChanged("MarginStatusDetail");
			}
		}
	}

	public float MarginRightMm
	{
		get
		{
			return _marginRightMm;
		}
		set
		{
			if (SetProperty(ref _marginRightMm, value, "MarginRightMm"))
			{
				SaveCurrentMarginStandard();
				OnPropertyChanged("MarginStatusDetail");
			}
		}
	}

	public string FirstLineIndentCmText
	{
		get
		{
			return _firstLineIndentCmText;
		}
		set
		{
			if (SetProperty(ref _firstLineIndentCmText, value, "FirstLineIndentCmText") && TryParseFlexibleFloat(value, out var result))
			{
				if (result >= 50f)
				{
					result /= 100f;
				}
				_firstLineIndentCm = result;
				OnPropertyChanged("FirstLineIndentCm");
				SaveCurrentMarginStandard();
			}
		}
	}

	public float FirstLineIndentCm
	{
		get
		{
			return _firstLineIndentCm;
		}
		set
		{
			if (value >= 50f)
			{
				value /= 100f;
			}
			if (SetProperty(ref _firstLineIndentCm, value, "FirstLineIndentCm"))
			{
				_firstLineIndentCmText = FormatIndentDisplay(value, _firstLineIndentCmText);
				OnPropertyChanged("FirstLineIndentCmText");
				SaveCurrentMarginStandard();
			}
		}
	}

	public string SpacingAfterPtText
	{
		get
		{
			return _spacingAfterPtText;
		}
		set
		{
			if (SetProperty(ref _spacingAfterPtText, value, "SpacingAfterPtText") && TryParseFlexibleFloat(value, out var result))
			{
				_spacingAfterPt = result;
				OnPropertyChanged("SpacingAfterPt");
				SaveCurrentMarginStandard();
			}
		}
	}

	public float SpacingAfterPt
	{
		get
		{
			return _spacingAfterPt;
		}
		set
		{
			if (SetProperty(ref _spacingAfterPt, value, "SpacingAfterPt"))
			{
				_spacingAfterPtText = FormatIndentDisplay(value, _spacingAfterPtText);
				OnPropertyChanged("SpacingAfterPtText");
				SaveCurrentMarginStandard();
			}
		}
	}

	public string[] LineSpacingRuleList => LineSpacingOptions;

	public string SpacingBeforePtText
	{
		get
		{
			return _spacingBeforePtText;
		}
		set
		{
			if (SetProperty(ref _spacingBeforePtText, value, "SpacingBeforePtText") && TryParseFlexibleFloat(value, out var result))
			{
				_spacingBeforePt = result;
				OnPropertyChanged("SpacingBeforePt");
				SaveCurrentMarginStandard();
			}
		}
	}

	public float SpacingBeforePt
	{
		get
		{
			return _spacingBeforePt;
		}
		set
		{
			if (SetProperty(ref _spacingBeforePt, value, "SpacingBeforePt"))
			{
				_spacingBeforePtText = FormatIndentDisplay(value, _spacingBeforePtText);
				OnPropertyChanged("SpacingBeforePtText");
				SaveCurrentMarginStandard();
			}
		}
	}

	public string LineSpacingRule
	{
		get
		{
			return _lineSpacingRule;
		}
		set
		{
			if (SetProperty(ref _lineSpacingRule, value, "LineSpacingRule"))
			{
				OnPropertyChanged("IsLineSpacingAtEnabled");
				SaveCurrentMarginStandard();
			}
		}
	}

	public string LineSpacingAtPtText
	{
		get
		{
			return _lineSpacingAtPtText;
		}
		set
		{
			if (SetProperty(ref _lineSpacingAtPtText, value, "LineSpacingAtPtText") && TryParseFlexibleFloat(value, out var result))
			{
				_lineSpacingAtPt = result;
				OnPropertyChanged("LineSpacingAtPt");
				SaveCurrentMarginStandard();
			}
		}
	}

	public float LineSpacingAtPt
	{
		get
		{
			return _lineSpacingAtPt;
		}
		set
		{
			if (SetProperty(ref _lineSpacingAtPt, value, "LineSpacingAtPt"))
			{
				_lineSpacingAtPtText = FormatIndentDisplay(value, _lineSpacingAtPtText);
				OnPropertyChanged("LineSpacingAtPtText");
				SaveCurrentMarginStandard();
			}
		}
	}

	public bool IsLineSpacingAtEnabled
	{
		get
		{
			if (string.IsNullOrEmpty(_lineSpacingRule)) return true;
			return _lineSpacingRule.Equals("Exactly", StringComparison.OrdinalIgnoreCase) ||
			       _lineSpacingRule.Equals("At least", StringComparison.OrdinalIgnoreCase) ||
			       _lineSpacingRule.Equals("Multiple", StringComparison.OrdinalIgnoreCase) ||
			       _lineSpacingRule.StartsWith("Exactly", StringComparison.OrdinalIgnoreCase);
		}
	}

	public bool NoSpaceBetweenParagraphsOfSameStyle
	{
		get
		{
			return _noSpaceBetweenParagraphsOfSameStyle;
		}
		set
		{
			if (SetProperty(ref _noSpaceBetweenParagraphsOfSameStyle, value, "NoSpaceBetweenParagraphsOfSameStyle"))
			{
				SaveCurrentMarginStandard();
			}
		}
	}

	public string ResetMarginStandardText => _selectedMarginStandard switch
	{
		0 => "Đặt lại lề chuẩn Nghị định 30 (20 - 20 - 30 - 15 mm)", 
		1 => "Đặt lại lề & đoạn chuẩn Hướng dẫn 05 (20-20-30-15 mm, Before 6pt, After 6pt, Exactly 18pt)", 
		2 => "Đặt lại lề chuẩn Thông tư 68 (20 - 20 - 30 - 15 mm)", 
		4 => "Đặt lại lề chuẩn Hướng dẫn 29 - Đoàn (20 - 20 - 30 - 20 mm, thụt 1,27 cm)", 
		_ => "Đặt lại lề mục Cá nhân về chuẩn NĐ 30", 
	};

	public int SelectedTypographyStandard
	{
		get
		{
			return _selectedTypographyStandard;
		}
		set
		{
			if (_selectedTypographyStandard != value)
			{
				SaveCurrentTypographyStandard();
				_selectedTypographyStandard = value;
				OnPropertyChanged("SelectedTypographyStandard");
				OnPropertyChanged("IsTypographyND30");
				OnPropertyChanged("IsTypographyHD05");
				OnPropertyChanged("IsTypographyTT68");
				OnPropertyChanged("IsTypographyCustom");
				OnPropertyChanged("IsTypographyHD29");
				OnPropertyChanged("TypographyDefaultVisibility");
				OnPropertyChanged("ResetTypographyStandardText");
				NotifyTypographyLabels();
				LoadCurrentTypographyStandard();
			}
		}
	}

	public bool IsTypographyND30
	{
		get
		{
			return _selectedTypographyStandard == 0;
		}
		set
		{
			if (value)
			{
				SelectedTypographyStandard = 0;
			}
		}
	}

	public bool IsTypographyHD05
	{
		get
		{
			return _selectedTypographyStandard == 1;
		}
		set
		{
			if (value)
			{
				SelectedTypographyStandard = 1;
			}
		}
	}

	public bool IsTypographyTT68
	{
		get
		{
			return _selectedTypographyStandard == 2;
		}
		set
		{
			if (value)
			{
				SelectedTypographyStandard = 2;
			}
		}
	}

	public bool IsTypographyCustom
	{
		get
		{
			return _selectedTypographyStandard == 3;
		}
		set
		{
			if (value)
			{
				SelectedTypographyStandard = 3;
			}
		}
	}

	public bool IsTypographyHD29
	{
		get
		{
			return _selectedTypographyStandard == 4;
		}
		set
		{
			if (value)
			{
				SelectedTypographyStandard = 4;
			}
		}
	}

	public bool UseCustomTypography
	{
		get
		{
			return _useCustomTypography;
		}
		set
		{
			if (SetProperty(ref _useCustomTypography, value, "UseCustomTypography"))
			{
				OnPropertyChanged("TypographyBadgeText");
				OnPropertyChanged("TypographyStatusDetail");
			}
		}
	}

	public float SizeQuocHieu
	{
		get
		{
			return _sizeQuocHieu;
		}
		set
		{
			if (SetProperty(ref _sizeQuocHieu, (float)Math.Round(value), "SizeQuocHieu"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeTieuNgu, (float)Math.Round(value), "SizeTieuNgu"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeCoQuan, (float)Math.Round(value), "SizeCoQuan"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeSoKyHieu, (float)Math.Round(value), "SizeSoKyHieu"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeDiaDanh, (float)Math.Round(value), "SizeDiaDanh"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeTenLoai, (float)Math.Round(value), "SizeTenLoai"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeTrichYeu, (float)Math.Round(value), "SizeTrichYeu"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeTrichYeuCongVan, (float)Math.Round(value), "SizeTrichYeuCongVan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldQuocHieu
	{
		get
		{
			return _boldQuocHieu;
		}
		set
		{
			if (SetProperty(ref _boldQuocHieu, value, "BoldQuocHieu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicQuocHieu
	{
		get
		{
			return _italicQuocHieu;
		}
		set
		{
			if (SetProperty(ref _italicQuocHieu, value, "ItalicQuocHieu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldTieuNgu
	{
		get
		{
			return _boldTieuNgu;
		}
		set
		{
			if (SetProperty(ref _boldTieuNgu, value, "BoldTieuNgu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicTieuNgu
	{
		get
		{
			return _italicTieuNgu;
		}
		set
		{
			if (SetProperty(ref _italicTieuNgu, value, "ItalicTieuNgu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldCoQuan
	{
		get
		{
			return _boldCoQuan;
		}
		set
		{
			if (SetProperty(ref _boldCoQuan, value, "BoldCoQuan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicCoQuan
	{
		get
		{
			return _italicCoQuan;
		}
		set
		{
			if (SetProperty(ref _italicCoQuan, value, "ItalicCoQuan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldSoKyHieu
	{
		get
		{
			return _boldSoKyHieu;
		}
		set
		{
			if (SetProperty(ref _boldSoKyHieu, value, "BoldSoKyHieu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicSoKyHieu
	{
		get
		{
			return _italicSoKyHieu;
		}
		set
		{
			if (SetProperty(ref _italicSoKyHieu, value, "ItalicSoKyHieu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldDiaDanh
	{
		get
		{
			return _boldDiaDanh;
		}
		set
		{
			if (SetProperty(ref _boldDiaDanh, value, "BoldDiaDanh"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicDiaDanh
	{
		get
		{
			return _italicDiaDanh;
		}
		set
		{
			if (SetProperty(ref _italicDiaDanh, value, "ItalicDiaDanh"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldTenLoai
	{
		get
		{
			return _boldTenLoai;
		}
		set
		{
			if (SetProperty(ref _boldTenLoai, value, "BoldTenLoai"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicTenLoai
	{
		get
		{
			return _italicTenLoai;
		}
		set
		{
			if (SetProperty(ref _italicTenLoai, value, "ItalicTenLoai"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldTrichYeu
	{
		get
		{
			return _boldTrichYeu;
		}
		set
		{
			if (SetProperty(ref _boldTrichYeu, value, "BoldTrichYeu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicTrichYeu
	{
		get
		{
			return _italicTrichYeu;
		}
		set
		{
			if (SetProperty(ref _italicTrichYeu, value, "ItalicTrichYeu"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldTrichYeuCongVan
	{
		get
		{
			return _boldTrichYeuCongVan;
		}
		set
		{
			if (SetProperty(ref _boldTrichYeuCongVan, value, "BoldTrichYeuCongVan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicTrichYeuCongVan
	{
		get
		{
			return _italicTrichYeuCongVan;
		}
		set
		{
			if (SetProperty(ref _italicTrichYeuCongVan, value, "ItalicTrichYeuCongVan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldThanBai
	{
		get
		{
			return _boldThanBai;
		}
		set
		{
			if (SetProperty(ref _boldThanBai, value, "BoldThanBai"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicThanBai
	{
		get
		{
			return _italicThanBai;
		}
		set
		{
			if (SetProperty(ref _italicThanBai, value, "ItalicThanBai"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldNoiNhan
	{
		get
		{
			return _boldNoiNhan;
		}
		set
		{
			if (SetProperty(ref _boldNoiNhan, value, "BoldNoiNhan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicNoiNhan
	{
		get
		{
			return _italicNoiNhan;
		}
		set
		{
			if (SetProperty(ref _italicNoiNhan, value, "ItalicNoiNhan"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool BoldChuKy
	{
		get
		{
			return _boldChuKy;
		}
		set
		{
			if (SetProperty(ref _boldChuKy, value, "BoldChuKy"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public bool ItalicChuKy
	{
		get
		{
			return _italicChuKy;
		}
		set
		{
			if (SetProperty(ref _italicChuKy, value, "ItalicChuKy"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeThanBai, (float)Math.Round(value), "SizeThanBai"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeNoiNhan, (float)Math.Round(value), "SizeNoiNhan"))
			{
				SaveCurrentTypographyStandard();
			}
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
			if (SetProperty(ref _sizeChuKy, (float)Math.Round(value), "SizeChuKy"))
			{
				SaveCurrentTypographyStandard();
			}
		}
	}

	public string ResetTypographyStandardText => _selectedTypographyStandard switch
	{
		0 => "Đặt lại toàn bộ cỡ chữ chuẩn Nghị định 30", 
		1 => "Đặt lại toàn bộ cỡ chữ chuẩn Hướng dẫn 05", 
		2 => "Đặt lại toàn bộ cỡ chữ chuẩn Thông tư 68", 
		4 => "Đặt lại toàn bộ cỡ chữ chuẩn Hướng dẫn 29 (Đoàn)", 
		_ => "Đặt lại toàn bộ cỡ chữ Cá nhân về chuẩn NĐ 30", 
	};

	public string LabelQuocHieu => _selectedTypographyStandard switch
	{
		0 => "Quốc hiệu (NĐ 30: 12-13pt, chuẩn 12pt):", 
		1 => "Tiêu đề Đảng hiệu (HĐ 05: chuẩn 15pt):", 
		2 => "Quốc hiệu (TT 68: chuẩn 13pt):", 
		4 => "Tiêu đề 'ĐOÀN TNCS HỒ CHÍ MINH' (HD 29: chuẩn 15pt):", 
		_ => "Quốc hiệu (Tùy chỉnh pt):", 
	};

	public string LabelTieuNgu => _selectedTypographyStandard switch
	{
		0 => "Tiêu ngữ (NĐ 30: 13-14pt, chuẩn 13pt):", 
		1 => "Tiêu ngữ / Tiêu đề phụ (HĐ 05: 14pt):", 
		2 => "Tiêu ngữ (TT 68: chuẩn 13pt):", 
		4 => "Tiêu ngữ (HD 29: không dùng - văn bản Đoàn không có Tiêu ngữ):", 
		_ => "Tiêu ngữ (Tùy chỉnh pt):", 
	};

	public string LabelCoQuan => _selectedTypographyStandard switch
	{
		0 => "Cơ quan ban hành (NĐ 30: 12-13pt, chuẩn 13pt):", 
		1 => "Cơ quan ban hành (HĐ 05: 13-14pt, chuẩn 14pt):", 
		2 => "Cơ quan ban hành (TT 68: chuẩn 13pt):", 
		4 => "Cơ quan ban hành (HD 29: chuẩn 14pt, cấp trên đứng thường):", 
		_ => "Cơ quan ban hành (Tùy chỉnh pt):", 
	};

	public string LabelSoKyHieu => _selectedTypographyStandard switch
	{
		0 => "Số và ký hiệu (NĐ 30: 13pt):", 
		1 => "Số và ký hiệu (HĐ 05: 14pt):", 
		2 => "Số và ký hiệu (TT 68: chuẩn 13pt):", 
		4 => "Số và ký hiệu (HD 29: chuẩn 14pt, 'Số: 245-CV/TWĐTN-VP'):", 
		_ => "Số và ký hiệu (Tùy chỉnh pt):", 
	};

	public string LabelDiaDanh => _selectedTypographyStandard switch
	{
		0 => "Địa danh, ngày tháng (NĐ 30: 13-14pt, chuẩn 14pt):", 
		1 => "Địa danh, ngày tháng (HĐ 05: 13-14pt, chuẩn 14pt):", 
		2 => "Địa danh, ngày tháng (TT 68: chuẩn 13pt):", 
		4 => "Địa điểm, ngày tháng (HD 29: chuẩn 13pt, nghiêng):", 
		_ => "Địa danh, ngày tháng (Tùy chỉnh pt):", 
	};

	public string LabelTenLoai => _selectedTypographyStandard switch
	{
		0 => "Tên loại văn bản (NĐ 30: 13-14pt, chuẩn 14pt):", 
		1 => "Tên loại văn bản (HĐ 05: 15-16pt, chuẩn 15pt):", 
		2 => "Tên loại văn bản (TT 68: chuẩn 14pt):", 
		4 => "Tên loại văn bản (HD 29: chuẩn 16pt):", 
		_ => "Tên loại văn bản (Tùy chỉnh pt):", 
	};

	public string LabelTrichYeu => _selectedTypographyStandard switch
	{
		0 => "Trích yếu văn bản có tên loại (NĐ 30: 13-14pt, chuẩn 14pt):", 
		1 => "Trích yếu văn bản có tên loại (HĐ 05: 14-15pt, chuẩn 14pt):", 
		2 => "Trích yếu văn bản có tên loại (TT 68: chuẩn 14pt):", 
		4 => "Trích yếu văn bản có tên loại (HD 29: chuẩn 14pt, đậm):", 
		_ => "Trích yếu văn bản có tên loại (Tùy chỉnh pt):", 
	};

	public string LabelTrichYeuCongVan => _selectedTypographyStandard switch
	{
		0 => "Trích yếu công văn 'V/v...' (NĐ 30: 12-13pt, chuẩn 12pt):", 
		1 => "Trích yếu công văn 'V/v...' (HĐ 05: chuẩn 12pt, nghiêng):", 
		2 => "Trích yếu công văn 'V/v...' (TT 68: chuẩn 12pt):", 
		4 => "Trích yếu công văn 'V/v...' (HD 29: chuẩn 12pt, nghiêng):", 
		_ => "Trích yếu công văn 'V/v...' (Tùy chỉnh pt):", 
	};

	public string LabelThanBai => _selectedTypographyStandard switch
	{
		0 => "Thân bài nội dung (NĐ 30: 13-14pt, chuẩn 14pt):", 
		1 => "Thân bài nội dung (HĐ 05: 14-15pt, chuẩn 14pt):", 
		2 => "Thân bài nội dung (TT 68: chuẩn 14pt):", 
		4 => "Thân bài nội dung (HD 29: chuẩn 14pt):", 
		_ => "Thân bài nội dung (Tùy chỉnh pt):", 
	};

	public string LabelNoiNhan => _selectedTypographyStandard switch
	{
		0 => "Khối Nơi nhận (NĐ 30: 11-12pt, chuẩn 11pt):", 
		1 => "Khối Nơi nhận (HĐ 05: 12pt, tiêu đề 14pt):", 
		2 => "Khối Nơi nhận (TT 68: chuẩn 11pt, TĐ 12pt):", 
		4 => "Khối Nơi nhận (HD 29: chuẩn 11pt, 'Nơi nhận:' 13pt đậm):", 
		_ => "Khối Nơi nhận (Tùy chỉnh pt):", 
	};

	public string LabelChuKy => _selectedTypographyStandard switch
	{
		0 => "Khối Người ký (NĐ 30: 13-14pt, chuẩn 14pt):", 
		1 => "Khối Người ký (HĐ 05: chuẩn 14pt):", 
		2 => "Khối Người ký (TT 68: chuẩn 14pt):", 
		4 => "Khối Người ký (HD 29: chuẩn 14pt, chức vụ không đậm):", 
		_ => "Khối Người ký (Tùy chỉnh pt):", 
	};

	public string StatusMessage
	{
		get
		{
			return _statusMessage;
		}
		set
		{
			SetProperty(ref _statusMessage, value, "StatusMessage");
		}
	}

	public bool IsSaved
	{
		get
		{
			return _isSaved;
		}
		set
		{
			SetProperty(ref _isSaved, value, "IsSaved");
		}
	}

	public ICommand SaveCommand { get; private set; }

	public ICommand ResetCommand { get; private set; }

	public ICommand CloseCommand { get; private set; }

	public ICommand ResetMarginsToND30Command { get; private set; }

	public ICommand ResetTypographyToND30Command { get; private set; }

	public ICommand ResetCurrentMarginStandardCommand { get; private set; }

	public ICommand ResetCurrentTypographyStandardCommand { get; private set; }

	public ICommand SelectSectionCommand { get; private set; }

	public ICommand SelectAgencyStandardCommand { get; private set; }

	public ICommand SelectSignerStandardCommand { get; private set; }

	public ICommand SelectMarginStandardCommand { get; private set; }

	public ICommand SelectTypographyStandardCommand { get; private set; }

	private VndReadingOptions CurrentVndReadingOptions()
	{
		VndReadingOptions vndReadingOptions = new VndReadingOptions();
		vndReadingOptions.FourAsTu = _amountFourAsTu;
		vndReadingOptions.CapitalizeFirst = _amountCapitalizeFirst;
		vndReadingOptions.ZeroTensAsLe = _amountZeroTensAsLe;
		vndReadingOptions.ThousandAsNgan = _amountThousandAsNgan;
		return vndReadingOptions;
	}

	private void RaiseAmountPreviewChanged()
	{
		OnPropertyChanged("AmountPreviewText");
		OnPropertyChanged("TextToolsBadgeText");
	}

	public static string FormatIndentDisplay(float value, string originalText = null)
	{
		if (!string.IsNullOrEmpty(originalText) && originalText.Contains(","))
		{
			return value.ToString("0.##", CultureInfo.InvariantCulture).Replace('.', ',');
		}
		return value.ToString("0.##", CultureInfo.InvariantCulture);
	}

	public static bool TryParseFlexibleFloat(string input, out float result)
	{
		return JsonProfileRepository.TryParseFlexibleFloat(input, out result);
	}

	public void NormalizeIndentText()
	{
		_firstLineIndentCmText = FormatIndentDisplay(_firstLineIndentCm, _firstLineIndentCmText);
		OnPropertyChanged("FirstLineIndentCmText");
	}

	public void NormalizeSpacingBeforeText()
	{
		_spacingBeforePtText = FormatIndentDisplay(_spacingBeforePt, _spacingBeforePtText);
		OnPropertyChanged("SpacingBeforePtText");
	}

	public void NormalizeSpacingAfterText()
	{
		_spacingAfterPtText = FormatIndentDisplay(_spacingAfterPt, _spacingAfterPtText);
		OnPropertyChanged("SpacingAfterPtText");
	}

	public void NormalizeLineSpacingAtText()
	{
		_lineSpacingAtPtText = FormatIndentDisplay(_lineSpacingAtPt, _lineSpacingAtPtText);
		OnPropertyChanged("LineSpacingAtPtText");
	}

	private void NotifyTypographyLabels()
	{
		OnPropertyChanged("LabelQuocHieu");
		OnPropertyChanged("LabelTieuNgu");
		OnPropertyChanged("LabelCoQuan");
		OnPropertyChanged("LabelSoKyHieu");
		OnPropertyChanged("LabelDiaDanh");
		OnPropertyChanged("LabelTenLoai");
		OnPropertyChanged("LabelTrichYeu");
		OnPropertyChanged("LabelTrichYeuCongVan");
		OnPropertyChanged("LabelThanBai");
		OnPropertyChanged("LabelNoiNhan");
		OnPropertyChanged("LabelChuKy");
	}

	public PersonalizationViewModel(string customConfigPath = null)
	{
		SlotOptions = new ObservableCollection<string> { "Mẫu 1 (Mặc định)", "Mẫu 2", "Mẫu 3", "Mẫu 4", "Mẫu 5", "Mẫu 6", "Mẫu 7", "Mẫu 8" };
		_repo = new JsonProfileRepository(customConfigPath);
		LoadData();
		SaveCommand = new RelayCommand(ExecuteSave);
		ResetCommand = new RelayCommand(ExecuteReset);
		CloseCommand = new RelayCommand(ExecuteClose);
		ResetCurrentMarginStandardCommand = new RelayCommand(ExecuteResetCurrentMarginStandard);
		ResetCurrentTypographyStandardCommand = new RelayCommand(ExecuteResetCurrentTypographyStandard);
		ResetMarginsToND30Command = ResetCurrentMarginStandardCommand;
		ResetTypographyToND30Command = ResetCurrentTypographyStandardCommand;
		Action<object> execute = (object p) =>
		{
			if (p != null && int.TryParse(p.ToString(), out var result))
			{
				SelectedSectionIndex = result;
			}
		};
		SelectSectionCommand = new RelayCommand(execute);
		SelectAgencyStandardCommand = new RelayCommand((object p) =>
		{
			if (p != null && int.TryParse(p.ToString(), out var result))
			{
				SelectedAgencyStandard = result;
			}
		});
		SelectSignerStandardCommand = new RelayCommand((object p) =>
		{
			if (p != null && int.TryParse(p.ToString(), out var result))
			{
				SelectedSignerStandard = result;
			}
		});
		SelectMarginStandardCommand = new RelayCommand((object p) =>
		{
			if (p != null && int.TryParse(p.ToString(), out var result))
			{
				SelectedMarginStandard = result;
			}
		});
		SelectTypographyStandardCommand = new RelayCommand((object p) =>
		{
			if (p != null && int.TryParse(p.ToString(), out var result))
			{
				SelectedTypographyStandard = result;
			}
		});
	}

	private void LoadData()
	{
		try
		{
			_personalization = _repo.LoadPersonalization();
			_selectedAgencyStandard = _personalization.AgencyStandard;
			_selectedSignerStandard = _personalization.SignerStandard;
			_selectedMarginStandard = _personalization.MarginStandard;
			_selectedTypographyStandard = _personalization.TypographyStandard;
			UseCustomAgency = _personalization.UseCustomAgency;
			UseCustomSigner = _personalization.UseCustomSigner;
			UseCustomMargins = _personalization.UseCustomMargins;
			UseCustomParagraph = _personalization.UseCustomParagraph;
			UseCustomTypography = _personalization.UseCustomTypography;
			AmountFourAsTu = _personalization.AmountFourAsTu;
			AmountCapitalizeFirst = _personalization.AmountCapitalizeFirst;
			AmountZeroTensAsLe = _personalization.AmountZeroTensAsLe;
			AmountThousandAsNgan = _personalization.AmountThousandAsNgan;
			HeaderDateInTopCell = _personalization.HeaderDateInTopCell;
			LoadCurrentAgencySlot();
			LoadCurrentSignerSlot();
			LoadCurrentMarginStandard();
			LoadCurrentTypographyStandard();
			OnPropertyChanged("SelectedAgencyStandard");
			OnPropertyChanged("SelectedSignerStandard");
			OnPropertyChanged("SelectedMarginStandard");
			OnPropertyChanged("SelectedTypographyStandard");
			OnPropertyChanged("IsAgencyND30");
			OnPropertyChanged("IsAgencyHD05");
			OnPropertyChanged("IsAgencyTT68");
			OnPropertyChanged("IsAgencyCustom");
			OnPropertyChanged("IsAgencyHD29");
			OnPropertyChanged("IsSignerND30");
			OnPropertyChanged("IsSignerHD05");
			OnPropertyChanged("IsSignerTT68");
			OnPropertyChanged("IsSignerCustom");
			OnPropertyChanged("IsSignerHD29");
			OnPropertyChanged("IsMarginND30");
			OnPropertyChanged("IsMarginHD05");
			OnPropertyChanged("IsMarginTT68");
			OnPropertyChanged("IsMarginCustom");
			OnPropertyChanged("IsMarginHD29");
			OnPropertyChanged("IsTypographyND30");
			OnPropertyChanged("IsTypographyHD05");
			OnPropertyChanged("IsTypographyTT68");
			OnPropertyChanged("IsTypographyCustom");
			OnPropertyChanged("IsTypographyHD29");
			OnPropertyChanged("AgencyBadgeText");
			OnPropertyChanged("AgencyStatusDetail");
			OnPropertyChanged("SignerBadgeText");
			OnPropertyChanged("SignerStatusDetail");
			OnPropertyChanged("MarginBadgeText");
			OnPropertyChanged("MarginStatusDetail");
			OnPropertyChanged("TypographyBadgeText");
			OnPropertyChanged("TypographyStatusDetail");
			OnPropertyChanged("AgencyDefaultVisibility");
			OnPropertyChanged("SignerDefaultVisibility");
			OnPropertyChanged("MarginDefaultVisibility");
			OnPropertyChanged("TypographyDefaultVisibility");
			OnPropertyChanged("ResetMarginStandardText");
			OnPropertyChanged("ResetTypographyStandardText");
			NotifyTypographyLabels();
		}
		catch (Exception ex)
		{
			StatusMessage = "Lỗi khi nạp dữ liệu: " + ex.Message;
		}
	}

	private List<AgencyProfile> GetCurrentAgencyList()
	{
		if (_personalization == null)
		{
			return new List<AgencyProfile>();
		}
		return _personalization.GetAgenciesByStandard(_selectedAgencyStandard);
	}

	private void SaveCurrentAgencySlot()
	{
		if (_personalization != null)
		{
			List<AgencyProfile> currentAgencyList = GetCurrentAgencyList();
			EnsureAgencyCapacity(currentAgencyList, _agencySlotIndex, _selectedAgencyStandard);
			if (_agencySlotIndex >= 0 && _agencySlotIndex < currentAgencyList.Count)
			{
				AgencyProfile agencyProfile = currentAgencyList[_agencySlotIndex];
				agencyProfile.Alias = _agencyAlias ?? "";
				agencyProfile.ParentAgency = _agencyParent ?? "";
				agencyProfile.AgencyName = _agencyName ?? "";
				agencyProfile.Location = _agencyLoc ?? "";
				agencyProfile.CodePrefix = _agencyPrefix ?? "";
			}
		}
	}

	private void LoadCurrentAgencySlot()
	{
		if (_personalization != null)
		{
			List<AgencyProfile> currentAgencyList = GetCurrentAgencyList();
			EnsureAgencyCapacity(currentAgencyList, _agencySlotIndex, _selectedAgencyStandard);
			if (_agencySlotIndex >= 0 && _agencySlotIndex < currentAgencyList.Count)
			{
				AgencyProfile agencyProfile = currentAgencyList[_agencySlotIndex];
				AgencyAlias = agencyProfile.Alias;
				ParentAgency = agencyProfile.ParentAgency;
				AgencyName = agencyProfile.AgencyName;
				Location = agencyProfile.Location;
				CodePrefix = agencyProfile.CodePrefix;
			}
		}
	}

	private void EnsureAgencyCapacity(List<AgencyProfile> list, int index, int standard)
	{
		string text = standard switch
		{
			4 => "youth_", 
			2 => "mil_", 
			1 => "party_", 
			0 => "admin_", 
			_ => "custom_", 
		};
		RulePackType rulePack = RulePackTraits.FromStandardIndex(standard);
		while (list.Count <= index)
		{
			list.Add(new AgencyProfile(text + (list.Count + 1), "Mẫu " + (list.Count + 1), "", "", "", "", rulePack));
		}
	}

	private List<SignerProfile> GetCurrentSignerList()
	{
		if (_personalization == null)
		{
			return new List<SignerProfile>();
		}
		return _personalization.GetSignersByStandard(_selectedSignerStandard);
	}

	private void SaveCurrentSignerSlot()
	{
		if (_personalization != null)
		{
			List<SignerProfile> currentSignerList = GetCurrentSignerList();
			EnsureSignerCapacity(currentSignerList, _signerSlotIndex, _selectedSignerStandard);
			if (_signerSlotIndex >= 0 && _signerSlotIndex < currentSignerList.Count)
			{
				SignerProfile signerProfile = currentSignerList[_signerSlotIndex];
				signerProfile.AuthorityTitle = _authorityTitle ?? "";
				signerProfile.Position = _position ?? "";
				signerProfile.MilitaryRank = _militaryRank ?? "";
				signerProfile.FullName = _fullName ?? "";
				signerProfile.RecipientList = _recipientList ?? "";
				signerProfile.Storage = _storage ?? "";
			}
		}
	}

	private void LoadCurrentSignerSlot()
	{
		if (_personalization == null)
		{
			return;
		}
		List<SignerProfile> currentSignerList = GetCurrentSignerList();
		EnsureSignerCapacity(currentSignerList, _signerSlotIndex, _selectedSignerStandard);
		if (_signerSlotIndex >= 0 && _signerSlotIndex < currentSignerList.Count)
		{
			SignerProfile signerProfile = currentSignerList[_signerSlotIndex];
			AuthorityPrefix = signerProfile.AuthorityTitle;
			Position = signerProfile.Position;
			MilitaryRank = signerProfile.MilitaryRank;
			SignerName = signerProfile.FullName;
			if (string.IsNullOrEmpty(signerProfile.RecipientList))
			{
				RecipientProfile selectedRecipient = _personalization.GetSelectedRecipient();
				signerProfile.RecipientList = ((selectedRecipient != null && !string.IsNullOrEmpty(selectedRecipient.Content)) ? selectedRecipient.Content : "- Như Điều ...;\n- Lưu: VT, TH.");
			}
			if (string.IsNullOrEmpty(signerProfile.Storage))
			{
				signerProfile.Storage = "- Lưu: VT, TH.";
			}
			RecipientList = signerProfile.RecipientList;
			Storage = signerProfile.Storage;
		}
	}

	private void EnsureSignerCapacity(List<SignerProfile> list, int index, int standard)
	{
		string text = standard switch
		{
			4 => "youth_sig_", 
			2 => "mil_sig_", 
			1 => "party_sig_", 
			0 => "admin_sig_", 
			_ => "custom_sig_", 
		};
		RulePackType rulePack = RulePackTraits.FromStandardIndex(standard);
		while (list.Count <= index)
		{
			list.Add(new SignerProfile(text + (list.Count + 1), "Mẫu " + (list.Count + 1), "TM. BAN GIÁM ĐỐC", "GIÁM ĐỐC", "Nguyễn Văn " + (char)(65 + list.Count), "Thủ trưởng đơn vị", "", rulePack, "- Như Điều ...;\n- Lưu: VT, TH.", "- Lưu: VT, TH."));
		}
	}

	private void SaveCurrentMarginStandard()
	{
		if (_personalization != null)
		{
			MarginProfile marginsByStandard = _personalization.GetMarginsByStandard(_selectedMarginStandard);
			if (marginsByStandard != null)
			{
				marginsByStandard.MarginTopMm = _marginTopMm;
				marginsByStandard.MarginBottomMm = _marginBottomMm;
				marginsByStandard.MarginLeftMm = _marginLeftMm;
				marginsByStandard.MarginRightMm = _marginRightMm;
				marginsByStandard.FirstLineIndentCm = _firstLineIndentCm;
				marginsByStandard.SpacingBeforePt = _spacingBeforePt;
				marginsByStandard.SpacingAfterPt = _spacingAfterPt;
				marginsByStandard.LineSpacingRule = _lineSpacingRule ?? "1.15 lines";
				marginsByStandard.LineSpacingAtPt = _lineSpacingAtPt;
				marginsByStandard.NoSpaceBetweenParagraphsOfSameStyle = _noSpaceBetweenParagraphsOfSameStyle;
			}
		}
	}

	private void LoadCurrentMarginStandard()
	{
		if (_personalization == null)
		{
			return;
		}
		MarginProfile marginsByStandard = _personalization.GetMarginsByStandard(_selectedMarginStandard);
		if (marginsByStandard != null)
		{
			_marginTopMm = marginsByStandard.MarginTopMm;
			_marginBottomMm = marginsByStandard.MarginBottomMm;
			_marginLeftMm = marginsByStandard.MarginLeftMm;
			_marginRightMm = marginsByStandard.MarginRightMm;
			float num = marginsByStandard.FirstLineIndentCm;
			if (num >= 50f)
			{
				num /= 100f;
			}
			_firstLineIndentCm = num;
			marginsByStandard.FirstLineIndentCm = num;
			_firstLineIndentCmText = FormatIndentDisplay(_firstLineIndentCm, _firstLineIndentCmText);
			_spacingBeforePt = marginsByStandard.SpacingBeforePt;
			_spacingBeforePtText = FormatIndentDisplay(_spacingBeforePt, _spacingBeforePtText);
			_spacingAfterPt = marginsByStandard.SpacingAfterPt;
			_spacingAfterPtText = FormatIndentDisplay(_spacingAfterPt, _spacingAfterPtText);
			_lineSpacingRule = marginsByStandard.LineSpacingRule ?? "1.15 lines";
			_lineSpacingAtPt = marginsByStandard.LineSpacingAtPt;
			_lineSpacingAtPtText = FormatIndentDisplay(_lineSpacingAtPt, _lineSpacingAtPtText);
			_noSpaceBetweenParagraphsOfSameStyle = marginsByStandard.NoSpaceBetweenParagraphsOfSameStyle;
			OnPropertyChanged("MarginTopMm");
			OnPropertyChanged("MarginBottomMm");
			OnPropertyChanged("MarginLeftMm");
			OnPropertyChanged("MarginRightMm");
			OnPropertyChanged("FirstLineIndentCm");
			OnPropertyChanged("FirstLineIndentCmText");
			OnPropertyChanged("SpacingBeforePt");
			OnPropertyChanged("SpacingBeforePtText");
			OnPropertyChanged("SpacingAfterPt");
			OnPropertyChanged("SpacingAfterPtText");
			OnPropertyChanged("LineSpacingRule");
			OnPropertyChanged("LineSpacingAtPt");
			OnPropertyChanged("LineSpacingAtPtText");
			OnPropertyChanged("IsLineSpacingAtEnabled");
			OnPropertyChanged("NoSpaceBetweenParagraphsOfSameStyle");
			OnPropertyChanged("MarginStatusDetail");
		}
	}

	private void ExecuteResetCurrentMarginStandard()
	{
		if (_personalization != null)
		{
			switch (_selectedMarginStandard)
			{
			case 0:
				_personalization.MarginsAdmin = MarginProfile.CreateND30Default();
				StatusMessage = "Đã khôi phục thông số căn lề chuẩn Nghị định 30 về mặc định.";
				break;
			case 1:
				_personalization.MarginsParty = MarginProfile.CreateHD05Default();
				StatusMessage = "Đã khôi phục thông số căn lề chuẩn Hướng dẫn 05 về mặc định.";
				break;
			case 2:
				_personalization.MarginsMilitary = MarginProfile.CreateTT68Default();
				StatusMessage = "Đã khôi phục thông số căn lề chuẩn Thông tư 68 về mặc định.";
				break;
			case 4:
				_personalization.MarginsYouth = MarginProfile.CreateHD29Default();
				StatusMessage = "Đã khôi phục thông số căn lề chuẩn Hướng dẫn 29 (Đoàn) về mặc định.";
				break;
			default:
				_personalization.MarginsCustom = MarginProfile.CreateND30Default();
				StatusMessage = "Đã khôi phục thông số căn lề Cá nhân về chuẩn Nghị định 30.";
				break;
			}
			LoadCurrentMarginStandard();
		}
	}

	private void SaveCurrentTypographyStandard()
	{
		if (_personalization != null)
		{
			TypographyProfile typographyByStandard = _personalization.GetTypographyByStandard(_selectedTypographyStandard);
			if (typographyByStandard != null)
			{
				RulePackType rulePack = RulePackTraits.FromStandardIndex(_selectedTypographyStandard);
				typographyByStandard.SizeQuocHieu = _sizeQuocHieu;
				typographyByStandard.SizeTieuNgu = _sizeTieuNgu;
				typographyByStandard.SizeCoQuan = _sizeCoQuan;
				typographyByStandard.SizeSoKyHieu = _sizeSoKyHieu;
				typographyByStandard.SizeDiaDanh = _sizeDiaDanh;
				typographyByStandard.SizeTenLoai = _sizeTenLoai;
				typographyByStandard.SizeTrichYeu = _sizeTrichYeu;
				typographyByStandard.SizeTrichYeuCongVan = _sizeTrichYeuCongVan;
				typographyByStandard.SizeThanBai = _sizeThanBai;
				typographyByStandard.SizeNoiNhan = _sizeNoiNhan;
				typographyByStandard.SizeChuKy = _sizeChuKy;
				typographyByStandard.BoldQuocHieu = _boldQuocHieu;
				typographyByStandard.ItalicQuocHieu = _italicQuocHieu;
				typographyByStandard.BoldTieuNgu = _boldTieuNgu;
				typographyByStandard.ItalicTieuNgu = _italicTieuNgu;
				typographyByStandard.BoldCoQuan = _boldCoQuan;
				typographyByStandard.ItalicCoQuan = _italicCoQuan;
				typographyByStandard.BoldSoKyHieu = _boldSoKyHieu;
				typographyByStandard.ItalicSoKyHieu = _italicSoKyHieu;
				typographyByStandard.BoldDiaDanh = _boldDiaDanh;
				typographyByStandard.ItalicDiaDanh = _italicDiaDanh;
				typographyByStandard.BoldTenLoai = _boldTenLoai;
				typographyByStandard.ItalicTenLoai = _italicTenLoai;
				typographyByStandard.BoldTrichYeu = _boldTrichYeu;
				typographyByStandard.ItalicTrichYeu = _italicTrichYeu;
				typographyByStandard.BoldTrichYeuCongVan = _boldTrichYeuCongVan;
				typographyByStandard.ItalicTrichYeuCongVan = _italicTrichYeuCongVan;
				typographyByStandard.BoldThanBai = _boldThanBai;
				typographyByStandard.ItalicThanBai = _italicThanBai;
				typographyByStandard.BoldNoiNhan = _boldNoiNhan;
				typographyByStandard.ItalicNoiNhan = _italicNoiNhan;
				typographyByStandard.BoldChuKy = _boldChuKy;
				typographyByStandard.ItalicChuKy = _italicChuKy;
				typographyByStandard.ValidateAndClamp(rulePack);
			}
		}
	}

	private void LoadCurrentTypographyStandard()
	{
		if (_personalization != null)
		{
			TypographyProfile typographyByStandard = _personalization.GetTypographyByStandard(_selectedTypographyStandard);
			if (typographyByStandard != null)
			{
				RulePackType rulePack = RulePackTraits.FromStandardIndex(_selectedTypographyStandard);
				typographyByStandard.ValidateAndClamp(rulePack);
				_sizeQuocHieu = typographyByStandard.SizeQuocHieu;
				_sizeTieuNgu = typographyByStandard.SizeTieuNgu;
				_sizeCoQuan = typographyByStandard.SizeCoQuan;
				_sizeSoKyHieu = typographyByStandard.SizeSoKyHieu;
				_sizeDiaDanh = typographyByStandard.SizeDiaDanh;
				_sizeTenLoai = typographyByStandard.SizeTenLoai;
				_sizeTrichYeu = typographyByStandard.SizeTrichYeu;
				_sizeTrichYeuCongVan = typographyByStandard.SizeTrichYeuCongVan;
				_sizeThanBai = typographyByStandard.SizeThanBai;
				_sizeNoiNhan = typographyByStandard.SizeNoiNhan;
				_sizeChuKy = typographyByStandard.SizeChuKy;
				_boldQuocHieu = typographyByStandard.BoldQuocHieu;
				_italicQuocHieu = typographyByStandard.ItalicQuocHieu;
				_boldTieuNgu = typographyByStandard.BoldTieuNgu;
				_italicTieuNgu = typographyByStandard.ItalicTieuNgu;
				_boldCoQuan = typographyByStandard.BoldCoQuan;
				_italicCoQuan = typographyByStandard.ItalicCoQuan;
				_boldSoKyHieu = typographyByStandard.BoldSoKyHieu;
				_italicSoKyHieu = typographyByStandard.ItalicSoKyHieu;
				_boldDiaDanh = typographyByStandard.BoldDiaDanh;
				_italicDiaDanh = typographyByStandard.ItalicDiaDanh;
				_boldTenLoai = typographyByStandard.BoldTenLoai;
				_italicTenLoai = typographyByStandard.ItalicTenLoai;
				_boldTrichYeu = typographyByStandard.BoldTrichYeu;
				_italicTrichYeu = typographyByStandard.ItalicTrichYeu;
				_boldTrichYeuCongVan = typographyByStandard.BoldTrichYeuCongVan;
				_italicTrichYeuCongVan = typographyByStandard.ItalicTrichYeuCongVan;
				_boldThanBai = typographyByStandard.BoldThanBai;
				_italicThanBai = typographyByStandard.ItalicThanBai;
				_boldNoiNhan = typographyByStandard.BoldNoiNhan;
				_italicNoiNhan = typographyByStandard.ItalicNoiNhan;
				_boldChuKy = typographyByStandard.BoldChuKy;
				_italicChuKy = typographyByStandard.ItalicChuKy;
				OnPropertyChanged("SizeQuocHieu");
				OnPropertyChanged("SizeTieuNgu");
				OnPropertyChanged("SizeCoQuan");
				OnPropertyChanged("SizeSoKyHieu");
				OnPropertyChanged("SizeDiaDanh");
				OnPropertyChanged("SizeTenLoai");
				OnPropertyChanged("SizeTrichYeu");
				OnPropertyChanged("SizeTrichYeuCongVan");
				OnPropertyChanged("SizeThanBai");
				OnPropertyChanged("SizeNoiNhan");
				OnPropertyChanged("SizeChuKy");
				OnPropertyChanged("BoldQuocHieu");
				OnPropertyChanged("ItalicQuocHieu");
				OnPropertyChanged("BoldTieuNgu");
				OnPropertyChanged("ItalicTieuNgu");
				OnPropertyChanged("BoldCoQuan");
				OnPropertyChanged("ItalicCoQuan");
				OnPropertyChanged("BoldSoKyHieu");
				OnPropertyChanged("ItalicSoKyHieu");
				OnPropertyChanged("BoldDiaDanh");
				OnPropertyChanged("ItalicDiaDanh");
				OnPropertyChanged("BoldTenLoai");
				OnPropertyChanged("ItalicTenLoai");
				OnPropertyChanged("BoldTrichYeu");
				OnPropertyChanged("ItalicTrichYeu");
				OnPropertyChanged("BoldTrichYeuCongVan");
				OnPropertyChanged("ItalicTrichYeuCongVan");
				OnPropertyChanged("BoldThanBai");
				OnPropertyChanged("ItalicThanBai");
				OnPropertyChanged("BoldNoiNhan");
				OnPropertyChanged("ItalicNoiNhan");
				OnPropertyChanged("BoldChuKy");
				OnPropertyChanged("ItalicChuKy");
				OnPropertyChanged("TypographyStatusDetail");
			}
		}
	}

	private void ExecuteResetCurrentTypographyStandard()
	{
		if (_personalization != null)
		{
			switch (_selectedTypographyStandard)
			{
			case 0:
				_personalization.TypographyAdmin = TypographyProfile.CreateND30Default();
				StatusMessage = "Đã khôi phục kích thước cỡ chữ chuẩn Nghị định 30 về mặc định.";
				break;
			case 1:
				_personalization.TypographyParty = TypographyProfile.CreateHD05Default();
				StatusMessage = "Đã khôi phục kích thước cỡ chữ chuẩn Hướng dẫn 05 về mặc định.";
				break;
			case 2:
				_personalization.TypographyMilitary = TypographyProfile.CreateTT68Default();
				StatusMessage = "Đã khôi phục kích thước cỡ chữ chuẩn Thông tư 68 về mặc định.";
				break;
			case 4:
				_personalization.TypographyYouth = TypographyProfile.CreateHD29Default();
				StatusMessage = "Đã khôi phục kích thước cỡ chữ chuẩn Hướng dẫn 29 (Đoàn) về mặc định.";
				break;
			default:
				_personalization.TypographyCustom = TypographyProfile.CreateND30Default();
				StatusMessage = "Đã khôi phục kích thước cỡ chữ Cá nhân về chuẩn Nghị định 30.";
				break;
			}
			LoadCurrentTypographyStandard();
		}
	}

	private void ExecuteSave()
	{
		try
		{
			SaveCurrentAgencySlot();
			SaveCurrentSignerSlot();
			SaveCurrentMarginStandard();
			SaveCurrentTypographyStandard();
			_personalization.AgencyStandard = _selectedAgencyStandard;
			_personalization.SignerStandard = _selectedSignerStandard;
			_personalization.MarginStandard = _selectedMarginStandard;
			_personalization.TypographyStandard = _selectedTypographyStandard;
			_personalization.UseCustomAgency = UseCustomAgency;
			_personalization.UseCustomSigner = UseCustomSigner;
			_personalization.UseCustomMargins = UseCustomMargins;
			_personalization.UseCustomParagraph = UseCustomParagraph;
			_personalization.UseCustomTypography = UseCustomTypography;
			_personalization.AmountFourAsTu = AmountFourAsTu;
			_personalization.AmountCapitalizeFirst = AmountCapitalizeFirst;
			_personalization.AmountZeroTensAsLe = AmountZeroTensAsLe;
			_personalization.AmountThousandAsNgan = AmountThousandAsNgan;
			_personalization.HeaderDateInTopCell = HeaderDateInTopCell;
			List<AgencyProfile> currentAgencyList = GetCurrentAgencyList();
			if (currentAgencyList.Count > _agencySlotIndex)
			{
				_personalization.SelectedAgencyId = currentAgencyList[_agencySlotIndex].Id;
			}
			List<SignerProfile> currentSignerList = GetCurrentSignerList();
			if (currentSignerList.Count > _signerSlotIndex)
			{
				_personalization.SelectedSignerId = currentSignerList[_signerSlotIndex].Id;
			}
			_repo.SavePersonalization(_personalization);
			IsSaved = true;
			StatusMessage = "Đã lưu thông tin cá nhân hóa thành công!";
		}
		catch (Exception ex)
		{
			IsSaved = false;
			StatusMessage = "Lỗi khi lưu: " + ex.Message;
		}
	}

	private void ExecuteReset()
	{
		_personalization = UserPersonalization.CreateDefault();
		LoadData();
		StatusMessage = "Đã khôi phục các giá trị mẫu mặc định của toàn bộ 4 quy chuẩn.";
	}

	private void ExecuteClose()
	{
		if (CloseAction != null)
		{
			CloseAction();
		}
	}
}
