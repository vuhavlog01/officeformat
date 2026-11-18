using System;
using System.Windows.Media;
using VFormat.Core.SpellCheck;
using VFormat.Core.SpellCheck.Models;
using VFormat.WordEngine.SpellCheck;

namespace VFormat.UI.SpellCheck;

public class SpellIssueViewModel
{
	public WordSpellIssueContext Context { get; set; }

	public SpellCheckIssue Issue => Context.Issue;

	public string LocationDisplay => $"Đoạn {Context.ParagraphIndex}";

	public bool IsSpacingIssue => Issue?.RuleId != null && Issue.RuleId.StartsWith("SPACE_", StringComparison.Ordinal);

	public string OriginalDisplay => IsSpacingIssue ? SpacingCheckScanner.FormatVisibleSpaces(Issue.Original) : (Issue?.Original ?? string.Empty);

	public bool IsForeignTermNote
	{
		get
		{
			if (IsSpacingIssue)
			{
				return false;
			}
			if (Issue.Type == SpellCheckIssueType.Spelling)
			{
				return Issue.Severity == SpellCheckSeverity.Info;
			}
			return false;
		}
	}

	public string TagLabel
	{
		get
		{
			if (IsSpacingIssue)
			{
				return Issue.RuleId switch
				{
					"SPACE_MULTIPLE" => "Dấu cách kép",
					"SPACE_NBSP" => "Khoảng trắng cứng",
					"SPACE_BEFORE_PUNCT" => "Thừa cách trước dấu câu",
					"SPACE_MISSING_AFTER_PUNCT" => "Thiếu cách sau dấu câu",
					"SPACE_AFTER_OPEN_BRACKET" => "Thừa cách sau ngoặc/nháy",
					"SPACE_BEFORE_CLOSE_BRACKET" => "Thừa cách trước ngoặc/nháy",
					"SPACE_LEADING" => "Khoảng trắng đầu đoạn",
					"SPACE_TRAILING" => "Khoảng trắng cuối đoạn",
					"SPACE_TAB_AFTER_MARKER" => "Dấu Tab sau đề mục",
					"SPACE_STRAIGHT_QUOTE" => "Dấu nháy thẳng (\" ➔ “)",
					"SPACE_EN_DASH" => "Dấu gạch dài (– ➔ -)",
					_ => "Lỗi khoảng trắng",
				};
			}
			if (IsForeignTermNote)
			{
				return "Thuật ngữ nước ngoài";
			}
			if (Issue.Type == SpellCheckIssueType.Spelling && Issue.Severity == SpellCheckSeverity.Warning)
			{
				return "Từ lạ – cần xem lại";
			}
			return Issue.Type switch
			{
				SpellCheckIssueType.Spelling => "Lỗi chính tả", 
				SpellCheckIssueType.Capitalization => "Viết hoa", 
				SpellCheckIssueType.Abbreviation => "Viết tắt", 
				SpellCheckIssueType.Variant => "Biến thể", 
				SpellCheckIssueType.Repetition => "Lặp từ", 
				_ => "Cần xem lại", 
			};
		}
	}

	public Brush TagColor
	{
		get
		{
			if (IsForeignTermNote)
			{
				return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#64748B"));
			}
			return Issue.Type switch
			{
				SpellCheckIssueType.Spelling => (Brush)new SolidColorBrush((Color)ColorConverter.ConvertFromString((Issue.Severity == SpellCheckSeverity.Warning) ? "#B45309" : "#DC2626")), 
				SpellCheckIssueType.Capitalization => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D97706")), 
				SpellCheckIssueType.Abbreviation => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4F46E5")), 
				SpellCheckIssueType.Variant => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#16A34A")), 
				SpellCheckIssueType.Repetition => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0891B2")), 
				_ => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#475569")), 
			};
		}
	}
}
