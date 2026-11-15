using System;
using VFormat.Core.Enums;

namespace VFormat.WordEngine;

public class FormattingResult
{
	public bool Success { get; set; }

	public string Message { get; set; }

	public RulePackType RulePack { get; set; }

	public int ParagraphsProcessed { get; set; }

	public int TablesProcessed { get; set; }

	public int ComponentsFormatted { get; set; }

	public long ElapsedMilliseconds { get; set; }

	public Exception Error { get; set; }

	public bool SafeModeApplied { get; set; }

	public string HeaderStatus { get; set; }

	public string SignatureStatus { get; set; }

	public bool CanInsertTemplate { get; set; }

	public static FormattingResult NotApplied(RulePackType rulePack, string message, bool canInsertTemplate)
	{
		FormattingResult formattingResult = new FormattingResult();
		formattingResult.Success = false;
		formattingResult.RulePack = rulePack;
		formattingResult.Message = message;
		formattingResult.CanInsertTemplate = canInsertTemplate;
		return formattingResult;
	}

	public static FormattingResult Succeeded(RulePackType rulePack, int paragraphs, int tables, long elapsedMs)
	{
		FormattingResult formattingResult = new FormattingResult();
		formattingResult.Success = true;
		formattingResult.RulePack = rulePack;
		formattingResult.ParagraphsProcessed = paragraphs;
		formattingResult.TablesProcessed = tables;
		formattingResult.ElapsedMilliseconds = elapsedMs;
		formattingResult.Message = $"Chuẩn hóa thành công {paragraphs} đoạn văn, {tables} bảng ({elapsedMs} ms)";
		return formattingResult;
	}

	public static FormattingResult Failed(Exception ex)
	{
		FormattingResult formattingResult = new FormattingResult();
		formattingResult.Success = false;
		formattingResult.Error = ex;
		formattingResult.Message = ((ex != null) ? ex.Message : "Đã có lỗi xảy ra trong quá trình chuẩn hóa.");
		return formattingResult;
	}
}
