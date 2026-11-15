using System;
using System.Collections.Generic;

namespace VFormat.Core.Common;

/// <summary>
/// Hỗ trợ xử lý và phòng tránh lỗi mồ côi chữ (orphan word / runt) trong trình bày văn bản tiếng Việt.
/// Thay thế khoảng trắng trước từ/âm tiết cuối cùng của mỗi đoạn văn (hoặc dòng ngắt mềm)
/// bằng dấu cách không ngắt dòng (Non-Breaking Space '\u00A0' - tương đương Ctrl+Shift+Space trong Word).
/// </summary>
public static class OrphanWordHelper
{
	public const char NonBreakingSpace = '\u00a0';

	/// <summary>
	/// Tìm danh sách các vị trí khoảng trắng trước từ cuối cùng cần thay bằng Non-Breaking Space.
	/// Danh sách trả về được sắp xếp theo vị trí giảm dần (Start lớn trước) để khi thay thế
	/// trong Word Range không làm xê dịch offset của các vị trí đứng trước.
	/// </summary>
	public static List<Tuple<int, int>> GetOrphanSpaceReplacements(string text)
	{
		List<Tuple<int, int>> list = new List<Tuple<int, int>>();
		if (string.IsNullOrEmpty(text) || text.Length < 6)
		{
			return list;
		}

		int searchEnd = text.Length - 1;
		while (searchEnd >= 0 && (text[searchEnd] == '\r' || text[searchEnd] == '\n' || text[searchEnd] == '\a' || text[searchEnd] == ' ' || text[searchEnd] == '\t'))
		{
			searchEnd--;
		}
		if (searchEnd < 3)
		{
			return list;
		}

		List<int> breakPoints = new List<int>();
		for (int idx = 0; idx <= searchEnd; idx++)
		{
			if (text[idx] == '\v')
			{
				breakPoints.Add(idx);
			}
		}
		breakPoints.Add(searchEnd + 1);

		for (int b = breakPoints.Count - 1; b >= 0; b--)
		{
			int segEnd = breakPoints[b] - 1;
			while (segEnd >= 0 && (text[segEnd] == ' ' || text[segEnd] == '\t' || text[segEnd] == '\v' || text[segEnd] == '\r' || text[segEnd] == '\n'))
			{
				segEnd--;
			}
			if (segEnd < 3)
			{
				continue;
			}

			int lastSpace = -1;
			for (int i = segEnd; i >= 0; i--)
			{
				if (text[i] == '\v' || text[i] == '\r' || text[i] == '\n')
				{
					break;
				}
				if (text[i] == ' ' || text[i] == NonBreakingSpace)
				{
					lastSpace = i;
					break;
				}
			}
			if (lastSpace <= 0)
			{
				continue;
			}

			int prevWord = lastSpace - 1;
			while (prevWord >= 0 && (text[prevWord] == ' ' || text[prevWord] == '\t' || text[prevWord] == NonBreakingSpace))
			{
				prevWord--;
			}
			if (prevWord < 0 || text[prevWord] == '\v' || text[prevWord] == '\r' || text[prevWord] == '\n')
			{
				continue;
			}

			if (text[lastSpace] == NonBreakingSpace && (lastSpace == 0 || (text[lastSpace - 1] != ' ' && text[lastSpace - 1] != '\t')))
			{
				continue;
			}

			int spaceStart = prevWord + 1;
			int spaceLength = lastSpace - spaceStart + 1;
			list.Add(Tuple.Create(spaceStart, spaceLength));
		}

		return list;
	}

	/// <summary>
	/// Thay thế khoảng trắng trước từ cuối cùng của mỗi đoạn bằng Non-Breaking Space trên chuỗi thuần.
	/// </summary>
	public static string FixOrphanWordsInText(string text)
	{
		List<Tuple<int, int>> replacements = GetOrphanSpaceReplacements(text);
		if (replacements == null || replacements.Count == 0)
		{
			return text;
		}

		System.Text.StringBuilder sb = new System.Text.StringBuilder(text);
		foreach (Tuple<int, int> rep in replacements)
		{
			sb.Remove(rep.Item1, rep.Item2);
			sb.Insert(rep.Item1, NonBreakingSpace);
		}
		return sb.ToString();
	}
}
