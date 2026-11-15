using System;
using System.Collections.Generic;
using VFormat.Core.Enums;

namespace VFormat.Core.Personalization;

public class RecipientProfile
{
	public string Id { get; set; }

	public string Title { get; set; }

	public string Content { get; set; }

	public RulePackType RulePack { get; set; }

	public RecipientProfile()
	{
		Id = Guid.NewGuid().ToString("N");
		Title = "Mẫu nơi nhận mới";
		Content = string.Empty;
		RulePack = RulePackType.ND30_2020;
	}

	public RecipientProfile(string id, string title, string content, RulePackType rulePack = RulePackType.ND30_2020)
	{
		Id = id;
		Title = title;
		Content = content;
		RulePack = rulePack;
	}

	public List<string> GetFormattedItems()
	{
		List<string> list = new List<string>();
		if (string.IsNullOrEmpty(Content))
		{
			return list;
		}
		string[] array = Content.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
		string[] array2 = array;
		foreach (string text in array2)
		{
			string text2 = text.Trim();
			if (!string.IsNullOrEmpty(text2))
			{
				if (!text2.StartsWith("-"))
				{
					text2 = "- " + text2;
				}
				list.Add(text2);
			}
		}
		return list;
	}

	public RecipientProfile Clone()
	{
		RecipientProfile recipientProfile = new RecipientProfile();
		recipientProfile.Id = Id;
		recipientProfile.Title = Title;
		recipientProfile.Content = Content;
		recipientProfile.RulePack = RulePack;
		return recipientProfile;
	}
}
