using System.Collections.Generic;

namespace VFormat.Core.Models;

public class SignatureData
{
	public List<string> Recipients { get; set; }

	public string Authority { get; set; }

	public string Title { get; set; }

	public string MilitaryRank { get; set; }

	public string SignerName { get; set; }

	public SignatureData()
	{
		Recipients = new List<string>();
		Authority = string.Empty;
		Title = string.Empty;
		MilitaryRank = string.Empty;
		SignerName = string.Empty;
	}

	public bool IsEmpty()
	{
		if (Recipients.Count == 0 && string.IsNullOrEmpty(Authority) && string.IsNullOrEmpty(Title))
		{
			return string.IsNullOrEmpty(SignerName);
		}
		return false;
	}
}
