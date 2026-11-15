namespace VFormat.Core.Models;

public class HeaderData
{
	public string SuperiorAgency { get; set; }

	public string IssuingAgency { get; set; }

	public string DocNumber { get; set; }

	public string Location { get; set; }

	public string DateText { get; set; }

	public bool HasAsterisk { get; set; }

	public string SecurityIndicator { get; set; }

	public string UrgencyIndicator { get; set; }

	public string DispatchSubject { get; set; }

	public bool HasExtractedAgency { get; set; }

	public HeaderData()
	{
		SuperiorAgency = string.Empty;
		IssuingAgency = string.Empty;
		DocNumber = string.Empty;
		Location = string.Empty;
		DateText = string.Empty;
		HasAsterisk = false;
		SecurityIndicator = string.Empty;
		UrgencyIndicator = string.Empty;
		DispatchSubject = string.Empty;
		HasExtractedAgency = false;
	}

	public bool IsEmpty()
	{
		if (string.IsNullOrEmpty(SuperiorAgency) && string.IsNullOrEmpty(IssuingAgency) && string.IsNullOrEmpty(DocNumber) && string.IsNullOrEmpty(DateText))
		{
			return string.IsNullOrEmpty(DispatchSubject);
		}
		return false;
	}
}
