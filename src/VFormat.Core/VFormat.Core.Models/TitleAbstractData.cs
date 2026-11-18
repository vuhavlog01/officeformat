namespace VFormat.Core.Models;

public class TitleAbstractData
{
	public string DocumentType { get; set; }

	public string SubjectText { get; set; }

	public bool IsDispatch { get; set; }

	public TitleAbstractData()
	{
		DocumentType = string.Empty;
		SubjectText = string.Empty;
		IsDispatch = false;
	}
}
