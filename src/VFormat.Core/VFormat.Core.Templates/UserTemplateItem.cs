using System;

namespace VFormat.Core.Templates;

[Serializable]
public class UserTemplateItem
{
	public string Id { get; set; }

	public string Name { get; set; }

	public string FileName { get; set; }

	public string OriginalFilePath { get; set; }

	public DateTime DateAdded { get; set; }

	public DateTime LastModified { get; set; }

	public long FileSizeBytes { get; set; }

	public string Description { get; set; }

	public string FormattedSize
	{
		get
		{
			if (FileSizeBytes < 1024)
			{
				return FileSizeBytes + " B";
			}
			if (FileSizeBytes < 1048576)
			{
				return ((double)FileSizeBytes / 1024.0).ToString("F1") + " KB";
			}
			return ((double)FileSizeBytes / 1048576.0).ToString("F2") + " MB";
		}
	}

	public string FormattedDate => LastModified.ToString("dd/MM/yyyy HH:mm");
}
