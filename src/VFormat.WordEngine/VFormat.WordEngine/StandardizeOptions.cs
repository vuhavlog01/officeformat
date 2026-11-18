namespace VFormat.WordEngine;

public sealed class StandardizeOptions
{
	public bool PageSetup { get; set; }

	public bool Typography { get; set; }

	public bool Tables { get; set; }

	public bool PageNumbers { get; set; }

	public static StandardizeOptions All => new StandardizeOptions();

	public bool IsEmpty
	{
		get
		{
			if (!PageSetup && !Typography && !Tables)
			{
				return !PageNumbers;
			}
			return false;
		}
	}

	public StandardizeOptions()
	{
		PageSetup = true;
		Typography = true;
		Tables = true;
		PageNumbers = true;
	}
}
