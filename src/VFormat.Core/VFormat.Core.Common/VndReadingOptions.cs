namespace VFormat.Core.Common;

public sealed class VndReadingOptions
{
	public bool FourAsTu { get; set; }

	public bool CapitalizeFirst { get; set; }

	public bool ZeroTensAsLe { get; set; }

	public bool ThousandAsNgan { get; set; }

	public static VndReadingOptions Default => new VndReadingOptions();

	public VndReadingOptions()
	{
		CapitalizeFirst = true;
	}
}
