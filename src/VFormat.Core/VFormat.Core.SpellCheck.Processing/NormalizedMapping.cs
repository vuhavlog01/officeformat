namespace VFormat.Core.SpellCheck.Processing;

public class NormalizedMapping
{
	private readonly int[] _normToOrigStart;

	private readonly int[] _normToOrigEnd;

	public string OriginalText { get; private set; }

	public string NormalizedText { get; private set; }

	public NormalizedMapping(string originalText, string normalizedText, int[] normToOrigStart, int[] normToOrigEnd)
	{
		OriginalText = originalText ?? string.Empty;
		NormalizedText = normalizedText ?? string.Empty;
		_normToOrigStart = normToOrigStart;
		_normToOrigEnd = normToOrigEnd;
	}

	public int ToOriginalStart(int normIndex)
	{
		if (_normToOrigStart == null || _normToOrigStart.Length == 0)
		{
			return normIndex;
		}
		if (normIndex < 0)
		{
			return 0;
		}
		if (normIndex >= _normToOrigStart.Length)
		{
			return OriginalText.Length;
		}
		return _normToOrigStart[normIndex];
	}

	public int ToOriginalEnd(int normIndex)
	{
		if (_normToOrigEnd == null || _normToOrigEnd.Length == 0)
		{
			return normIndex;
		}
		if (normIndex <= 0)
		{
			return 0;
		}
		if (normIndex >= _normToOrigEnd.Length)
		{
			return OriginalText.Length;
		}
		return _normToOrigEnd[normIndex];
	}
}
