using System;
using System.Globalization;
using System.Windows.Data;
using VFormat.Core.Serialization;

namespace VFormat.UI.Common;

[ValueConversion(typeof(float), typeof(string))]
public class FlexibleFloatConverter : IValueConverter
{
	private string _lastInputString;

	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		if (value is float num)
		{
			if (!string.IsNullOrEmpty(_lastInputString) && JsonProfileRepository.TryParseFlexibleFloat(_lastInputString, out var result) && Math.Abs(result - num) < 0.0001f)
			{
				return _lastInputString;
			}
			return num.ToString("0.##", CultureInfo.InvariantCulture);
		}
		if (value == null)
		{
			return "";
		}
		return value.ToString();
	}

	public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
	{
		string text = (_lastInputString = value as string);
		if (string.IsNullOrWhiteSpace(text))
		{
			return 0f;
		}
		if (JsonProfileRepository.TryParseFlexibleFloat(text, out var result))
		{
			if (parameter as string == "indent" && result >= 50f)
			{
				result /= 100f;
			}
			return result;
		}
		return Binding.DoNothing;
	}
}
