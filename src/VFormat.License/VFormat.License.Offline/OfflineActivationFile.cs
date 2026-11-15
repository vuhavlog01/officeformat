using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using VFormat.License.Models;

namespace VFormat.License.Offline;

public sealed class OfflineActivationFile
{
	public string LicenseKey { get; private set; }

	public string OfflineCode { get; private set; }

	public static OfflineActivationFile Load(string path, string machineHwid)
	{
		string input;
		using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
		{
			if (fileStream.Length == 0 || fileStream.Length > 16384)
			{
				throw new InvalidDataException("Tệp kích hoạt trống hoặc vượt quá dung lượng cho phép.");
			}
			using StreamReader streamReader = new StreamReader(fileStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true), detectEncodingFromByteOrderMarks: true);
			input = streamReader.ReadToEnd();
		}
		JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
		javaScriptSerializer.MaxJsonLength = 16384;
		javaScriptSerializer.RecursionLimit = 8;
		JavaScriptSerializer javaScriptSerializer2 = javaScriptSerializer;
		Dictionary<string, object> dictionary = javaScriptSerializer2.Deserialize<Dictionary<string, object>>(input);
		if (dictionary == null || Read(dictionary, "format") != "vformat-offline" || !dictionary.ContainsKey("version") || !(dictionary["version"] is int) || (int)dictionary["version"] != 1)
		{
			throw new InvalidDataException("Không đúng định dạng tệp kích hoạt OfficeFormat hoặc phiên bản chưa được hỗ trợ.");
		}
		string licenseKey = Read(dictionary, "license_key").Trim().ToUpperInvariant();
		string text = Read(dictionary, "hwid").Trim();
		string text2 = Read(dictionary, "offline_code").Trim();
		if (string.IsNullOrEmpty(machineHwid) || string.IsNullOrEmpty(text) || Regex.Replace(text.ToUpperInvariant(), "[^A-Z0-9\\-]", "") != Regex.Replace(machineHwid.ToUpperInvariant(), "[^A-Z0-9\\-]", ""))
		{
			throw new InvalidDataException("Tệp kích hoạt được cấp cho máy khác. Vui lòng gửi đúng mã máy để cấp lại tệp.");
		}
		string[] array = text2.Split(new char[1] { '-' }, 4);
		if (array.Length != 4 || (array[3].Length != 512 && array[3].Length != 344))
		{
			throw new InvalidDataException("Tệp phải chứa chữ ký RSA hợp lệ. Vui lòng yêu cầu cấp lại tệp kích hoạt.");
		}
		LicenseCheckResult licenseCheckResult = OfflineLicenseValidator.VerifyOfflineCode(licenseKey, machineHwid, text2);
		if (!licenseCheckResult.IsSuccess)
		{
			throw new InvalidDataException(licenseCheckResult.Message);
		}
		OfflineActivationFile offlineActivationFile = new OfflineActivationFile();
		offlineActivationFile.LicenseKey = licenseKey;
		offlineActivationFile.OfflineCode = text2;
		return offlineActivationFile;
	}

	private static string Read(Dictionary<string, object> data, string key)
	{
		if (!data.TryGetValue(key, out var value) || !(value is string))
		{
			throw new InvalidDataException("Tệp kích hoạt thiếu thông tin bắt buộc: " + key);
		}
		return (string)value;
	}
}
