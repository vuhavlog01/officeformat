using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using VFormat.License.Models;

namespace VFormat.License.Offline;

public static class OfflineLicenseValidator
{
	private const string RsaPublicKeyXml = "<RSAKeyValue><Modulus>uVOdgqa2R6F9uISawTrZnQMBULMZB8Oi3Fx3DaPXTf5DORvfAVqCjquRaHzY6e9uwR3LU+Dc7GWYiwqSF8mV4iAY8p2StFADujkjPjHDKGa/ywQ91Dhrg5Hkk9E1fYmLxuUnmf4q47m1qbP09xg/EOOCDEEMSuQBArliY2Oxuhi5GaMNGkcH+O45v4sq6OcKst7CRQ4CZbP/yU69vlv8FegwD8iS35BSEReD9OG6t7oblXUscuOhapPf3Y9iXYz71zEGOKmti+MAEZJFlXz2XkMao56ZdpyjbnDwr/Bzo2h/hSOw/K/qWbmgdciAM2e4mP3E2rwf0ROGvTRpiay9dQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

	private const string LegacySecretSalt = "vformat_sec_k98f23490fdsmf902348mksdf9823m490f";

	public static LicenseCheckResult VerifyOfflineCode(string licenseKey, string hwid, string offlineCode)
	{
		LicenseCheckResult licenseCheckResult = new LicenseCheckResult();
		licenseCheckResult.LicenseKey = licenseKey;
		licenseCheckResult.Hwid = hwid;
		licenseCheckResult.IsOffline = true;
		LicenseCheckResult licenseCheckResult2 = licenseCheckResult;
		if (string.IsNullOrEmpty(licenseKey) || string.IsNullOrEmpty(hwid) || string.IsNullOrEmpty(offlineCode))
		{
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.StatusCodeString = "EMPTY_PARAMETERS";
			licenseCheckResult2.Message = "Vui lòng cung cấp đầy đủ Mã bản quyền, Mã máy HWID và Mã kích hoạt Offline!";
			return licenseCheckResult2;
		}
		string text = Regex.Replace(licenseKey.ToUpperInvariant(), "[^A-Z0-9]", "");
		string text2 = Regex.Replace(hwid.ToUpperInvariant(), "[^A-Z0-9\\-]", "");
		string text3 = offlineCode.Trim();
		if (text.Length < 8 || text2.Length < 4)
		{
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.StatusCodeString = "INVALID_KEY_OR_HWID";
			licenseCheckResult2.Message = "Mã bản quyền hoặc Mã máy HWID không hợp lệ!";
			return licenseCheckResult2;
		}
		string[] array = text3.Split(new char[1] { '-' }, 4);
		if (array.Length != 4 || !string.Equals(array[0], "VFOFF", StringComparison.OrdinalIgnoreCase))
		{
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.StatusCodeString = "INVALID_OFFLINE_FORMAT";
			licenseCheckResult2.Message = "Định dạng mã kích hoạt Offline không chính xác! (Định dạng chuẩn: VFOFF-YYYYMMDD-PLAN-SIGNATURE)";
			return licenseCheckResult2;
		}
		string text4 = array[1].ToUpperInvariant();
		string text5 = array[2].ToUpperInvariant();
		string text6 = array[3].Trim();
		string payload = $"VFOFF|{text}|{text2}|{text4}|{text5}";
		bool flag = false;
		if (text6.Length == 512)
		{
			flag = VerifyRsaSignature(payload, text6);
		}
		else if (text6.Length >= 340 && text6.Length <= 350)
		{
			flag = VerifyRsaBase64Signature(payload, text6);
		}
		else if (text6.Length == 16)
		{
			flag = VerifyLegacyHmac(payload, text6);
		}
		if (!flag)
		{
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.StatusCodeString = "SIGNATURE_MISMATCH";
			licenseCheckResult2.Message = "Mã kích hoạt Offline không trùng khớp với Mã máy tính này hoặc sai Mã bản quyền!";
			return licenseCheckResult2;
		}
		if (text4 == "99991231")
		{
			licenseCheckResult2.Status = LicenseStatus.OfflineActive;
			licenseCheckResult2.StatusCodeString = "ACTIVE";
			licenseCheckResult2.PlanCode = "LIFETIME";
			licenseCheckResult2.ExpiresAt = "Vĩnh viễn";
			licenseCheckResult2.DaysLeft = 99999;
			licenseCheckResult2.IsSuccess = true;
			licenseCheckResult2.Message = "Bản quyền Offline Vĩnh viễn đã được xác thực thành công!";
			return licenseCheckResult2;
		}
		if (DateTime.TryParseExact(text4, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
		{
			if (DateTime.Today > result)
			{
				licenseCheckResult2.Status = LicenseStatus.Expired;
				licenseCheckResult2.StatusCodeString = "KEY_EXPIRED";
				licenseCheckResult2.PlanCode = text5;
				licenseCheckResult2.ExpiresAt = result.ToString("dd/MM/yyyy");
				licenseCheckResult2.DaysLeft = 0;
				licenseCheckResult2.Message = "Mã kích hoạt Offline đã hết hạn sử dụng vào ngày " + result.ToString("dd/MM/yyyy") + "!";
				return licenseCheckResult2;
			}
			int num = (int)(result - DateTime.Today).TotalDays + 1;
			licenseCheckResult2.Status = LicenseStatus.OfflineActive;
			licenseCheckResult2.StatusCodeString = "ACTIVE";
			licenseCheckResult2.PlanCode = text5;
			licenseCheckResult2.ExpiresAt = result.ToString("dd/MM/yyyy");
			licenseCheckResult2.DaysLeft = num;
			licenseCheckResult2.IsSuccess = true;
			licenseCheckResult2.Message = "Bản quyền Offline hợp lệ. Còn lại " + num + " ngày sử dụng.";
			return licenseCheckResult2;
		}
		licenseCheckResult2.Status = LicenseStatus.Invalid;
		licenseCheckResult2.StatusCodeString = "INVALID_EXPIRATION_DATE";
		licenseCheckResult2.Message = "Ngày hết hạn trong mã Offline không hợp lệ!";
		return licenseCheckResult2;
	}

	private static bool VerifyRsaSignature(string payload, string sigHex)
	{
		try
		{
			if (string.IsNullOrEmpty(sigHex) || sigHex.Length != 512)
			{
				return false;
			}
			byte[] array = new byte[sigHex.Length / 2];
			for (int i = 0; i < array.Length; i++)
			{
				array[i] = Convert.ToByte(sigHex.Substring(i * 2, 2), 16);
			}
			return VerifyRsaCore(payload, array);
		}
		catch
		{
			return false;
		}
	}

	private static bool VerifyRsaBase64Signature(string payload, string sigBase64)
	{
		try
		{
			if (string.IsNullOrEmpty(sigBase64))
			{
				return false;
			}
			byte[] sigBytes = Convert.FromBase64String(sigBase64);
			return VerifyRsaCore(payload, sigBytes);
		}
		catch
		{
			return false;
		}
	}

	private static bool VerifyRsaCore(string payload, byte[] sigBytes)
	{
		using RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
		rSACryptoServiceProvider.FromXmlString("<RSAKeyValue><Modulus>uVOdgqa2R6F9uISawTrZnQMBULMZB8Oi3Fx3DaPXTf5DORvfAVqCjquRaHzY6e9uwR3LU+Dc7GWYiwqSF8mV4iAY8p2StFADujkjPjHDKGa/ywQ91Dhrg5Hkk9E1fYmLxuUnmf4q47m1qbP09xg/EOOCDEEMSuQBArliY2Oxuhi5GaMNGkcH+O45v4sq6OcKst7CRQ4CZbP/yU69vlv8FegwD8iS35BSEReD9OG6t7oblXUscuOhapPf3Y9iXYz71zEGOKmti+MAEZJFlXz2XkMao56ZdpyjbnDwr/Bzo2h/hSOw/K/qWbmgdciAM2e4mP3E2rwf0ROGvTRpiay9dQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>");
		using SHA256 sHA = SHA256.Create();
		byte[] bytes = Encoding.UTF8.GetBytes(payload);
		byte[] rgbHash = sHA.ComputeHash(bytes);
		RSAPKCS1SignatureDeformatter rSAPKCS1SignatureDeformatter = new RSAPKCS1SignatureDeformatter(rSACryptoServiceProvider);
		rSAPKCS1SignatureDeformatter.SetHashAlgorithm("SHA256");
		return rSAPKCS1SignatureDeformatter.VerifySignature(rgbHash, sigBytes);
	}

	private static bool VerifyLegacyHmac(string payload, string sigHex)
	{
		try
		{
			using HMACSHA256 hMACSHA = new HMACSHA256(Encoding.UTF8.GetBytes("vformat_sec_k98f23490fdsmf902348mksdf9823m490f"));
			byte[] array = hMACSHA.ComputeHash(Encoding.UTF8.GetBytes(payload));
			string text = BitConverter.ToString(array).Replace("-", "").ToUpperInvariant();
			string b = text.Substring(0, 16);
			return string.Equals(sigHex, b, StringComparison.OrdinalIgnoreCase);
		}
		catch
		{
			return false;
		}
	}
}
