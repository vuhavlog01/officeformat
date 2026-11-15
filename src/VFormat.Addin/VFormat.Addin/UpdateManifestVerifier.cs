using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VFormat.Addin;

internal static class UpdateManifestVerifier
{
	public const string AllowedUpdateDomain = "";

	private const string PublicKeyXml = "<RSAKeyValue><Modulus>uVOdgqa2R6F9uISawTrZnQMBULMZB8Oi3Fx3DaPXTf5DORvfAVqCjquRaHzY6e9uwR3LU+Dc7GWYiwqSF8mV4iAY8p2StFADujkjPjHDKGa/ywQ91Dhrg5Hkk9E1fYmLxuUnmf4q47m1qbP09xg/EOOCDEEMSuQBArliY2Oxuhi5GaMNGkcH+O45v4sq6OcKst7CRQ4CZbP/yU69vlv8FegwD8iS35BSEReD9OG6t7oblXUscuOhapPf3Y9iXYz71zEGOKmti+MAEZJFlXz2XkMao56ZdpyjbnDwr/Bzo2h/hSOw/K/qWbmgdciAM2e4mP3E2rwf0ROGvTRpiay9dQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

	public static bool TryVerify(IDictionary<string, object> data, out UpdateManifest manifest, out string error)
	{
		manifest = null;
		error = null;
		try
		{
			if (data == null)
			{
				throw new Exception("Chưa cấu hình máy chủ cập nhật.");
			}
			string text = Get(data, "manifest_version");
			string text2 = Get(data, "version");
			string text3 = Get(data, "update_url");
			string text4 = Get(data, "sha256").ToUpperInvariant();
			string s = Get(data, "update_signature");
			if (text != "1" || !Version.TryParse(text2, out var result))
			{
				throw new Exception("Phiên bản manifest không hợp lệ.");
			}
			if (!Uri.TryCreate(text3, UriKind.Absolute, out var result2) || result2.Scheme != Uri.UriSchemeHttps || (!string.IsNullOrEmpty(AllowedUpdateDomain) && !result2.Host.Equals(AllowedUpdateDomain, StringComparison.OrdinalIgnoreCase)))
			{
				throw new Exception("Địa chỉ gói cập nhật không hợp lệ.");
			}
			if (!Regex.IsMatch(text4, "^[0-9A-F]{64}$"))
			{
				throw new Exception("Mã SHA-256 không hợp lệ.");
			}
			byte[] signature = Convert.FromBase64String(s);
			string s2 = "VFORMAT_UPDATE|1|" + text2 + "|" + text3 + "|" + text4;
			using (RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider())
			{
				rSACryptoServiceProvider.FromXmlString(PublicKeyXml);
				if (!rSACryptoServiceProvider.VerifyData(Encoding.UTF8.GetBytes(s2), CryptoConfig.MapNameToOID("SHA256"), signature))
				{
					throw new Exception("Chữ ký bản cập nhật không hợp lệ.");
				}
			}
			manifest = new UpdateManifest
			{
				Version = result,
				VersionText = text2,
				UpdateUrl = text3,
				Sha256 = text4
			};
			return true;
		}
		catch (Exception ex)
		{
			error = ex.Message;
			return false;
		}
	}

	private static string Get(IDictionary<string, object> data, string key)
	{
		if (data == null || !data.ContainsKey(key) || data[key] == null)
		{
			return string.Empty;
		}
		return Convert.ToString(data[key], CultureInfo.InvariantCulture).Trim();
	}
}
