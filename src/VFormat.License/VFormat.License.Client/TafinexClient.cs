using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using VFormat.License.Models;

namespace VFormat.License.Client;

public class TafinexClient
{
	public const string DefaultBaseUrl = "";

	public const string BuyUrl = "";

	public const string TransferPurchaseUrl = "";

	public const string OfflinePortalUrl = "";

	public const string UpdateVersionUrl = "";

	public const string UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OfficeFormat/1.0";

	private readonly string _baseUrl;

	private readonly JavaScriptSerializer _serializer;

	public TafinexClient(string baseUrl = null)
	{
		_baseUrl = ((!string.IsNullOrEmpty(baseUrl)) ? baseUrl : "");
		_serializer = new JavaScriptSerializer();
		try
		{
			ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
		}
		catch
		{
		}
	}

	public LicenseCheckResult CheckLicense(string hwid, string licenseKey, string machineName = null)
	{
		string url = _baseUrl + "/check.php";
		Dictionary<string, object> dictionary = new Dictionary<string, object>();
		dictionary.Add("hwid", hwid);
		dictionary.Add("license_key", licenseKey ?? "");
		dictionary.Add("machine_name", machineName ?? Environment.MachineName);
		dictionary.Add("product_code", "officeformat");
		Dictionary<string, object> payload = dictionary;
		return SendLicenseRequest(url, payload, hwid, licenseKey);
	}

	public LicenseCheckResult ActivateLicense(string hwid, string licenseKey, string machineName = null)
	{
		string url = _baseUrl + "/activate.php";
		Dictionary<string, object> dictionary = new Dictionary<string, object>();
		dictionary.Add("hwid", hwid);
		dictionary.Add("license_key", licenseKey);
		dictionary.Add("machine_name", machineName ?? Environment.MachineName);
		dictionary.Add("product_code", "officeformat");
		Dictionary<string, object> payload = dictionary;
		return SendLicenseRequest(url, payload, hwid, licenseKey);
	}

	public Dictionary<string, object> CheckVersion(string customVersionUrl = null)
	{
		string requestUriString = ((!string.IsNullOrEmpty(customVersionUrl)) ? customVersionUrl : UpdateVersionUrl);
		if (string.IsNullOrWhiteSpace(requestUriString))
		{
			return null;
		}
		try
		{
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(requestUriString);
			httpWebRequest.Method = "GET";
			httpWebRequest.UserAgent = UserAgent;
			httpWebRequest.Timeout = 6000;
			httpWebRequest.ReadWriteTimeout = 6000;
			using HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			using StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8);
			string input = streamReader.ReadToEnd();
			return _serializer.Deserialize<Dictionary<string, object>>(input);
		}
		catch
		{
			return null;
		}
	}

	public DeviceTransferStatus GetTransferStatus(string hwid, string licenseKey, string pendingRequestId)
	{
		Dictionary<string, object> dictionary = NewTransferPayload("status", hwid, licenseKey);
		if (!string.IsNullOrEmpty(pendingRequestId))
		{
			dictionary["request_id"] = pendingRequestId;
		}
		return ToTransferStatus(PostSignedJson(_baseUrl + "/transfer.php", dictionary, 8000), hwid, pendingRequestId);
	}

	public DeviceTransferStatus ExecuteTransfer(string hwid, string licenseKey, string requestId)
	{
		Dictionary<string, object> dictionary = NewTransferPayload("execute", hwid, licenseKey);
		dictionary["request_id"] = requestId;
		return ToTransferStatus(PostSignedJson(_baseUrl + "/transfer.php", dictionary, 15000), hwid, requestId);
	}

	public string GetTransferPurchaseLink(string hwid, string licenseKey)
	{
		if (string.IsNullOrWhiteSpace(_baseUrl) || string.IsNullOrWhiteSpace(TransferPurchaseUrl))
		{
			return null;
		}
		Dictionary<string, object> dictionary = PostSignedJson(_baseUrl + "/transfer.php", NewTransferPayload("purchase_link", hwid, licenseKey), 8000);
		if (dictionary == null || GetString(dictionary, "status") != "PURCHASE_LINK" || GetString(dictionary, "hwid") != (hwid ?? ""))
		{
			return null;
		}
		string text = GetString(dictionary, "ticket");
		if (text.Length < 20 || text.Length > 64)
		{
			return null;
		}
		string text2 = text;
		foreach (char c in text2)
		{
			if ((c < 'A' || c > 'Z') && (c < 'a' || c > 'z') && (c < '0' || c > '9') && c != '-' && c != '_')
			{
				return null;
			}
		}
		return TransferPurchaseUrl + "?t=" + text;
	}

	private static Dictionary<string, object> NewTransferPayload(string action, string hwid, string licenseKey)
	{
		Dictionary<string, object> dictionary = new Dictionary<string, object>();
		dictionary.Add("action", action);
		dictionary.Add("hwid", hwid ?? "");
		dictionary.Add("license_key", licenseKey ?? "");
		dictionary.Add("machine_name", Environment.MachineName);
		dictionary.Add("product_code", "officeformat");
		return dictionary;
	}

	private static DeviceTransferStatus ToTransferStatus(Dictionary<string, object> dict, string hwid, string requestId)
	{
		if (dict == null)
		{
			DeviceTransferStatus deviceTransferStatus = new DeviceTransferStatus();
			deviceTransferStatus.State = DeviceTransferState.Unknown;
			deviceTransferStatus.IsNetworkError = true;
			return deviceTransferStatus;
		}
		if (GetString(dict, "hwid") != (hwid ?? "") || (!string.IsNullOrEmpty(requestId) && GetString(dict, "request_id") != requestId))
		{
			DeviceTransferStatus deviceTransferStatus2 = new DeviceTransferStatus();
			deviceTransferStatus2.State = DeviceTransferState.Unknown;
			deviceTransferStatus2.IsNetworkError = true;
			return deviceTransferStatus2;
		}
		DeviceTransferStatus deviceTransferStatus3 = DeviceTransferStatus.FromServerCode(GetString(dict, "transfer_code"));
		deviceTransferStatus3.CreditsKnown = GetInt(dict, "credits_known") == 1;
		deviceTransferStatus3.CreditsRemaining = Math.Max(0, GetInt(dict, "credits_remaining"));
		deviceTransferStatus3.CanBuyCredits = GetInt(dict, "can_buy") == 1;
		deviceTransferStatus3.AwaitingNewDevice = GetInt(dict, "awaiting_new_device") == 1;
		deviceTransferStatus3.RequestState = GetString(dict, "request_state");
		deviceTransferStatus3.Message = GetString(dict, "message");
		if (deviceTransferStatus3.State == DeviceTransferState.Unknown)
		{
			deviceTransferStatus3.IsNetworkError = true;
		}
		return deviceTransferStatus3;
	}

	private Dictionary<string, object> PostSignedJson(string url, Dictionary<string, object> payload, int timeoutMs)
	{
		if (string.IsNullOrWhiteSpace(_baseUrl) || string.IsNullOrWhiteSpace(url))
		{
			return null;
		}
		try
		{
			byte[] bytes = Encoding.UTF8.GetBytes(_serializer.Serialize(payload));
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
			httpWebRequest.Method = "POST";
			httpWebRequest.ContentType = "application/json; charset=utf-8";
			httpWebRequest.UserAgent = UserAgent;
			httpWebRequest.Timeout = timeoutMs;
			httpWebRequest.ReadWriteTimeout = timeoutMs;
			httpWebRequest.ContentLength = bytes.Length;
			using (Stream stream = httpWebRequest.GetRequestStream())
			{
				stream.Write(bytes, 0, bytes.Length);
			}
			using HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			using StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8);
			Dictionary<string, object> dictionary = _serializer.Deserialize<Dictionary<string, object>>(streamReader.ReadToEnd());
			if (dictionary == null || !dictionary.ContainsKey("rsa_signature") || dictionary["rsa_signature"] == null)
			{
				return null;
			}
			return VerifyTafinexSignature(dictionary, dictionary["rsa_signature"].ToString()) ? dictionary : null;
		}
		catch
		{
			return null;
		}
	}

	private static int GetInt(Dictionary<string, object> dict, string key)
	{
		if (dict.ContainsKey(key) && dict[key] is bool)
		{
			if (!(bool)dict[key])
			{
				return 0;
			}
			return 1;
		}
		if (!int.TryParse(GetString(dict, key), out var result))
		{
			return 0;
		}
		return result;
	}

	private LicenseCheckResult SendLicenseRequest(string url, Dictionary<string, object> payload, string hwid, string licenseKey)
	{
		LicenseCheckResult licenseCheckResult = new LicenseCheckResult();
		licenseCheckResult.Hwid = hwid;
		licenseCheckResult.LicenseKey = licenseKey;
		licenseCheckResult.VerifiedAt = DateTime.Now;
		LicenseCheckResult licenseCheckResult2 = licenseCheckResult;
		if (string.IsNullOrWhiteSpace(_baseUrl) || string.IsNullOrWhiteSpace(url))
		{
			licenseCheckResult2.Status = LicenseStatus.NoLicense;
			licenseCheckResult2.StatusCodeString = "NO_DOMAIN";
			licenseCheckResult2.Message = "Chưa cấu hình máy chủ kích hoạt.";
			return licenseCheckResult2;
		}
		try
		{
			string s = _serializer.Serialize(payload);
			byte[] bytes = Encoding.UTF8.GetBytes(s);
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
			httpWebRequest.Method = "POST";
			httpWebRequest.ContentType = "application/json; charset=utf-8";
			httpWebRequest.UserAgent = UserAgent;
			httpWebRequest.Timeout = 8000;
			httpWebRequest.ReadWriteTimeout = 8000;
			httpWebRequest.ContentLength = bytes.Length;
			using (Stream stream = httpWebRequest.GetRequestStream())
			{
				stream.Write(bytes, 0, bytes.Length);
			}
			using HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			using StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8);
			string input = streamReader.ReadToEnd();
			Dictionary<string, object> dict = _serializer.Deserialize<Dictionary<string, object>>(input);
			return ParseServerResponse(dict, hwid, licenseKey);
		}
		catch (WebException ex)
		{
			if (ex.Response != null)
			{
				try
				{
					using StreamReader streamReader2 = new StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8);
					string input2 = streamReader2.ReadToEnd();
					Dictionary<string, object> dict2 = _serializer.Deserialize<Dictionary<string, object>>(input2);
					return ParseServerResponse(dict2, hwid, licenseKey);
				}
				catch
				{
				}
			}
			licenseCheckResult2.Status = LicenseStatus.NoLicense;
			licenseCheckResult2.StatusCodeString = "NETWORK_ERROR";
			licenseCheckResult2.Message = "Không thể kết nối đến máy chủ. Vui lòng kiểm tra kết nối Internet.";
			return licenseCheckResult2;
		}
		catch (Exception ex2)
		{
			licenseCheckResult2.Status = LicenseStatus.NoLicense;
			licenseCheckResult2.StatusCodeString = "CLIENT_ERROR";
			licenseCheckResult2.Message = "Lỗi xử lý kết nối: " + ex2.Message;
			return licenseCheckResult2;
		}
	}

	private LicenseCheckResult ParseServerResponse(Dictionary<string, object> dict, string hwid, string licenseKey)
	{
		LicenseCheckResult licenseCheckResult = new LicenseCheckResult();
		licenseCheckResult.Hwid = hwid;
		licenseCheckResult.LicenseKey = licenseKey;
		licenseCheckResult.VerifiedAt = DateTime.Now;
		LicenseCheckResult licenseCheckResult2 = licenseCheckResult;
		if (dict == null)
		{
			licenseCheckResult2.Status = LicenseStatus.NoLicense;
			licenseCheckResult2.StatusCodeString = "EMPTY_RESPONSE";
			licenseCheckResult2.Message = "Máy chủ trả về phản hồi rỗng.";
			return licenseCheckResult2;
		}
		if (!dict.ContainsKey("rsa_signature") || dict["rsa_signature"] == null || string.IsNullOrWhiteSpace(dict["rsa_signature"].ToString()))
		{
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.StatusCodeString = "UNSIGNED_SERVER_RESPONSE";
			licenseCheckResult2.Message = "Máy chủ trả về dữ liệu chưa được ký số. Kích hoạt đã bị từ chối để bảo vệ bản quyền.";
			return licenseCheckResult2;
		}
		if (!VerifyTafinexSignature(dict, dict["rsa_signature"].ToString()))
		{
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.StatusCodeString = "TAMPERED_SERVER_RESPONSE";
			licenseCheckResult2.Message = "Phản hồi từ máy chủ không hợp lệ hoặc đã bị can thiệp.";
			return licenseCheckResult2;
		}
		string text = (licenseCheckResult2.StatusCodeString = GetString(dict, "status").ToUpperInvariant());
		licenseCheckResult2.CustomerName = GetString(dict, "customer_name");
		licenseCheckResult2.PlanCode = GetString(dict, "plan_code");
		if (string.IsNullOrEmpty(licenseCheckResult2.PlanCode))
		{
			licenseCheckResult2.PlanCode = GetString(dict, "type");
		}
		licenseCheckResult2.ExpiresAt = GetString(dict, "expires_at");
		licenseCheckResult2.Message = GetString(dict, "message");
		if (dict.ContainsKey("days_left") && int.TryParse(dict["days_left"].ToString(), out var result))
		{
			licenseCheckResult2.DaysLeft = result;
		}
		switch (text)
		{
		case "ACTIVE":
			licenseCheckResult2.Status = LicenseStatus.Active;
			licenseCheckResult2.IsSuccess = true;
			break;
		case "TRIAL_ACTIVE":
			licenseCheckResult2.Status = LicenseStatus.Active;
			licenseCheckResult2.IsSuccess = true;
			break;
		case "KEY_EXPIRED":
			licenseCheckResult2.Status = LicenseStatus.Expired;
			licenseCheckResult2.Message = ((!string.IsNullOrEmpty(licenseCheckResult2.Message)) ? licenseCheckResult2.Message : "Mã bản quyền đã hết hạn sử dụng.");
			break;
		case "KEY_LOCKED":
		case "KEY_REVOKED":
		case "KEY_BLOCKED":
		case "LICENSE_SUSPENDED":
			licenseCheckResult2.Status = LicenseStatus.Locked;
			licenseCheckResult2.Message = ((!string.IsNullOrEmpty(licenseCheckResult2.Message)) ? licenseCheckResult2.Message : "Mã bản quyền đã bị khóa bởi Quản trị viên!");
			break;
		case "MAX_DEVICES_EXCEEDED":
		case "MAX_DEVICES_REACHED":
			licenseCheckResult2.Status = LicenseStatus.MaxDevices;
			licenseCheckResult2.Message = ((!string.IsNullOrEmpty(licenseCheckResult2.Message)) ? licenseCheckResult2.Message : "Mã bản quyền đã đạt số lượng thiết bị tối đa cho phép.");
			break;
		case "DEVICE_BLOCKED":
		case "DEVICE_TRANSFERRED":
			licenseCheckResult2.Status = LicenseStatus.DeviceBlocked;
			licenseCheckResult2.Message = ((!string.IsNullOrEmpty(licenseCheckResult2.Message)) ? licenseCheckResult2.Message : "Bản quyền này đã được chuyển sang thiết bị mới. Thiết bị này không còn quyền sử dụng OfficeFormat!");
			break;
		case "KEY_INVALID":
		case "INVALID":
		case "KEY_NOT_FOUND":
			licenseCheckResult2.Status = LicenseStatus.Invalid;
			licenseCheckResult2.Message = ((!string.IsNullOrEmpty(licenseCheckResult2.Message)) ? licenseCheckResult2.Message : "Mã bản quyền không tồn tại hoặc không chính xác.");
			break;
		default:
			licenseCheckResult2.Status = LicenseStatus.NoLicense;
			break;
		}
		return licenseCheckResult2;
	}

	private static bool VerifyTafinexSignature(Dictionary<string, object> dict, string serverSig)
	{
		try
		{
			SortedDictionary<string, object> sortedDictionary = new SortedDictionary<string, object>(StringComparer.Ordinal);
			foreach (KeyValuePair<string, object> item in dict)
			{
				if (!item.Key.Equals("signature", StringComparison.OrdinalIgnoreCase) && !item.Key.Equals("rsa_signature", StringComparison.OrdinalIgnoreCase))
				{
					sortedDictionary[item.Key] = item.Value;
				}
			}
			JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
			string s = javaScriptSerializer.Serialize(sortedDictionary);
			byte[] signature = Convert.FromBase64String(serverSig);
			using RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
			rSACryptoServiceProvider.FromXmlString("<RSAKeyValue><Modulus>uVOdgqa2R6F9uISawTrZnQMBULMZB8Oi3Fx3DaPXTf5DORvfAVqCjquRaHzY6e9uwR3LU+Dc7GWYiwqSF8mV4iAY8p2StFADujkjPjHDKGa/ywQ91Dhrg5Hkk9E1fYmLxuUnmf4q47m1qbP09xg/EOOCDEEMSuQBArliY2Oxuhi5GaMNGkcH+O45v4sq6OcKst7CRQ4CZbP/yU69vlv8FegwD8iS35BSEReD9OG6t7oblXUscuOhapPf3Y9iXYz71zEGOKmti+MAEZJFlXz2XkMao56ZdpyjbnDwr/Bzo2h/hSOw/K/qWbmgdciAM2e4mP3E2rwf0ROGvTRpiay9dQ==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>");
			return rSACryptoServiceProvider.VerifyData(Encoding.UTF8.GetBytes(s), CryptoConfig.MapNameToOID("SHA256"), signature);
		}
		catch
		{
			return false;
		}
	}

	private static string GetString(Dictionary<string, object> dict, string key)
	{
		if (dict != null && dict.ContainsKey(key) && dict[key] != null)
		{
			return dict[key].ToString();
		}
		return string.Empty;
	}
}
