using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using VFormat.Core.Common;

namespace VFormat.Core.Templates;

public class UserTemplateService
{
	private static readonly Lazy<UserTemplateService> _instance = new Lazy<UserTemplateService>(() => new UserTemplateService());

	private readonly object _lock = new object();

	private readonly JavaScriptSerializer _serializer;

	private readonly string _storageRoot;

	public static UserTemplateService Instance => _instance.Value;

	public string TemplateStorageDirectory => Path.Combine(_storageRoot ?? Constants.AppDataDirectory, "user_templates");

	public string MetadataFilePath => Path.Combine(TemplateStorageDirectory, "user_templates.json");

	public UserTemplateService(string storageRoot = null)
	{
		_storageRoot = storageRoot;
		_serializer = new JavaScriptSerializer
		{
			MaxJsonLength = 10485760
		};
		EnsureDirectoryExists();
	}

	private void EnsureDirectoryExists()
	{
		try
		{
			if (!Directory.Exists(TemplateStorageDirectory))
			{
				Directory.CreateDirectory(TemplateStorageDirectory);
			}
		}
		catch
		{
		}
	}

	public List<UserTemplateItem> GetAllTemplates()
	{
		lock (_lock)
		{
			EnsureDirectoryExists();
			if (!File.Exists(MetadataFilePath) && File.Exists(MetadataFilePath + ".bak"))
			{
				return ReadMetadata(MetadataFilePath + ".bak");
			}
			if (!File.Exists(MetadataFilePath))
			{
				EnsureSampleTemplatesExist();
			}
			if (!File.Exists(MetadataFilePath))
			{
				return new List<UserTemplateItem>();
			}
			try
			{
				return ReadMetadata(MetadataFilePath);
			}
			catch (Exception innerException)
			{
				string path = MetadataFilePath + ".bak";
				if (!File.Exists(path))
				{
					throw new IOException("Danh sách biểu mẫu bị lỗi. Giữ nguyên dữ liệu để khôi phục từ bản sao.", innerException);
				}
				return ReadMetadata(path);
			}
		}
	}

	public UserTemplateItem GetTemplateById(string templateId)
	{
		if (string.IsNullOrEmpty(templateId))
		{
			return null;
		}
		List<UserTemplateItem> allTemplates = GetAllTemplates();
		return allTemplates.Find((UserTemplateItem x) => string.Equals(x.Id, templateId, StringComparison.OrdinalIgnoreCase));
	}

	public string SuggestTemplateName(string filePath)
	{
		if (string.IsNullOrWhiteSpace(filePath))
		{
			return "Biểu mẫu mới";
		}
		try
		{
			string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
			if (string.IsNullOrWhiteSpace(fileNameWithoutExtension))
			{
				return "Biểu mẫu mới";
			}
			string input = fileNameWithoutExtension.Replace('_', ' ').Replace('-', ' ');
			input = Regex.Replace(input, "\\s+", " ").Trim();
			return string.IsNullOrEmpty(input) ? "Biểu mẫu mới" : input;
		}
		catch
		{
			return "Biểu mẫu mới";
		}
	}

	public string GetTemplateFullPath(UserTemplateItem item)
	{
		if (item == null || string.IsNullOrEmpty(item.FileName))
		{
			return null;
		}
		if (item.FileName != Path.GetFileName(item.FileName) || item.FileName.IndexOf(':') >= 0)
		{
			throw new IOException("Tên tệp biểu mẫu không hợp lệ.");
		}
		return Path.Combine(TemplateStorageDirectory, item.FileName);
	}

	public UserTemplateItem ImportTemplate(string sourceFilePath, string displayName, string description, out string errorMessage)
	{
		errorMessage = null;
		if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
		{
			errorMessage = "Tệp nguồn không tồn tại hoặc đường dẫn không hợp lệ.";
			return null;
		}
		string text = Path.GetExtension(sourceFilePath).ToLowerInvariant();
		if (!(text != ".docx") || !(text != ".doc") || !(text != ".dotx") || !(text != ".dot"))
		{
			lock (_lock)
			{
				EnsureDirectoryExists();
				try
				{
					string name = (string.IsNullOrWhiteSpace(displayName) ? SuggestTemplateName(sourceFilePath) : displayName.Trim());
					string text2 = SanitizeFileName(name);
					if (string.IsNullOrWhiteSpace(text2))
					{
						text2 = "UserTemplate_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
					}
					string text3 = text2 + text;
					string text4 = Path.Combine(TemplateStorageDirectory, text3);
					int num = 1;
					while (File.Exists(text4))
					{
						text3 = $"{text2} ({num}){text}";
						text4 = Path.Combine(TemplateStorageDirectory, text3);
						num++;
					}
					List<UserTemplateItem> allTemplates = GetAllTemplates();
					File.Copy(sourceFilePath, text4, overwrite: true);
					FileInfo fileInfo = new FileInfo(text4);
					UserTemplateItem userTemplateItem = new UserTemplateItem();
					userTemplateItem.Id = Guid.NewGuid().ToString("N");
					userTemplateItem.Name = name;
					userTemplateItem.FileName = text3;
					userTemplateItem.OriginalFilePath = sourceFilePath;
					userTemplateItem.DateAdded = DateTime.Now;
					userTemplateItem.LastModified = DateTime.Now;
					userTemplateItem.FileSizeBytes = fileInfo.Length;
					userTemplateItem.Description = description ?? "";
					UserTemplateItem userTemplateItem2 = userTemplateItem;
					allTemplates.Insert(0, userTemplateItem2);
					try
					{
						SaveMetadataInternal(allTemplates);
					}
					catch
					{
						try
						{
							File.Delete(text4);
						}
						catch
						{
						}
						throw;
					}
					return userTemplateItem2;
				}
				catch (Exception ex)
				{
					errorMessage = "Lỗi khi nhập biểu mẫu: " + ex.Message;
					return null;
				}
			}
		}
		errorMessage = "Định dạng tệp không được hỗ trợ. Vui lòng chọn tệp Word (.docx, .doc, .dotx).";
		return null;
	}

	public bool ReplaceTemplateFile(string templateId, string newSourceFilePath, out string errorMessage)
	{
		errorMessage = null;
		if (string.IsNullOrWhiteSpace(newSourceFilePath) || !File.Exists(newSourceFilePath))
		{
			errorMessage = "Tệp thay thế không tồn tại.";
			return false;
		}
		lock (_lock)
		{
			try
			{
				List<UserTemplateItem> allTemplates = GetAllTemplates();
				UserTemplateItem userTemplateItem = allTemplates.Find((UserTemplateItem x) => string.Equals(x.Id, templateId, StringComparison.OrdinalIgnoreCase));
				if (userTemplateItem == null)
				{
					errorMessage = "Không tìm thấy biểu mẫu cần thay thế.";
					return false;
				}
				string text = Path.GetExtension(newSourceFilePath).ToLowerInvariant();
				if (text != ".docx" && text != ".doc" && text != ".dotx" && text != ".dot")
				{
					throw new IOException("Chỉ hỗ trợ tệp Word .docx, .doc, .dotx, .dot.");
				}
				string text2 = "template_" + Guid.NewGuid().ToString("N") + text;
				string text3 = Path.Combine(TemplateStorageDirectory, text2);
				File.Copy(newSourceFilePath, text3, overwrite: false);
				userTemplateItem.FileName = text2;
				userTemplateItem.OriginalFilePath = newSourceFilePath;
				userTemplateItem.LastModified = DateTime.Now;
				userTemplateItem.FileSizeBytes = new FileInfo(text3).Length;
				try
				{
					SaveMetadataInternal(allTemplates);
				}
				catch
				{
					try
					{
						File.Delete(text3);
					}
					catch
					{
					}
					throw;
				}
				return true;
			}
			catch (Exception ex)
			{
				errorMessage = "Lỗi khi thay thế tệp biểu mẫu: " + ex.Message;
				return false;
			}
		}
	}

	public bool RenameTemplate(string templateId, string newDisplayName, out string errorMessage)
	{
		errorMessage = null;
		if (string.IsNullOrWhiteSpace(newDisplayName))
		{
			errorMessage = "Tên biểu mẫu không được để trống.";
			return false;
		}
		lock (_lock)
		{
			try
			{
				List<UserTemplateItem> allTemplates = GetAllTemplates();
				UserTemplateItem userTemplateItem = allTemplates.Find((UserTemplateItem x) => string.Equals(x.Id, templateId, StringComparison.OrdinalIgnoreCase));
				if (userTemplateItem == null)
				{
					errorMessage = "Không tìm thấy biểu mẫu cần đổi tên.";
					return false;
				}
				userTemplateItem.Name = newDisplayName.Trim();
				userTemplateItem.LastModified = DateTime.Now;
				SaveMetadataInternal(allTemplates);
				return true;
			}
			catch (Exception ex)
			{
				errorMessage = "Lỗi khi đổi tên biểu mẫu: " + ex.Message;
				return false;
			}
		}
	}

	public bool DeleteTemplate(string templateId, out string errorMessage)
	{
		errorMessage = null;
		lock (_lock)
		{
			try
			{
				List<UserTemplateItem> allTemplates = GetAllTemplates();
				UserTemplateItem userTemplateItem = allTemplates.Find((UserTemplateItem x) => string.Equals(x.Id, templateId, StringComparison.OrdinalIgnoreCase));
				if (userTemplateItem == null)
				{
					errorMessage = "Không tìm thấy biểu mẫu cần xóa.";
					return false;
				}
				allTemplates.Remove(userTemplateItem);
				SaveMetadataInternal(allTemplates);
				return true;
			}
			catch (Exception ex)
			{
				errorMessage = "Lỗi khi xóa biểu mẫu: " + ex.Message;
				return false;
			}
		}
	}

	private List<UserTemplateItem> ReadMetadata(string path)
	{
		string text = File.ReadAllText(path, Encoding.UTF8);
		if (text.IndexOf("/Date(", StringComparison.Ordinal) >= 0 && text.IndexOf("\\/Date(", StringComparison.Ordinal) < 0)
		{
			text = text.Replace("/Date(", "\\/Date(").Replace(")/", ")\\/");
		}
		List<UserTemplateItem> list = _serializer.Deserialize<List<UserTemplateItem>>(text);
		if (list == null)
		{
			throw new IOException("Danh sách biểu mẫu không hợp lệ.");
		}
		foreach (UserTemplateItem item in list)
		{
			if (item == null || string.IsNullOrWhiteSpace(item.FileName) || item.FileName != Path.GetFileName(item.FileName) || item.FileName.IndexOf(':') >= 0)
			{
				throw new IOException("Tên tệp biểu mẫu không hợp lệ.");
			}
		}
		return list;
	}

	private void SaveMetadataInternal(List<UserTemplateItem> items)
	{
		Directory.CreateDirectory(TemplateStorageDirectory);
		string text = MetadataFilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
		try
		{
			byte[] bytes = Encoding.UTF8.GetBytes(_serializer.Serialize(items));
			using (FileStream fileStream = new FileStream(text, FileMode.CreateNew, FileAccess.Write, FileShare.None))
			{
				fileStream.Write(bytes, 0, bytes.Length);
				fileStream.Flush(flushToDisk: true);
			}
			if (File.Exists(MetadataFilePath))
			{
				bool flag = false;
				try
				{
					ReadMetadata(MetadataFilePath);
					flag = true;
				}
				catch
				{
				}
				File.Replace(text, MetadataFilePath, flag ? (MetadataFilePath + ".bak") : null);
			}
			else
			{
				File.Move(text, MetadataFilePath);
			}
		}
		finally
		{
			if (File.Exists(text))
			{
				File.Delete(text);
			}
		}
	}

	private static string SanitizeFileName(string name)
	{
		if (string.IsNullOrWhiteSpace(name))
		{
			return "";
		}
		char[] invalidFileNameChars = Path.GetInvalidFileNameChars();
		StringBuilder stringBuilder = new StringBuilder();
		foreach (char c in name)
		{
			if (Array.IndexOf(invalidFileNameChars, c) < 0 && c != '/' && c != '\\' && c != ':' && c != '*' && c != '?' && c != '"' && c != '<' && c != '>' && c != '|')
			{
				stringBuilder.Append(c);
			}
			else
			{
				stringBuilder.Append('_');
			}
		}
		return stringBuilder.ToString().Trim();
	}

	public string EnsureValidTemplateFile(UserTemplateItem item)
	{
		return GetTemplateFullPath(item);
	}

	public void EnsureSampleTemplatesExist()
	{
		try
		{
			EnsureDirectoryExists();
			string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			string[] array = new string[4]
			{
				Path.Combine(directoryName ?? "", "templates", "user_samples", "user_templates.json"),
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates", "user_samples", "user_templates.json"),
				Path.Combine("D:\\Tool\\V-Format Addin\\templates\\user_samples", "user_templates.json"),
				Path.Combine(Constants.AppDataDirectory, "templates", "user_samples", "user_templates.json")
			};
			string text = null;
			string[] array2 = array;
			foreach (string text2 in array2)
			{
				if (File.Exists(text2))
				{
					text = text2;
					break;
				}
			}
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			string directoryName2 = Path.GetDirectoryName(text);
			if (Directory.Exists(directoryName2))
			{
				string[] files = Directory.GetFiles(directoryName2, "*.docx");
				string[] array3 = files;
				foreach (string text3 in array3)
				{
					string text4 = Path.Combine(TemplateStorageDirectory, Path.GetFileName(text3));
					if (!File.Exists(text4))
					{
						File.Copy(text3, text4, overwrite: true);
					}
				}
			}
			if (!File.Exists(MetadataFilePath))
			{
				File.Copy(text, MetadataFilePath, overwrite: true);
			}
		}
		catch
		{
		}
	}
}
