using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using VFormat.Core.Common;

namespace VFormat.Core.Templates;

public class TemplateLibraryService
{
	private static readonly Lazy<TemplateLibraryService> _instance = new Lazy<TemplateLibraryService>(() => new TemplateLibraryService());

	private readonly object _lock = new object();

	private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer
	{
		MaxJsonLength = int.MaxValue
	};

	private List<TemplateItem> _allTemplates;

	private string _resolvedMetadataPath;

	public static TemplateLibraryService Instance => _instance.Value;

	public string MetadataFilePath
	{
		get
		{
			if (!string.IsNullOrEmpty(_resolvedMetadataPath) && File.Exists(_resolvedMetadataPath))
			{
				return _resolvedMetadataPath;
			}
			string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			string[] array = new string[8]
			{
				Path.Combine(directoryName, "templates", "templates_metadata.json"),
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates", "templates_metadata.json"),
				Path.Combine(directoryName, "..", "templates", "templates_metadata.json"),
				Path.Combine(directoryName, "..", "..", "templates", "templates_metadata.json"),
				"D:\\Tool\\V-Format Addin\\templates\\templates_metadata.json",
				Path.Combine(Constants.ProgramDirectory, "templates", "templates_metadata.json"),
				Path.Combine(Constants.AppDataDirectory, "templates", "templates_metadata.json"),
				"d:\\Tool\\V-Format\\templates\\templates_metadata.json"
			};
			string[] array2 = array;
			foreach (string path in array2)
			{
				try
				{
					if (File.Exists(path))
					{
						_resolvedMetadataPath = Path.GetFullPath(path);
						return _resolvedMetadataPath;
					}
				}
				catch
				{
				}
			}
			return Path.Combine(Constants.AppDataDirectory, "templates", "templates_metadata.json");
		}
	}

	public void Reload()
	{
		lock (_lock)
		{
			_allTemplates = null;
			_resolvedMetadataPath = null;
		}
	}

	public List<TemplateItem> GetAllTemplates()
	{
		lock (_lock)
		{
			if (_allTemplates != null)
			{
				return _allTemplates;
			}
			List<TemplateItem> list = new List<TemplateItem>();
			string metadataFilePath = MetadataFilePath;
			if (File.Exists(metadataFilePath))
			{
				try
				{
					string text = File.ReadAllText(metadataFilePath, Encoding.UTF8);
					if (!string.IsNullOrWhiteSpace(text))
					{
						List<TemplateItem> list2 = _serializer.Deserialize<List<TemplateItem>>(text);
						if (list2 != null)
						{
							list.AddRange(list2);
						}
					}
				}
				catch
				{
				}
			}
			try
			{
				List<UserTemplateItem> allTemplates = UserTemplateService.Instance.GetAllTemplates();
				if (allTemplates != null && allTemplates.Count > 0)
				{
					foreach (UserTemplateItem item in allTemplates)
					{
						list.Add(new TemplateItem
						{
							Id = item.Id,
							Name = item.Name,
							Title = item.Name,
							ShortTitle = item.Name,
							File = item.FileName,
							Category = "Biểu mẫu của tôi",
							Group = "Mẫu cá nhân của tôi",
							TemplateNumber = "Cá nhân",
							Source = "Thư viện cá nhân",
							Description = (string.IsNullOrEmpty(item.Description) ? "Biểu mẫu người dùng tự thêm vào hệ thống" : item.Description),
							IsUserTemplate = true,
							Keywords = new List<string>
							{
								"ca nhan",
								"bieu mau cua toi",
								item.Name.ToLower()
							},
							Type = "USER_TEMPLATE"
						});
					}
				}
			}
			catch
			{
			}
			_allTemplates = list;
			return _allTemplates;
		}
	}

	public IReadOnlyList<TemplateCategoryItem> GetCategories()
	{
		List<TemplateItem> allTemplates = GetAllTemplates();
		List<TemplateCategoryItem> list = new List<TemplateCategoryItem>();
		List<string> list2 = new List<string>();
		list2.Add("Thi đua – Khen thưởng");
		list2.Add("Văn bản hành chính");
		list2.Add("Lực lượng vũ trang & Quân đội");
		list2.Add("Văn bản Đảng");
		list2.Add("Văn bản Đoàn TNCS Hồ Chí Minh");
		list2.Add("Đơn thư & Cá nhân");
		list2.Add("Chuyên ngành & Giáo dục");
		list2.Add("Biểu mẫu của tôi");
		List<string> list3 = list2;
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		dictionary.Add("Thi đua – Khen thưởng", "\ud83c\udfc6");
		dictionary.Add("Văn bản hành chính", "\ud83c\udfdb\ufe0f");
		dictionary.Add("Văn bản Đảng", "\ud83d\udea9");
		dictionary.Add("Văn bản Đoàn TNCS Hồ Chí Minh", "\ud83c\udf1f");
		dictionary.Add("Đơn thư & Cá nhân", "✍\ufe0f");
		dictionary.Add("Lực lượng vũ trang & Quân đội", "\ud83c\udf96\ufe0f");
		dictionary.Add("Chuyên ngành & Giáo dục", "\ud83d\udcd6");
		dictionary.Add("Biểu mẫu của tôi", "⭐");
		Dictionary<string, string> dictionary2 = dictionary;
		Dictionary<string, string> dictionary3 = new Dictionary<string, string>();
		dictionary3.Add("Thi đua – Khen thưởng", "Đầy đủ 13 mẫu Phụ lục Nghị định 152/2025/NĐ-CP chuẩn Quốc gia");
		dictionary3.Add("Văn bản hành chính", "Biểu mẫu thể thức hành chính chuẩn Nghị định 30/2020/NĐ-CP");
		dictionary3.Add("Văn bản Đảng", "Biểu mẫu văn bản Đảng chuẩn Hướng dẫn 05-HD/VPTW");
		dictionary3.Add("Văn bản Đoàn TNCS Hồ Chí Minh", "Biểu mẫu văn bản Đoàn chuẩn Hướng dẫn 29-HD/TWĐTN-VP");
		dictionary3.Add("Đơn thư & Cá nhân", "Đơn từ, tường trình, cam kết và giấy ủy quyền chuẩn thể thức");
		dictionary3.Add("Lực lượng vũ trang & Quân đội", "Biểu mẫu văn bản Quân đội theo Thông tư 68/2024/TT-BQP");
		dictionary3.Add("Chuyên ngành & Giáo dục", "Văn bản hợp nhất (VBHN), Kế hoạch bài dạy 5512...");
		dictionary3.Add("Biểu mẫu của tôi", "Kho lưu trữ biểu mẫu cá nhân hóa do bạn quản lý và thêm mới");
		Dictionary<string, string> dictionary4 = dictionary3;
		foreach (string catName in list3)
		{
			Func<TemplateItem, bool> predicate = (TemplateItem t) => string.Equals(t.Category, catName, StringComparison.OrdinalIgnoreCase);
			List<TemplateItem> list4 = allTemplates.Where(predicate).ToList();
			if (list4.Count > 0 || catName == "Biểu mẫu của tôi")
			{
				List<string> groups = list4.Select((TemplateItem t) => t.Group ?? "Chung").Distinct().ToList();
				list.Add(new TemplateCategoryItem
				{
					Id = catName.ToLower().Replace(" ", "_").Replace("–", "-"),
					Name = catName,
					Icon = (dictionary2.ContainsKey(catName) ? dictionary2[catName] : "\ud83d\udcc4"),
					Description = (dictionary4.ContainsKey(catName) ? dictionary4[catName] : ""),
					TemplateCount = list4.Count,
					Groups = groups
				});
			}
		}
		return list;
	}

	public IReadOnlyList<TemplateGroupSection> GetGroupSectionsByCategory(string categoryName, string searchQuery)
	{
		List<TemplateItem> allTemplates = GetAllTemplates();
		IEnumerable<TemplateItem> source = allTemplates.Where((TemplateItem t) => string.Equals(t.Category, categoryName, StringComparison.OrdinalIgnoreCase));
		if (!string.IsNullOrWhiteSpace(searchQuery))
		{
			string normQuery = NormalizeForSearch(searchQuery);
			source = source.Where((TemplateItem t) => MatchesTemplate(t, normQuery));
		}
		List<TemplateItem> source2 = source.ToList();
		List<string> list = null;
		if (string.Equals(categoryName, "Thi đua – Khen thưởng", StringComparison.OrdinalIgnoreCase))
		{
			List<string> list2 = new List<string>();
			list2.Add("Hồ sơ đề nghị");
			list2.Add("Danh hiệu thi đua");
			list2.Add("Khen thưởng công trạng");
			list2.Add("Khen thưởng đột xuất");
			list2.Add("Khen thưởng quá trình cống hiến");
			list2.Add("Khen thưởng chuyên đề");
			list2.Add("Danh hiệu Anh hùng");
			list2.Add("Khen thưởng đối ngoại");
			list2.Add("Sáng kiến – Khoa học, công nghệ và đổi mới sáng tạo");
			list = list2;
		}
		IEnumerable<IGrouping<string, TemplateItem>> enumerable = from t in source2
			group t by t.Group ?? "Khác";
		List<TemplateGroupSection> list3 = new List<TemplateGroupSection>();
		if (list != null)
		{
			foreach (string gName in list)
			{
				Func<TemplateItem, bool> predicate = (TemplateItem t) => string.Equals(t.Group, gName, StringComparison.OrdinalIgnoreCase);
				List<TemplateItem> list4 = source2.Where(predicate).ToList();
				if (list4.Count > 0)
				{
					list3.Add(new TemplateGroupSection
					{
						GroupName = gName,
						Templates = list4
					});
				}
			}
			foreach (IGrouping<string, TemplateItem> grp in enumerable)
			{
				List<string> source3 = list;
				Func<string, bool> predicate2 = (string g) => string.Equals(g, grp.Key, StringComparison.OrdinalIgnoreCase);
				if (!source3.Any(predicate2))
				{
					list3.Add(new TemplateGroupSection
					{
						GroupName = grp.Key,
						Templates = grp.ToList()
					});
				}
			}
		}
		else
		{
			foreach (IGrouping<string, TemplateItem> item in enumerable)
			{
				list3.Add(new TemplateGroupSection
				{
					GroupName = item.Key,
					Templates = item.ToList()
				});
			}
		}
		return list3;
	}

	public IReadOnlyList<TemplateGroupSection> SearchAll(string searchQuery)
	{
		if (string.IsNullOrWhiteSpace(searchQuery))
		{
			return new List<TemplateGroupSection>();
		}
		string normQuery = NormalizeForSearch(searchQuery);
		List<TemplateItem> allTemplates = GetAllTemplates();
		List<TemplateItem> source = allTemplates.Where((TemplateItem t) => MatchesTemplate(t, normQuery)).ToList();
		IEnumerable<IGrouping<string, TemplateItem>> enumerable = from t in source
			group t by t.Category ?? "Khác";
		List<TemplateGroupSection> list = new List<TemplateGroupSection>();
		foreach (IGrouping<string, TemplateItem> item in enumerable)
		{
			list.Add(new TemplateGroupSection
			{
				GroupName = item.Key,
				Templates = item.ToList()
			});
		}
		return list;
	}

	public TemplateItem GetTemplateById(string id)
	{
		if (string.IsNullOrEmpty(id))
		{
			return null;
		}
		return GetAllTemplates().FirstOrDefault((TemplateItem t) => string.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase));
	}

	public string GetTemplateFullPath(TemplateItem item)
	{
		if (item == null)
		{
			return null;
		}
		if (item.IsUserTemplate)
		{
			UserTemplateItem userTemplateItem = UserTemplateService.Instance.GetTemplateById(item.Id);
			if (userTemplateItem == null)
			{
				UserTemplateItem userTemplateItem2 = new UserTemplateItem();
				userTemplateItem2.FileName = item.File;
				userTemplateItem = userTemplateItem2;
			}
			UserTemplateItem item2 = userTemplateItem;
			string text = UserTemplateService.Instance.GetTemplateFullPath(item2);
			if (string.IsNullOrEmpty(text) || !File.Exists(text) || new FileInfo(text).Length < 500)
			{
				text = UserTemplateService.Instance.EnsureValidTemplateFile(item2);
			}
			return text;
		}
		string file = item.File;
		if (string.IsNullOrEmpty(file))
		{
			return null;
		}
		string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
		string[] array = new string[8]
		{
			Path.Combine(directoryName, "templates", file),
			Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates", file),
			Path.Combine(directoryName, "..", "templates", file),
			Path.Combine(directoryName, "..", "..", "templates", file),
			Path.Combine("D:\\Tool\\V-Format Addin\\templates", file),
			Path.Combine(Constants.ProgramDirectory, "templates", file),
			Path.Combine(Constants.AppDataDirectory, "templates", file),
			Path.Combine("d:\\Tool\\V-Format\\templates", file)
		};
		string[] array2 = array;
		foreach (string path in array2)
		{
			try
			{
				if (File.Exists(path))
				{
					return Path.GetFullPath(path);
				}
			}
			catch
			{
			}
		}
		return file;
	}

	private bool MatchesTemplate(TemplateItem item, string normalizedQuery)
	{
		if (item == null || string.IsNullOrEmpty(normalizedQuery))
		{
			return false;
		}
		if (!string.IsNullOrEmpty(item.TemplateNumber))
		{
			string text = NormalizeForSearch(item.TemplateNumber);
			if (text.Contains(normalizedQuery) || normalizedQuery.Contains(text))
			{
				return true;
			}
			if (("mau " + text).Contains(normalizedQuery))
			{
				return true;
			}
			if (("mau so " + text).Contains(normalizedQuery))
			{
				return true;
			}
		}
		if (!string.IsNullOrEmpty(item.Title) && NormalizeForSearch(item.Title).Contains(normalizedQuery))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(item.ShortTitle) && NormalizeForSearch(item.ShortTitle).Contains(normalizedQuery))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(item.Name) && NormalizeForSearch(item.Name).Contains(normalizedQuery))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(item.Description) && NormalizeForSearch(item.Description).Contains(normalizedQuery))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(item.Group) && NormalizeForSearch(item.Group).Contains(normalizedQuery))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(item.Category) && NormalizeForSearch(item.Category).Contains(normalizedQuery))
		{
			return true;
		}
		if (!string.IsNullOrEmpty(item.Source) && NormalizeForSearch(item.Source).Contains(normalizedQuery))
		{
			return true;
		}
		if (item.Keywords != null)
		{
			foreach (string keyword in item.Keywords)
			{
				if (!string.IsNullOrEmpty(keyword) && NormalizeForSearch(keyword).Contains(normalizedQuery))
				{
					return true;
				}
			}
		}
		return false;
	}

	private string NormalizeForSearch(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return "";
		}
		string text2 = text.ToLowerInvariant().Trim();
		text2 = RemoveDiacritics(text2);
		return text2.Replace("–", "-").Replace("—", "-");
	}

	private string RemoveDiacritics(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return "";
		}
		string text2 = text.Normalize(NormalizationForm.FormD);
		StringBuilder stringBuilder = new StringBuilder();
		string text3 = text2;
		foreach (char c in text3)
		{
			UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
			if (unicodeCategory != UnicodeCategory.NonSpacingMark)
			{
				stringBuilder.Append(c);
			}
		}
		string text4 = stringBuilder.ToString().Normalize(NormalizationForm.FormC);
		return text4.Replace('đ', 'd').Replace('Đ', 'd');
	}
}
