using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using VFormat.Core.SpellCheck.Processing;

namespace VFormat.Core.SpellCheck.Engine;

public class SpellDictionaryManager
{
	private static SpellDictionaryManager _instance;

	private static readonly object _lock = new object();

	private readonly HashSet<string> _masterSet;

	private readonly HashSet<string> _coreSet;

	private readonly HashSet<string> _adminLegalSet;

	private readonly HashSet<string> _properNamesSet;

	private readonly HashSet<string> _abbreviationsSet;

	private readonly HashSet<string> _foreignTermsSet;

	private readonly HashSet<string> _userSet;

	private readonly HashSet<string> _ignoredSet;

	private readonly Dictionary<string, TypoMapEntry> _typoMap;

	private readonly Dictionary<string, ConfusablePhraseEntry> _confusableMap;

	private readonly Dictionary<string, int> _frequencyMap;

	private readonly Dictionary<string, int> _bigramFrequencyMap;

	private readonly Dictionary<int, List<string>> _lengthBuckets;

	private readonly HashSet<string> _foldedAccentedForms;

	private string _dataDirectory;

	private string _userDicPath;

	private string _legacyUserDicPath;

	private readonly string _userDicPathOverride;

	public int MaxConfusablePhraseWords { get; private set; }

	public static SpellDictionaryManager Instance
	{
		get
		{
			if (_instance == null)
			{
				lock (_lock)
				{
					if (_instance == null)
					{
						_instance = new SpellDictionaryManager();
					}
				}
			}
			return _instance;
		}
	}

	public ICollection<string> ConfusablePhraseKeys => _confusableMap.Keys;

	public int MasterWordCount => _masterSet.Count;

	public int CoreWordCount => _coreSet.Count;

	public int AdminWordCount => _adminLegalSet.Count;

	public int ProperNameCount => _properNamesSet.Count;

	public int AbbrCount => _abbreviationsSet.Count;

	public int ForeignTermCount => _foreignTermsSet.Count;

	public SpellDictionaryManager(string customDataDir = null)
		: this(customDataDir, null)
	{
	}

	internal SpellDictionaryManager(string customDataDir, string userDicPath)
	{
		_userDicPathOverride = userDicPath;
		_masterSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_coreSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_adminLegalSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_properNamesSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_abbreviationsSet = new HashSet<string>(StringComparer.Ordinal);
		_foreignTermsSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_userSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_ignoredSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		_typoMap = new Dictionary<string, TypoMapEntry>(StringComparer.OrdinalIgnoreCase);
		_confusableMap = new Dictionary<string, ConfusablePhraseEntry>(StringComparer.OrdinalIgnoreCase);
		_frequencyMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		_bigramFrequencyMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		_lengthBuckets = new Dictionary<int, List<string>>();
		_foldedAccentedForms = new HashSet<string>(StringComparer.Ordinal);
		LocateAndLoadData(customDataDir);
	}

	private void LocateAndLoadData(string customDataDir)
	{
		if (!string.IsNullOrEmpty(customDataDir) && Directory.Exists(customDataDir))
		{
			_dataDirectory = customDataDir;
		}
		else
		{
			string codeBase = Assembly.GetExecutingAssembly().CodeBase;
			UriBuilder uriBuilder = new UriBuilder(codeBase);
			string path = Uri.UnescapeDataString(uriBuilder.Path);
			string directoryName = Path.GetDirectoryName(path);
			string text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VFormat", "SpellCheck", "Data");
			string[] array = new string[7]
			{
				Path.Combine(directoryName, "SpellCheck", "Data"),
				text,
				Path.Combine(directoryName, "Data"),
				Path.Combine(directoryName, "..", "..", "VFormat.Core", "SpellCheck", "Data"),
				Path.Combine(directoryName, "..", "..", "src", "spellcheck", "dictionaries"),
				"D:\\Tool\\V-Format Addin\\VFormat.Core\\SpellCheck\\Data",
				"D:\\Tool\\V-Format Addin\\src\\spellcheck"
			};
			for (int i = 0; i < array.Length; i++)
			{
				if (Directory.Exists(array[i]))
				{
					_dataDirectory = Path.GetFullPath(array[i]);
					break;
				}
			}
		}
		if (string.IsNullOrEmpty(_dataDirectory) || !Directory.Exists(_dataDirectory))
		{
			_dataDirectory = AppDomain.CurrentDomain.BaseDirectory;
		}
		LoadAllDictionaries();
		LoadTypoMap();
		LoadConfusablePhrases();
		LoadFrequencies();
		BuildBuckets();
	}

	private void LoadAllDictionaries()
	{
		LoadWordsFromFile(ResolveFilePath("vi_core.dic"), _coreSet);
		LoadWordsFromFile(ResolveFilePath("vi_admin_legal.dic"), _adminLegalSet);
		LoadWordsFromFile(ResolveFilePath("vi_proper_names.dic"), _properNamesSet);
		LoadWordsFromFile(ResolveFilePath("vi_abbreviations.dic"), _abbreviationsSet);
		LoadWordsFromFile(ResolveFilePath("vi_foreign_terms.dic"), _foreignTermsSet);
		string text = _userDicPathOverride ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VFormat", "vi_user.dic");
		string text2 = ResolveFilePath("vi_user.dic");
		_legacyUserDicPath = null;
		if (!string.IsNullOrEmpty(text2) && File.Exists(text2) && !string.Equals(Path.GetFullPath(text2), Path.GetFullPath(text), StringComparison.OrdinalIgnoreCase))
		{
			_legacyUserDicPath = text2;
			LoadWordsFromFile(text2, _userSet);
		}
		if (File.Exists(text))
		{
			LoadWordsFromFile(text, _userSet);
		}
		_userDicPath = text;
		foreach (string item in _coreSet)
		{
			_masterSet.Add(item);
		}
		foreach (string item2 in _adminLegalSet)
		{
			_masterSet.Add(item2);
		}
		foreach (string item3 in _properNamesSet)
		{
			_masterSet.Add(item3);
		}
		foreach (string item4 in _foreignTermsSet)
		{
			_masterSet.Add(item4);
		}
		foreach (string item5 in _userSet)
		{
			_masterSet.Add(item5);
		}
		LoadWordsFromFile(ResolveFilePath("vi_vformat.dic"), _masterSet);
	}

	private string ResolveFilePath(string fileName)
	{
		if (string.IsNullOrEmpty(_dataDirectory))
		{
			return fileName;
		}
		string text = Path.Combine(_dataDirectory, fileName);
		if (File.Exists(text))
		{
			return text;
		}
		string[] array = new string[4] { "dictionaries", "rules", "generated", "Data" };
		for (int i = 0; i < array.Length; i++)
		{
			string text2 = Path.Combine(_dataDirectory, array[i], fileName);
			if (File.Exists(text2))
			{
				return text2;
			}
			string path = Path.Combine(_dataDirectory, "..", array[i], fileName);
			if (File.Exists(path))
			{
				return Path.GetFullPath(path);
			}
		}
		return text;
	}

	private void LoadWordsFromFile(string filePath, HashSet<string> targetSet)
	{
		if (!File.Exists(filePath))
		{
			return;
		}
		try
		{
			using StreamReader streamReader = new StreamReader(filePath, Encoding.UTF8);
			string text;
			while ((text = streamReader.ReadLine()) != null)
			{
				string text2 = VietnameseUnicodeNormalizer.NormalizeNfc(text.Trim());
				if (text2.Length <= 0 || text2.StartsWith("#"))
				{
					continue;
				}
				targetSet.Add(text2);
				if (text2.IndexOf(' ') < 0)
				{
					continue;
				}
				string[] array = text2.Split(new char[5] { ' ', '-', '–', '—', '/' }, StringSplitOptions.RemoveEmptyEntries);
				for (int i = 0; i < array.Length; i++)
				{
					string text3 = array[i].Trim();
					if (text3.Length > 0 && char.IsLetter(text3[0]))
					{
						targetSet.Add(text3);
					}
				}
			}
		}
		catch
		{
		}
	}

	private void LoadTypoMap()
	{
		string path = ResolveFilePath("typo_map.json");
		if (!File.Exists(path))
		{
			return;
		}
		try
		{
			string input = File.ReadAllText(path, Encoding.UTF8);
			JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
			Dictionary<string, object> dictionary = javaScriptSerializer.Deserialize<Dictionary<string, object>>(input);
			if (dictionary == null)
			{
				return;
			}
			foreach (KeyValuePair<string, object> item in dictionary)
			{
				if (!(item.Value is Dictionary<string, object> dictionary2))
				{
					continue;
				}
				TypoMapEntry typoMapEntry = new TypoMapEntry();
				typoMapEntry.suggestions = new List<string>();
				if (dictionary2.ContainsKey("suggestions") && dictionary2["suggestions"] is ArrayList arrayList)
				{
					foreach (object item2 in arrayList)
					{
						typoMapEntry.suggestions.Add(item2.ToString());
					}
				}
				if (dictionary2.ContainsKey("confidence"))
				{
					typoMapEntry.confidence = Convert.ToDouble(dictionary2["confidence"]);
				}
				if (dictionary2.ContainsKey("category"))
				{
					typoMapEntry.category = dictionary2["category"].ToString();
				}
				_typoMap[item.Key] = typoMapEntry;
			}
		}
		catch
		{
		}
	}

	private void LoadConfusablePhrases()
	{
		MaxConfusablePhraseWords = 2;
		string path = ResolveFilePath("confusable_phrases.json");
		if (!File.Exists(path))
		{
			return;
		}
		try
		{
			string input = File.ReadAllText(path, Encoding.UTF8);
			JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
			Dictionary<string, object> dictionary = javaScriptSerializer.Deserialize<Dictionary<string, object>>(input);
			if (dictionary == null)
			{
				return;
			}
			foreach (KeyValuePair<string, object> item in dictionary)
			{
				if (!(item.Value is Dictionary<string, object> dictionary2))
				{
					continue;
				}
				ConfusablePhraseEntry confusablePhraseEntry = new ConfusablePhraseEntry();
				if (dictionary2.ContainsKey("replacement"))
				{
					confusablePhraseEntry.replacement = dictionary2["replacement"].ToString();
				}
				if (dictionary2.ContainsKey("description"))
				{
					confusablePhraseEntry.description = dictionary2["description"].ToString();
				}
				if (dictionary2.ContainsKey("category"))
				{
					confusablePhraseEntry.category = dictionary2["category"].ToString();
				}
				if (dictionary2.ContainsKey("severity"))
				{
					confusablePhraseEntry.severity = dictionary2["severity"].ToString();
				}
				if (dictionary2.ContainsKey("source"))
				{
					confusablePhraseEntry.source = dictionary2["source"].ToString();
				}
				if (dictionary2.ContainsKey("when_to_use") && dictionary2["when_to_use"] != null)
				{
					confusablePhraseEntry.when_to_use = dictionary2["when_to_use"].ToString();
				}
				if (dictionary2.ContainsKey("when_not_to_use") && dictionary2["when_not_to_use"] != null)
				{
					confusablePhraseEntry.when_not_to_use = dictionary2["when_not_to_use"].ToString();
				}
				if (dictionary2.ContainsKey("usage_guide") && dictionary2["usage_guide"] != null)
				{
					confusablePhraseEntry.usage_guide = dictionary2["usage_guide"].ToString();
				}
				if (dictionary2.ContainsKey("examples") && dictionary2["examples"] != null && dictionary2["examples"] is ArrayList arrayList)
				{
					confusablePhraseEntry.examples = new List<string>();
					foreach (object item2 in arrayList)
					{
						if (item2 != null)
						{
							confusablePhraseEntry.examples.Add(item2.ToString());
						}
					}
				}
				confusablePhraseEntry.not_after = ReadLowerList(dictionary2, "not_after");
				confusablePhraseEntry.not_before = ReadLowerList(dictionary2, "not_before");
				confusablePhraseEntry.only_before = ReadLowerList(dictionary2, "only_before");
				_confusableMap[VietnameseUnicodeNormalizer.NormalizeNfc(item.Key)] = confusablePhraseEntry;
				MaxConfusablePhraseWords = Math.Max(MaxConfusablePhraseWords, Math.Min(32, item.Key.Split(new char[1] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length));
			}
		}
		catch
		{
		}
	}

	private static List<string> ReadLowerList(Dictionary<string, object> entryObj, string key)
	{
		if (!entryObj.ContainsKey(key) || entryObj[key] == null)
		{
			return null;
		}
		if (!(entryObj[key] is ArrayList arrayList))
		{
			return null;
		}
		List<string> list = new List<string>();
		foreach (object item in arrayList)
		{
			if (item != null)
			{
				string text = VietnameseUnicodeNormalizer.NormalizeNfc(item.ToString().Trim());
				if (text.Length > 0)
				{
					list.Add((text == "<END>") ? text : text.ToLowerInvariant());
				}
			}
		}
		if (list.Count <= 0)
		{
			return null;
		}
		return list;
	}

	private void LoadFrequencies()
	{
		string path = ResolveFilePath("frequency.json");
		if (File.Exists(path))
		{
			try
			{
				string input = File.ReadAllText(path, Encoding.UTF8);
				JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
				Dictionary<string, object> dictionary = javaScriptSerializer.Deserialize<Dictionary<string, object>>(input);
				if (dictionary != null)
				{
					foreach (KeyValuePair<string, object> item in dictionary)
					{
						_frequencyMap[item.Key] = Convert.ToInt32(item.Value);
					}
				}
			}
			catch
			{
			}
		}
		string path2 = ResolveFilePath("bigram_frequency.json");
		if (!File.Exists(path2))
		{
			return;
		}
		try
		{
			string input2 = File.ReadAllText(path2, Encoding.UTF8);
			JavaScriptSerializer javaScriptSerializer2 = new JavaScriptSerializer();
			Dictionary<string, object> dictionary2 = javaScriptSerializer2.Deserialize<Dictionary<string, object>>(input2);
			if (dictionary2 == null)
			{
				return;
			}
			foreach (KeyValuePair<string, object> item2 in dictionary2)
			{
				_bigramFrequencyMap[item2.Key] = Convert.ToInt32(item2.Value);
			}
		}
		catch
		{
		}
	}

	private void BuildBuckets()
	{
		_lengthBuckets.Clear();
		_foldedAccentedForms.Clear();
		foreach (string item in _masterSet)
		{
			string text = item.ToLowerInvariant();
			string text2 = FoldDiacritics(text);
			if (!string.Equals(text2, text, StringComparison.Ordinal))
			{
				_foldedAccentedForms.Add(text2);
			}
			int length = item.Length;
			if (!_lengthBuckets.TryGetValue(length, out var value))
			{
				value = new List<string>();
				_lengthBuckets[length] = value;
			}
			value.Add(item);
		}
	}

	public bool ContainsWord(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return false;
		}
		string text = VietnameseUnicodeNormalizer.NormalizeNfc(word);
		if (_masterSet.Contains(text) || _ignoredSet.Contains(text))
		{
			return true;
		}
		string alternativeTonePlacement = VietnameseUnicodeNormalizer.GetAlternativeTonePlacement(text);
		if (!string.IsNullOrEmpty(alternativeTonePlacement) && (_masterSet.Contains(alternativeTonePlacement) || _ignoredSet.Contains(alternativeTonePlacement)))
		{
			return true;
		}
		return false;
	}

	public static string FoldDiacritics(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return string.Empty;
		}
		string text2 = text.Normalize(NormalizationForm.FormD);
		StringBuilder stringBuilder = new StringBuilder(text2.Length);
		for (int i = 0; i < text2.Length; i++)
		{
			char c = text2[i];
			if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
			{
				switch (c)
				{
				case 'đ':
					c = 'd';
					break;
				case 'Đ':
					c = 'D';
					break;
				}
				stringBuilder.Append(c);
			}
		}
		return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
	}

	public bool HasAccentedVariant(string asciiWord)
	{
		if (string.IsNullOrEmpty(asciiWord))
		{
			return false;
		}
		return _foldedAccentedForms.Contains(FoldDiacritics(asciiWord.ToLowerInvariant()));
	}

	public bool IsUserAccepted(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		string item = VietnameseUnicodeNormalizer.NormalizeNfc(text.Trim());
		if (!_ignoredSet.Contains(item))
		{
			return _userSet.Contains(item);
		}
		return true;
	}

	public bool IsIgnored(string text)
	{
		if (!string.IsNullOrEmpty(text))
		{
			return _ignoredSet.Contains(VietnameseUnicodeNormalizer.NormalizeNfc(text.Trim()));
		}
		return false;
	}

	public bool IsAbbreviation(string token)
	{
		if (string.IsNullOrEmpty(token))
		{
			return false;
		}
		return _abbreviationsSet.Contains(token);
	}

	public bool IsProperName(string token)
	{
		if (string.IsNullOrEmpty(token))
		{
			return false;
		}
		string item = VietnameseUnicodeNormalizer.NormalizeNfc(token);
		return _properNamesSet.Contains(item);
	}

	public bool IsAdminLegal(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return false;
		}
		string item = VietnameseUnicodeNormalizer.NormalizeNfc(word);
		return _adminLegalSet.Contains(item);
	}

	public bool IsForeignTerm(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return false;
		}
		string item = VietnameseUnicodeNormalizer.NormalizeNfc(word);
		return _foreignTermsSet.Contains(item);
	}

	public bool TryGetTypoRule(string token, out TypoMapEntry entry)
	{
		string key = VietnameseUnicodeNormalizer.NormalizeNfc(token);
		return _typoMap.TryGetValue(key, out entry);
	}

	public bool TryGetConfusablePhrase(string phrase, out ConfusablePhraseEntry entry)
	{
		if (string.IsNullOrEmpty(phrase))
		{
			entry = null;
			return false;
		}
		string key = VietnameseUnicodeNormalizer.NormalizeNfc(phrase.Trim());
		return _confusableMap.TryGetValue(key, out entry);
	}

	public int GetWordFrequency(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return 0;
		}
		string key = VietnameseUnicodeNormalizer.NormalizeNfc(word).ToLowerInvariant();
		if (_frequencyMap.TryGetValue(key, out var value))
		{
			return value;
		}
		return 0;
	}

	public int GetBigramFrequency(string w1, string w2)
	{
		if (string.IsNullOrEmpty(w1) || string.IsNullOrEmpty(w2))
		{
			return 0;
		}
		string key = $"{w1.Trim().ToLowerInvariant()} {w2.Trim().ToLowerInvariant()}";
		if (_bigramFrequencyMap.TryGetValue(key, out var value))
		{
			return value;
		}
		return 0;
	}

	public IEnumerable<string> GetCandidatePool(int targetLength)
	{
		for (int l = targetLength - 2; l <= targetLength + 2; l++)
		{
			if (_lengthBuckets.TryGetValue(l, out var list))
			{
				for (int i = 0; i < list.Count; i++)
				{
					yield return list[i];
				}
			}
		}
	}

	private void MigrateLegacyUserDictionary()
	{
		if (string.IsNullOrEmpty(_legacyUserDicPath) || File.Exists(_userDicPath) || !File.Exists(_legacyUserDicPath))
		{
			return;
		}
		try
		{
			File.Copy(_legacyUserDicPath, _userDicPath, overwrite: false);
		}
		catch (IOException)
		{
			if (!File.Exists(_userDicPath))
			{
				throw;
			}
		}
	}

	private static bool NeedsLeadingNewLine(string path)
	{
		if (!File.Exists(path))
		{
			return false;
		}
		using FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
		if (fileStream.Length == 0)
		{
			return false;
		}
		fileStream.Seek(-1L, SeekOrigin.End);
		int num = fileStream.ReadByte();
		return num != 10 && num != 13;
	}

	public void AddUserWord(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return;
		}
		string text = VietnameseUnicodeNormalizer.NormalizeNfc(word.Trim());
		if (!_userSet.Contains(text))
		{
			if (string.IsNullOrEmpty(_userDicPath))
			{
				_userDicPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VFormat", "vi_user.dic");
			}
			Directory.CreateDirectory(Path.GetDirectoryName(_userDicPath));
			MigrateLegacyUserDictionary();
			byte[] bytes = Encoding.UTF8.GetBytes(text + Environment.NewLine);
			if (NeedsLeadingNewLine(_userDicPath))
			{
				bytes = Encoding.UTF8.GetBytes(Environment.NewLine + text + Environment.NewLine);
			}
			using (FileStream fileStream = new FileStream(_userDicPath, FileMode.Append, FileAccess.Write, FileShare.Read))
			{
				fileStream.Write(bytes, 0, bytes.Length);
				fileStream.Flush(flushToDisk: true);
			}
			_userSet.Add(text);
			_masterSet.Add(text);
			if (!_lengthBuckets.TryGetValue(text.Length, out var value))
			{
				value = new List<string>();
				_lengthBuckets[text.Length] = value;
			}
			value.Add(text);
		}
	}

	public void IgnoreWord(string word)
	{
		if (!string.IsNullOrEmpty(word))
		{
			_ignoredSet.Add(VietnameseUnicodeNormalizer.NormalizeNfc(word.Trim()));
		}
	}
}
