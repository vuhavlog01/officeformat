#define TRACE
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using VFormat.Addin.Registration;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Personalization;
using VFormat.Core.Rules;
using VFormat.Core.Serialization;
using VFormat.Core.Templates;
using VFormat.License;
using VFormat.License.Client;
using VFormat.License.Models;
using VFormat.UI.Common;
using VFormat.UI.License;
using VFormat.UI.Personalization;
using VFormat.UI.SpellCheck;
using VFormat.UI.Templates;
using VFormat.WordEngine;
using VFormat.WordEngine.Interop;

namespace VFormat.Addin;

[ClassInterface(ClassInterfaceType.AutoDual)]
[ProgId("VFormat.Addin")]
[ComVisible(true)]
[Guid("7F28B490-A833-4F85-B2A8-3E2A398246E1")]
public class Addin : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
{
	private static Dictionary<string, int> _regDict;

	private sealed class DocumentPrintLayoutState
	{
		public PrintLayoutMode Mode;
	}

	private enum LongOperationOutcome
	{
		Completed,
		Cancelled,
		Failed,
		Busy
	}

	private Microsoft.Office.Interop.Word.Application _wordApp;

	private ICTPFactory _ctpFactory;

	private CustomTaskPane _taskPane;

	private IRibbonUI _ribbon;

	private TaskPaneControl _taskPaneControl;

	private WordFormatEngine _engineInstance;

	private Timer _documentDetectionTimer;

	private bool _disconnecting;

	private bool _startupReady;

	private JsonProfileRepository _profileRepoInstance;

	private LicenseManager _licenseManagerInstance;

	private RulePackType _selectedRulePack = RulePackType.ND30_2020;

	private int _selectedRegulationIndex;

	private int _selectedOrgLevelIndex;

	private int _selectedDocTypeIndex;

	private PrintLayoutMode _selectedPrintLayout;

	private bool _syncingTaskPane;

	private int _selectedPrintLayoutIndex;

	private readonly ConditionalWeakTable<Document, DocumentPrintLayoutState> _docPrintLayouts = new ConditionalWeakTable<Document, DocumentPrintLayoutState>();

	private readonly HashSet<string> _autoDetectedDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

	private readonly Dictionary<string, RulePackType> _docSelectedRules = new Dictionary<string, RulePackType>(StringComparer.OrdinalIgnoreCase);

	private readonly HashSet<string> _manuallySelectedDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

	private readonly Dictionary<string, string> _docDetectionFingerprints = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

	private readonly Dictionary<string, RulePackDetectionResult> _docDetectionResults = new Dictionary<string, RulePackDetectionResult>(StringComparer.OrdinalIgnoreCase);

	private static bool _suppressDialogs;

	private bool _longOperationRunning;

	internal static bool SuppressedConfirmationResult;

	private WordFormatEngine _engine => _engineInstance ?? (_engineInstance = new WordFormatEngine());

	private JsonProfileRepository _profileRepo => _profileRepoInstance ?? (_profileRepoInstance = new JsonProfileRepository());

	private LicenseManager _licenseManager => _licenseManagerInstance ?? (_licenseManagerInstance = new LicenseManager());

	public static bool SuppressDialogs
	{
		get
		{
			return _suppressDialogs;
		}
		set
		{
			_suppressDialogs = value;
		}
	}

	private static string VersionLabel
	{
		get
		{
			try
			{
				return "OfficeFormat 1.0.0";
			}
			catch
			{
				return "OfficeFormat";
			}
		}
	}

	static Addin()
	{
		ClearWordResiliencyMarkers();
		_suppressDialogs = false;
		SuppressedConfirmationResult = true;
		AppDomain.CurrentDomain.AssemblyResolve += ResolveAddinAssembly;
	}

	private static Assembly ResolveAddinAssembly(object sender, ResolveEventArgs args)
	{
		try
		{
			string path = new AssemblyName(args.Name).Name + ".dll";
			string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			if (!string.IsNullOrEmpty(directoryName))
			{
				string text = Path.Combine(directoryName, path);
				if (File.Exists(text))
				{
					return Assembly.LoadFrom(text);
				}
			}
		}
		catch
		{
		}
		return null;
	}

	private DocumentPrintLayoutState CreatePrintLayoutState(Document doc)
	{
		PrintLayoutMode mode = PrintLayoutMode.SingleSided;
		try
		{
			if (doc.PageSetup.MirrorMargins == -1)
			{
				mode = PrintLayoutMode.TwoSidedMirror;
			}
		}
		catch
		{
		}
		DocumentPrintLayoutState documentPrintLayoutState = new DocumentPrintLayoutState();
		documentPrintLayoutState.Mode = mode;
		return documentPrintLayoutState;
	}

	private void RestoreDocumentPrintLayout(Document doc)
	{
		PrintLayoutMode printLayoutMode = ((doc != null) ? _docPrintLayouts.GetValue(doc, CreatePrintLayoutState).Mode : PrintLayoutMode.SingleSided);
		bool flag = _selectedPrintLayout != printLayoutMode;
		_selectedPrintLayout = printLayoutMode;
		_selectedPrintLayoutIndex = ((printLayoutMode == PrintLayoutMode.TwoSidedMirror) ? 1 : 0);
		SyncTaskPanePrintLayout(printLayoutMode);
		if (flag && _ribbon != null)
		{
			try
			{
				_ribbon.InvalidateControl("ddPrintLayout");
			}
			catch
			{
			}
		}
	}

	private void SelectDocumentPrintLayout(Document doc, PrintLayoutMode mode)
	{
		if (doc != null)
		{
			_docPrintLayouts.GetValue(doc, CreatePrintLayoutState).Mode = mode;
			RestoreDocumentPrintLayout(doc);
		}
	}

	private string GetDocumentKey(Document doc)
	{
		if (doc == null)
		{
			return null;
		}
		try
		{
			string fullName = doc.FullName;
			if (!string.IsNullOrEmpty(fullName))
			{
				return fullName;
			}
		}
		catch
		{
		}
		try
		{
			string name = doc.Name;
			if (!string.IsNullOrEmpty(name))
			{
				return name;
			}
		}
		catch
		{
		}
		return null;
	}

	private int GetRegulationIndexFromRulePack(RulePackType rulePack)
	{
		return rulePack switch
		{
			RulePackType.HD05_VPTW => 1, 
			RulePackType.TT68_2023_BQP => 2, 
			RulePackType.HD29_TWDTN => 3, 
			_ => 0, 
		};
	}

	private bool AutoDetectCurrentDocumentSilently(Document doc)
	{
		if (doc == null)
		{
			return false;
		}
		string documentKey = GetDocumentKey(doc);
		if (string.IsNullOrEmpty(documentKey))
		{
			return false;
		}
		string detectionSample = GetDetectionSample(doc);
		string text = CreateDetectionFingerprint(detectionSample);
		if (_manuallySelectedDocs.Contains(documentKey) && _docSelectedRules.TryGetValue(documentKey, out var value))
		{
			_selectedRulePack = value;
			_selectedRegulationIndex = GetRegulationIndexFromRulePack(value);
			SyncTaskPaneRulePack(value);
			if (_ribbon != null)
			{
				try
				{
					_ribbon.InvalidateControl("ddRegulation");
				}
				catch
				{
				}
			}
			return true;
		}
		if (_docDetectionFingerprints.TryGetValue(documentKey, out var value2) && string.Equals(value2, text, StringComparison.Ordinal) && _docSelectedRules.TryGetValue(documentKey, out value))
		{
			_selectedRulePack = value;
			_selectedRegulationIndex = GetRegulationIndexFromRulePack(value);
			SyncTaskPaneRulePack(value);
			if (_ribbon != null)
			{
				try
				{
					_ribbon.InvalidateControl("ddRegulation");
				}
				catch
				{
				}
			}
			return true;
		}
		try
		{
			RulePackDetectionResult rulePackDetectionResult = _engine.RuleEngine.DetectRulePackDetailed(detectionSample);
			RulePackType rulePackType = (_selectedRulePack = rulePackDetectionResult.RulePack);
			_selectedRegulationIndex = GetRegulationIndexFromRulePack(rulePackType);
			_autoDetectedDocs.Add(documentKey);
			_docSelectedRules[documentKey] = rulePackType;
			_docDetectionFingerprints[documentKey] = text;
			_docDetectionResults[documentKey] = rulePackDetectionResult;
			SyncTaskPaneRulePack(rulePackType);
			if (_ribbon != null)
			{
				try
				{
					_ribbon.InvalidateControl("ddRegulation");
				}
				catch
				{
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			Trace.WriteLine("AutoDetectCurrentDocumentSilently Error: " + ex.Message);
			return false;
		}
	}

	private string GetDetectionSample(Document doc)
	{
		if (doc == null)
		{
			return string.Empty;
		}
		try
		{
			int num = Math.Min(doc.Content.End, 6000);
			if (num <= 0)
			{
				return string.Empty;
			}
			object Start = 0;
			object End = num;
			Range range = doc.Range(ref Start, ref End);
			try
			{
				return range.Text ?? string.Empty;
			}
			finally
			{
				if (range != null)
				{
					Marshal.ReleaseComObject(range);
				}
			}
		}
		catch
		{
			return string.Empty;
		}
	}

	private static string CreateDetectionFingerprint(string sample)
	{
		if (sample == null)
		{
			return "0";
		}
		int num = 17;
		foreach (char c in sample)
		{
			num = num * 31 + c;
		}
		return sample.Length + ":" + num.ToString("X");
	}

	private bool IsAutomaticStandardizationSafe(Document doc)
	{
		string documentKey = GetDocumentKey(doc);
		if (string.IsNullOrEmpty(documentKey))
		{
			_selectedRulePack = RulePackType.ND30_2020;
			_selectedRegulationIndex = GetRegulationIndexFromRulePack(_selectedRulePack);
			return true;
		}
		if (_manuallySelectedDocs.Contains(documentKey))
		{
			return true;
		}
		if (!AutoDetectCurrentDocumentSilently(doc))
		{
			_selectedRulePack = RulePackType.ND30_2020;
			_selectedRegulationIndex = GetRegulationIndexFromRulePack(_selectedRulePack);
			_docSelectedRules[documentKey] = _selectedRulePack;
		}
		return true;
	}

	private void OnWordDocumentChange()
	{
		if (!_disconnecting && !_longOperationRunning && _wordApp != null)
		{
			try
			{
				RestoreDocumentPrintLayout((_wordApp.Documents.Count > 0) ? _wordApp.ActiveDocument : null);
			}
			catch (Exception error)
			{
				LogLifecycle("DocumentPrintLayout", error);
			}
		}
		ScheduleDocumentDetection();
	}

	private void ScheduleDocumentDetection()
	{
		if (_disconnecting || !_startupReady || _wordApp == null)
		{
			return;
		}
		try
		{
			if (_documentDetectionTimer == null)
			{
				_documentDetectionTimer = new Timer
				{
					Interval = 750
				};
				_documentDetectionTimer.Tick += OnDocumentDetectionTick;
			}
			_documentDetectionTimer.Stop();
			_documentDetectionTimer.Start();
		}
		catch (Exception error)
		{
			LogLifecycle("ScheduleDocumentDetection", error);
		}
	}

	private void OnDocumentDetectionTick(object sender, EventArgs args)
	{
		try
		{
			if (sender is Timer timer)
			{
				timer.Stop();
			}
			if (!_disconnecting && _startupReady)
			{
				ProcessWordDocumentChange();
			}
		}
		catch (Exception error)
		{
			LogLifecycle("DocumentDetectionTick", error);
		}
	}

	private void ProcessWordDocumentChange()
	{
		if (_longOperationRunning)
		{
			return;
		}
		try
		{
			Document document = null;
			try
			{
				if (_wordApp != null && _wordApp.Documents.Count > 0)
				{
					document = _wordApp.ActiveDocument;
				}
			}
			catch
			{
			}
			RestoreDocumentPrintLayout(document);
			if (document == null)
			{
				return;
			}
			string documentKey = GetDocumentKey(document);
			if (string.IsNullOrEmpty(documentKey))
			{
				return;
			}
			if (_docSelectedRules.TryGetValue(documentKey, out var value))
			{
				_selectedRulePack = value;
				_selectedRegulationIndex = GetRegulationIndexFromRulePack(value);
				SyncTaskPaneRulePack(value);
				if (_ribbon != null)
				{
					try
					{
						_ribbon.InvalidateControl("ddRegulation");
						return;
					}
					catch
					{
						return;
					}
				}
			}
			else
			{
				AutoDetectCurrentDocumentSilently(document);
			}
		}
		catch (Exception error)
		{
			LogLifecycle("OnWordDocumentChange", error);
		}
	}

	public void SetSuppressDialogs(bool suppress)
	{
		_suppressDialogs = suppress;
	}

	internal static void EnsureWpfApplication()
	{
		if (System.Windows.Application.Current == null)
		{
			try
			{
				System.Windows.Application application = new System.Windows.Application();
				application.ShutdownMode = ShutdownMode.OnExplicitShutdown;
			}
			catch
			{
			}
		}
		try
		{
			System.Windows.Application current = System.Windows.Application.Current;
			if (current == null)
			{
				return;
			}
			bool flag = false;
			foreach (ResourceDictionary mergedDictionary in current.Resources.MergedDictionaries)
			{
				if (mergedDictionary.Source != null && mergedDictionary.Source.OriginalString.Contains("FluentColors.xaml"))
				{
					flag = true;
					break;
				}
			}
			if (!flag)
			{
				current.Resources.MergedDictionaries.Add(new ResourceDictionary
				{
					Source = new Uri("pack://application:,,,/VFormat.UI;component/Styles/FluentColors.xaml", UriKind.Absolute)
				});
				current.Resources.MergedDictionaries.Add(new ResourceDictionary
				{
					Source = new Uri("pack://application:,,,/VFormat.UI;component/Styles/FluentControls.xaml", UriKind.Absolute)
				});
				current.Resources.MergedDictionaries.Add(new ResourceDictionary
				{
					Source = new Uri("pack://application:,,,/VFormat.UI;component/Styles/VectorIcons.xaml", UriKind.Absolute)
				});
			}
		}
		catch
		{
		}
	}

	public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
	{
		try
		{
			ClearWordResiliencyMarkers();
			_wordApp = (Microsoft.Office.Interop.Word.Application)Application;
			_disconnecting = false;
			_startupReady = ConnectMode != ext_ConnectMode.ext_cm_Startup;
			try
			{
				new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "DocumentChange").AddEventHandler(_wordApp, new ApplicationEvents4_DocumentChangeEventHandler(OnWordDocumentChange));
			}
			catch (Exception error)
			{
				LogLifecycle("SubscribeDocumentChange", error);
			}
			if (AddInInst is COMAddIn cOMAddIn)
			{
				cOMAddIn.Object = this;
			}
			if (_startupReady)
			{
				ScheduleDocumentDetection();
			}
		}
		catch (Exception error2)
		{
			LogLifecycle("OnConnection", error2);
		}
	}

	private static void ClearWordResiliencyMarkers()
	{
		try
		{
			using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Word\Resiliency\StartupItems", writable: true))
			{
				if (key != null)
				{
					foreach (string name in key.GetValueNames())
					{
						key.DeleteValue(name, throwOnMissingValue: false);
					}
				}
			}
			using (Microsoft.Win32.RegistryKey key2 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems", writable: true))
			{
				if (key2 != null)
				{
					foreach (string name2 in key2.GetValueNames())
					{
						key2.DeleteValue(name2, throwOnMissingValue: false);
					}
				}
			}
			using (Microsoft.Win32.RegistryKey key3 = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Word\Addins\VFormat.Addin", writable: true))
			{
				key3?.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
			}
		}
		catch
		{
		}
	}

	public void OnStartupComplete(ref Array custom)
	{
		try
		{
			_startupReady = true;
			ClearWordResiliencyMarkers();
			ScheduleDocumentDetection();
		}
		catch (Exception error)
		{
			LogLifecycle("OnStartupComplete", error);
		}
	}

	public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
	{
		_disconnecting = true;
		_startupReady = false;
		try
		{
			StopDocumentDetection();
			try
			{
				if (_wordApp != null)
				{
					new ComAwareEventInfo(typeof(ApplicationEvents4_Event), "DocumentChange").RemoveEventHandler(_wordApp, new ApplicationEvents4_DocumentChangeEventHandler(OnWordDocumentChange));
				}
			}
			catch
			{
			}
			if (_taskPane != null)
			{
				_taskPane.Delete();
			}
		}
		catch (Exception error)
		{
			LogLifecycle("OnDisconnection", error);
		}
		finally
		{
			_taskPane = null;
			_taskPaneControl = null;
			_ribbon = null;
			_ctpFactory = null;
			_wordApp = null;
			_engineInstance = null;
			_profileRepoInstance = null;
			_licenseManagerInstance = null;
		}
	}

	public void OnBeginShutdown(ref Array custom)
	{
		_disconnecting = true;
		_startupReady = false;
		StopDocumentDetection();
	}

	private void StopDocumentDetection()
	{
		Timer documentDetectionTimer = _documentDetectionTimer;
		_documentDetectionTimer = null;
		if (documentDetectionTimer == null)
		{
			return;
		}
		try
		{
			documentDetectionTimer.Stop();
			documentDetectionTimer.Tick -= OnDocumentDetectionTick;
			documentDetectionTimer.Dispose();
		}
		catch (Exception error)
		{
			LogLifecycle("StopDocumentDetection", error);
		}
	}

	public void OnAddInsUpdate(ref Array custom)
	{
	}

	public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
	{
		_ctpFactory = CTPFactoryInst;
	}

	private void EnsureTaskPane(bool makeVisible)
	{
		if (_taskPane == null && _ctpFactory != null)
		{
			try
			{
				_taskPane = _ctpFactory.CreateCTP("VFormat.Addin.TaskPaneControl", "OfficeFormat 1.0.0 - Chuẩn Hóa Văn Bản", Type.Missing);
				_taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
				_taskPane.Width = 360;
				_taskPaneControl = _taskPane.ContentControl as TaskPaneControl;
				if (_taskPaneControl != null)
				{
					_taskPaneControl.Initialize(_wordApp, ShowPersonalizationDialog, ShowLicenseDialog, ShowSpellCheckDialog, () =>
					{
						ShowTemplateLibraryDialog();
					});
					if (_taskPaneControl.ViewModel != null)
					{
						SyncTaskPaneRulePack(_selectedRulePack);
						SyncTaskPanePrintLayout(_selectedPrintLayout);
						_taskPaneControl.ViewModel.OnRulePackChangedAction = (RulePackType rule) =>
						{
							if (!_syncingTaskPane)
							{
								ApplyUserSelectedRulePack(rule);
								if (_ribbon != null)
								{
									try
									{
										_ribbon.InvalidateControl("ddRegulation");
									}
									catch
									{
									}
								}
							}
						};
						_taskPaneControl.ViewModel.OnPrintLayoutChangedAction = (PrintLayoutMode mode) =>
						{
							if (!_syncingTaskPane && EnsureActiveDocument(out var doc))
							{
								SelectDocumentPrintLayout(doc, mode);
							}
						};
						_taskPaneControl.ViewModel.StandardizeAllAction = (StandardizeOptions options) =>
						{
							RunAutoStandardize(options);
						};
						_taskPaneControl.ViewModel.StandardizeSelectionAction = TaskPane_StandardizeSelection;
						_taskPaneControl.ViewModel.InsertHeaderAction = () =>
						{
							Ribbon_InsertQuocHieuAuto(null);
						};
						_taskPaneControl.ViewModel.InsertSignerAction = () =>
						{
							Ribbon_InsertSigAuto(null);
						};
					}
				}
			}
			catch (Exception error)
			{
				LogLifecycle("EnsureTaskPane", error);
				try
				{
					if (_taskPane != null)
					{
						_taskPane.Delete();
					}
				}
				catch
				{
				}
				_taskPane = null;
				_taskPaneControl = null;
			}
		}
		if (_taskPane != null && makeVisible)
		{
			_taskPane.Visible = true;
		}
	}

	public string GetCustomUI(string RibbonID)
	{
		try
		{
			Assembly executingAssembly = Assembly.GetExecutingAssembly();
			string[] names = executingAssembly.GetManifestResourceNames();
			foreach (string name in names)
			{
				if (name.EndsWith("customUI14.xml", StringComparison.OrdinalIgnoreCase))
				{
					using (Stream stream = executingAssembly.GetManifestResourceStream(name))
					{
						if (stream != null)
						{
							using (StreamReader streamReader = new StreamReader(stream, Encoding.UTF8))
							{
								return streamReader.ReadToEnd();
							}
						}
					}
				}
			}
			string dir = Path.GetDirectoryName(executingAssembly.Location);
			string xmlPath = Path.Combine(dir, "VFormat.Addin.Ribbon.customUI14.xml");
			if (File.Exists(xmlPath))
			{
				return File.ReadAllText(xmlPath, Encoding.UTF8);
			}
		}
		catch (Exception error)
		{
			LogLifecycle("GetCustomUI", error);
		}
		return null;
	}

	public System.Drawing.Bitmap Ribbon_GetIcon(IRibbonControl control)
	{
		try
		{
			return RibbonIconRenderer.GetIcon(control?.Id);
		}
		catch
		{
			return null;
		}
	}
	public void Ribbon_OnLoad(IRibbonUI ribbon)
	{
		_ribbon = ribbon;
		ScheduleDocumentDetection();
	}

	private static void LogLifecycle(string stage, Exception error)
	{
		Trace.WriteLine(stage + ": " + error);
		try
		{
			string text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VFormat", "Logs");
			Directory.CreateDirectory(text);
			string text2 = Path.Combine(text, "addin-startup.log");
			if (File.Exists(text2) && new FileInfo(text2).Length > 1048576)
			{
				File.Delete(text2);
			}
			File.AppendAllText(text2, string.Concat(DateTime.UtcNow.ToString("o"), " ", stage, " ", error, Environment.NewLine));
		}
		catch
		{
		}
	}

	[DllImport("user32.dll")]
	private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

	[DllImport("user32.dll")]
	private static extern bool IsWindowEnabled(IntPtr hWnd);

	[DllImport("user32.dll")]
	private static extern bool SetForegroundWindow(IntPtr hWnd);

	private LongOperationOutcome RunLongOperation(string title, string initialStatus, Action work, out Exception error, int expectedSteps = 0)
	{
		error = null;
		if (_longOperationRunning)
		{
			return LongOperationOutcome.Busy;
		}
		_longOperationRunning = true;
		IntPtr wordHwnd = GetWordHwnd();
		bool flag = false;
		ProgressDialogHost progressDialogHost = null;
		try
		{
			if (!SuppressDialogs)
			{
				try
				{
					progressDialogHost = ProgressDialogHost.Start(title, initialStatus, wordHwnd, allowCancel: true);
				}
				catch
				{
					progressDialogHost = null;
				}
			}
			if (progressDialogHost != null && wordHwnd != IntPtr.Zero && IsWindowEnabled(wordHwnd))
			{
				EnableWindow(wordHwnd, bEnable: false);
				flag = true;
			}
			ProgressDialogHost progressHost = progressDialogHost;
			using (OperationProgress.Begin((int pct, string msg) =>
			{
				if (progressHost != null)
				{
					progressHost.Update(pct, msg);
				}
			}, () => progressHost != null && progressHost.IsCancelRequested))
			{
				using (new SafeUndoScope((Microsoft.Office.Interop.Word.Application)_wordApp, title))
				{
					if (expectedSteps > 0)
					{
						OperationProgress.Expect(expectedSteps);
					}
					try
					{
						work();
					}
					catch (OperationCanceledException)
					{
						return LongOperationOutcome.Cancelled;
					}
					catch (Exception ex2)
					{
						if (OperationProgress.WasCancelled)
						{
							return LongOperationOutcome.Cancelled;
						}
						error = ex2;
						return LongOperationOutcome.Failed;
					}
					return OperationProgress.WasCancelled ? LongOperationOutcome.Cancelled : LongOperationOutcome.Completed;
				}
			}
		}
		finally
		{
			if (flag)
			{
				try
				{
					EnableWindow(wordHwnd, bEnable: true);
				}
				catch
				{
				}
			}
			if (progressDialogHost != null)
			{
				try
				{
					progressDialogHost.Dispose();
				}
				catch
				{
				}
			}
			if (flag)
			{
				try
				{
					SetForegroundWindow(wordHwnd);
				}
				catch
				{
				}
			}
			_longOperationRunning = false;
		}
	}

	private bool RunLongOperationWithNotices(string title, string initialStatus, Action work, int expectedSteps = 0)
	{
		Exception error;
		switch (RunLongOperation(title, initialStatus, work, out error, expectedSteps))
		{
		case LongOperationOutcome.Completed:
			return true;
		case LongOperationOutcome.Cancelled:
			SetTaskPaneStatus("Đã dừng theo yêu cầu. Ctrl+Z để hoàn tác phần đã thay đổi.");
			ShowModernInfo("Đã dừng theo yêu cầu.\r\n\r\nPhần văn bản đã thay đổi trước khi dừng có thể khôi phục bằng Ctrl+Z (Hoàn tác) một lần.", title);
			return false;
		case LongOperationOutcome.Busy:
			ShowModernInfo("OfficeFormat đang xử lý một tác vụ khác. Vui lòng chờ tác vụ đó hoàn tất.", "OFFICEFORMAT");
			return false;
		default:
			ShowModernError(title + ": " + ((error != null) ? error.Message : "Lỗi không xác định."), "OFFICEFORMAT");
			return false;
		}
	}

	private IntPtr GetWordHwnd()
	{
		try
		{
			if (_wordApp != null && _wordApp.Windows.Count > 0)
			{
				int hwnd = _wordApp.ActiveWindow.Hwnd;
				if (hwnd != 0)
				{
					return new IntPtr(hwnd);
				}
			}
		}
		catch
		{
		}
		try
		{
			return Process.GetCurrentProcess().MainWindowHandle;
		}
		catch
		{
			return IntPtr.Zero;
		}
	}

	private void ShowModernInfo(string message, string title = "OfficeFormat 1.0.0")
	{
		if (!SuppressDialogs)
		{
			EnsureWpfApplication();
			ModernMessageDialog.ShowInfo(message, title, GetWordHwnd());
		}
	}

	private void ShowModernSuccess(string message, string title = "OfficeFormat 1.0.0")
	{
		if (!SuppressDialogs)
		{
			EnsureWpfApplication();
			ModernMessageDialog.ShowSuccess(message, title, GetWordHwnd());
		}
	}

	private void ShowModernWarning(string message, string title = "OfficeFormat 1.0.0")
	{
		if (!SuppressDialogs)
		{
			EnsureWpfApplication();
			ModernMessageDialog.ShowWarning(message, title, GetWordHwnd());
		}
	}

	private void ShowModernError(string message, string title = "OfficeFormat 1.0.0")
	{
		if (!SuppressDialogs)
		{
			EnsureWpfApplication();
			ModernMessageDialog.ShowError(message, title, GetWordHwnd());
		}
	}

	private bool ShowModernConfirm(string message, string title = "OfficeFormat 1.0.0", string okText = "Đồng ý", string cancelText = "Hủy Bỏ")
	{
		if (SuppressDialogs)
		{
			return SuppressedConfirmationResult;
		}
		EnsureWpfApplication();
		return ModernMessageDialog.ShowConfirm(message, title, okText, cancelText, GetWordHwnd());
	}

	private bool EnsureLicenseActive()
	{
		string message = string.Empty;
		if (_licenseManager != null && _licenseManager.IsLicenseActive(out message))
		{
			return true;
		}
		if (!SuppressDialogs)
		{
			if (IsNotActivatedYet())
			{
				LicenseRequiredChoice licenseRequiredChoice = LicenseRequiredChoice.Close;
				try
				{
					EnsureWpfApplication();
					licenseRequiredChoice = LicenseRequiredDialog.Ask(GetWordHwnd());
				}
				catch (Exception ex)
				{
					Trace.WriteLine("License required dialog: " + ex.GetType().Name);
				}
				switch (licenseRequiredChoice)
				{
				case LicenseRequiredChoice.EnterKey:
					ShowLicenseDialog();
					if (_licenseManager != null)
					{
						return _licenseManager.IsLicenseActive(out message);
					}
					return false;
				case LicenseRequiredChoice.Buy:
					try
					{
						ExternalLinkHelper.OpenBuyLicensePage(null, GetWordHwnd());
					}
					catch (Exception ex2)
					{
						Trace.WriteLine("Open buy page: " + ex2.GetType().Name);
					}
					break;
				}
				return false;
			}
			ShowModernWarning((!string.IsNullOrEmpty(message)) ? message : "Vui lòng kích hoạt bản quyền OfficeFormat để sử dụng tính năng này!", "OfficeFormat 1.0.0 - Yêu Cầu Bản Quyền");
			ShowLicenseDialog();
		}
		if (_licenseManager != null)
		{
			return _licenseManager.IsLicenseActive(out message);
		}
		return false;
	}

	private bool IsNotActivatedYet()
	{
		try
		{
			LicenseCheckResult licenseCheckResult = ((_licenseManager != null) ? _licenseManager.CurrentStatus : null);
			return licenseCheckResult != null && licenseCheckResult.Status == LicenseStatus.NoLicense && string.Equals(licenseCheckResult.StatusCodeString, "NO_LICENSE", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(_licenseManager.GetSavedLicenseKey());
		}
		catch
		{
			return false;
		}
	}

	private bool EnsureActiveDocument(out Document doc)
	{
		doc = null;
		if (_wordApp == null)
		{
			try
			{
				_wordApp = (Microsoft.Office.Interop.Word.Application)Marshal.GetActiveObject("Word.Application");
			}
			catch
			{
			}
		}
		if (_wordApp == null || _wordApp.Documents.Count == 0)
		{
			ShowModernInfo("Vui lòng mở một tài liệu Word trước khi thực hiện thao tác.");
			return false;
		}
		try
		{
			doc = _wordApp.ActiveDocument;
		}
		catch
		{
			doc = null;
		}
		if (doc != null)
		{
			RestoreDocumentPrintLayout(doc);
		}
		return doc != null;
	}

	public void OnToggleTaskPaneClicked(IRibbonControl control, bool isPressed)
	{
		try
		{
			if (_taskPane == null)
			{
				EnsureTaskPane(isPressed);
			}
			else
			{
				_taskPane.Visible = isPressed;
			}
			if (_ribbon != null)
			{
				_ribbon.InvalidateControl("btnToggleTaskPane");
			}
		}
		catch (Exception ex)
		{
			Trace.WriteLine("OnToggleTaskPaneClicked Error: " + ex.Message);
		}
	}

	public void OnToggleTaskPaneClicked(IRibbonControl control)
	{
		try
		{
			if (_taskPane == null)
			{
				EnsureTaskPane(makeVisible: true);
			}
			else
			{
				_taskPane.Visible = !_taskPane.Visible;
			}
			if (_ribbon != null)
			{
				_ribbon.InvalidateControl("btnToggleTaskPane");
			}
		}
		catch (Exception ex)
		{
			Trace.WriteLine("OnToggleTaskPaneClicked Error: " + ex.Message);
		}
	}

	public bool GetTaskPanePressed(IRibbonControl control)
	{
		if (_taskPane != null)
		{
			return _taskPane.Visible;
		}
		return false;
	}

	public void Ribbon_AutoStandardize(IRibbonControl control)
	{
		RunAutoStandardize(null);
	}

	private void RunAutoStandardize(StandardizeOptions options)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc) || !IsAutomaticStandardizationSafe(doc))
		{
			return;
		}
		string documentKey = GetDocumentKey(doc);
		if (!string.IsNullOrEmpty(documentKey) && !_autoDetectedDocs.Contains(documentKey))
		{
			AutoDetectCurrentDocumentSilently(doc);
		}
		Selection sel = null;
		try
		{
			sel = _wordApp?.Selection;
		}
		catch
		{
		}
		bool hasSelection = false;
		try
		{
			hasSelection = sel != null && sel.Type != WdSelectionType.wdSelectionIP && sel.Start != sel.End && !string.IsNullOrWhiteSpace(sel.Text);
		}
		catch
		{
		}
		if (hasSelection)
		{
			FormattingResult selectedResult = null;
			if (!RunLongOperationWithNotices("OFFICEFORMAT - Chuẩn Hóa Vùng Chọn", "Đang chuẩn hóa theo quy chuẩn cho vùng chọn...", () =>
			{
				selectedResult = _engine.StandardizeSelection((Selection)sel, _selectedRulePack, includeAllRoles: true);
			}))
			{
				return;
			}
			if (!selectedResult.Success)
			{
				ShowModernError("Không thể chuẩn hóa vùng chọn: " + selectedResult.Message, "OFFICEFORMAT");
				return;
			}
			if (selectedResult.ParagraphsProcessed == 0 && selectedResult.TablesProcessed == 0)
			{
				ShowModernInfo("Vùng chọn không chứa nội dung phù hợp để chuẩn hóa.", "OFFICEFORMAT");
				return;
			}
			SetTaskPaneStatus(selectedResult.Message);
			try
			{
				if (_wordApp != null)
				{
					_wordApp.StatusBar = $"OFFICEFORMAT: Đã chuẩn hóa vùng chọn ({selectedResult.ParagraphsProcessed} đoạn văn, {selectedResult.TablesProcessed} bảng) thành công!";
				}
			}
			catch
			{
			}
			return;
		}
		RulePackDetectionResult value = null;
		if (!string.IsNullOrEmpty(documentKey))
		{
			_docDetectionResults.TryGetValue(documentKey, out value);
		}
		bool safeMode = value != null && !value.IsConfident && (string.IsNullOrEmpty(documentKey) || !_manuallySelectedDocs.Contains(documentKey));
		FormattingResult res = null;
		if (!RunLongOperationWithNotices("OFFICEFORMAT - Chuẩn Hóa Tự Động", "Đang khởi tạo và nhận diện quy chuẩn...", () =>
		{
			res = _engine.StandardizeDocument((Document)doc, _selectedRulePack, _selectedPrintLayout, safeMode, options);
		}))
		{
			return;
		}
		SetTaskPaneStatus(res.Message);
		if (!res.Success)
		{
			if (res.Error == null)
			{
				ShowModernInfo(res.Message, "OFFICEFORMAT - Chuẩn hóa tự động");
			}
			else
			{
				ShowModernError("Không thể chuẩn hóa tự động: " + res.Message, "OFFICEFORMAT");
			}
			return;
		}
		string text = (res.SafeModeApplied ? " Chế độ thận trọng: không tái tạo từ dữ liệu thiếu neo và không chèn khối mẫu." : string.Empty);
		try
		{
			if (_wordApp != null)
			{
				_wordApp.StatusBar = $"OFFICEFORMAT: Đã chuẩn hóa {res.ParagraphsProcessed} đoạn, {res.TablesProcessed} bảng. {res.HeaderStatus} {res.SignatureStatus}{text}";
			}
		}
		catch
		{
		}
	}

	public void Ribbon_DetectDocument(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			string detectionSample = GetDetectionSample(doc);
			RulePackDetectionResult rulePackDetectionResult = _engine.RuleEngine.DetectRulePackDetailed(detectionSample);
			RulePackType rulePackType = (_selectedRulePack = rulePackDetectionResult.RulePack);
			_selectedRegulationIndex = GetRegulationIndexFromRulePack(rulePackType);
			string documentKey = GetDocumentKey(doc);
			if (!string.IsNullOrEmpty(documentKey))
			{
				_manuallySelectedDocs.Remove(documentKey);
				_autoDetectedDocs.Add(documentKey);
				_docSelectedRules[documentKey] = rulePackType;
				_docDetectionFingerprints[documentKey] = CreateDetectionFingerprint(detectionSample);
				_docDetectionResults[documentKey] = rulePackDetectionResult;
			}
			SyncTaskPaneRulePack(rulePackType);
			if (_ribbon != null)
			{
				try
				{
					_ribbon.InvalidateControl("ddRegulation");
				}
				catch
				{
				}
			}
			string text = null;
			string text2 = null;
			if (!string.IsNullOrEmpty(detectionSample))
			{
				string[] array = detectionSample.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (string text3 in array)
				{
					string text4 = text3.Trim();
					if (text2 == null && (RegexPatterns.MilitaryNumber.IsMatch(text4) || RegexPatterns.AdminNumber.IsMatch(text4) || RegexPatterns.PartyNumber.IsMatch(text4) || RegexPatterns.DocumentNumberGeneral.IsMatch(text4)))
					{
						text2 = text4;
					}
					else if (text == null && (RegexPatterns.YouthUnionAgency.IsMatch(text4) || RegexPatterns.MilitaryAgency.IsMatch(text4) || RegexPatterns.AdminAgency.IsMatch(text4) || RegexPatterns.PartyAgency.IsMatch(text4)))
					{
						text = text4;
					}
				}
			}
			if (!SuppressDialogs)
			{
				EnsureWpfApplication();
				RegulationDetectDialog.ShowDetectResult(rulePackType, text, text2, GetWordHwnd());
				if (!rulePackDetectionResult.IsConfident)
				{
					ShowModernInfo("Kết quả chỉ có tính gợi ý: " + rulePackDetectionResult.Summary + "\r\nHãy chọn lại quy chuẩn trước khi chuẩn hóa tự động.", "OFFICEFORMAT - Nhận diện cần xác nhận");
				}
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi nhận diện văn bản: " + ex.Message);
		}
	}

	public int Ribbon_GetDefaultRegulation(IRibbonControl control)
	{
		try
		{
			Document document = null;
			try
			{
				if (_wordApp != null && _wordApp.Documents.Count > 0)
				{
					document = _wordApp.ActiveDocument;
				}
			}
			catch
			{
			}
			if (document != null)
			{
				string documentKey = GetDocumentKey(document);
				if (!string.IsNullOrEmpty(documentKey))
				{
					RulePackType value;
					if (!_autoDetectedDocs.Contains(documentKey))
					{
						AutoDetectCurrentDocumentSilently(document);
					}
					else if (_docSelectedRules.TryGetValue(documentKey, out value))
					{
						_selectedRulePack = value;
						_selectedRegulationIndex = GetRegulationIndexFromRulePack(value);
					}
				}
			}
		}
		catch
		{
		}
		return _selectedRegulationIndex;
	}

	public void Ribbon_SelectRegulation(IRibbonControl control, string selectedId, int selectedIndex)
	{
		string text = selectedId ?? string.Empty;
		string key = text;
		RulePackType rule;
		if (text != null)
		{
			if (_regDict == null)
			{
				_regDict = new Dictionary<string, int>(6)
				{
					{ "regDANG", 0 },
					{ "regTT68", 1 },
					{ "regDOAN", 2 },
					{ "regND30", 3 },
					{ "regVBHN", 4 },
					{ "regKHBD", 5 }
				};
			}
			if (_regDict.TryGetValue(key, out var value))
			{
				switch (value)
				{
				case 0:
					break;
				case 1:
					goto IL_00a4;
				case 2:
					goto IL_00a8;
				case 3:
				case 4:
				case 5:
					goto IL_00ac;
				default:
					goto IL_00b0;
				}
				rule = RulePackType.HD05_VPTW;
				goto IL_00d6;
			}
		}
		goto IL_00b0;
		IL_00d6:
		ApplyUserSelectedRulePack(rule);
		SyncTaskPaneRulePack(rule);
		return;
		IL_00b0:
		rule = selectedIndex switch
		{
			1 => RulePackType.HD05_VPTW, 
			2 => RulePackType.TT68_2023_BQP, 
			3 => RulePackType.HD29_TWDTN, 
			_ => RulePackType.ND30_2020, 
		};
		goto IL_00d6;
		IL_00a4:
		rule = RulePackType.TT68_2023_BQP;
		goto IL_00d6;
		IL_00ac:
		rule = RulePackType.ND30_2020;
		goto IL_00d6;
		IL_00a8:
		rule = RulePackType.HD29_TWDTN;
		goto IL_00d6;
	}

	private void ApplyUserSelectedRulePack(RulePackType rule)
	{
		_selectedRulePack = rule;
		_selectedRegulationIndex = GetRegulationIndexFromRulePack(rule);
		try
		{
			Document doc = null;
			try
			{
				if (_wordApp != null && _wordApp.Documents.Count > 0)
				{
					doc = _wordApp.ActiveDocument;
				}
			}
			catch
			{
			}
			string documentKey = GetDocumentKey(doc);
			if (!string.IsNullOrEmpty(documentKey))
			{
				_docSelectedRules[documentKey] = rule;
				_autoDetectedDocs.Add(documentKey);
				_manuallySelectedDocs.Add(documentKey);
			}
		}
		catch
		{
		}
	}

	private void SyncTaskPaneRulePack(RulePackType rule)
	{
		if (_taskPaneControl == null || _taskPaneControl.ViewModel == null)
		{
			return;
		}
		_syncingTaskPane = true;
		try
		{
			_taskPaneControl.ViewModel.SelectedRulePack = rule;
		}
		finally
		{
			_syncingTaskPane = false;
		}
	}

	private void SyncTaskPanePrintLayout(PrintLayoutMode mode)
	{
		if (_taskPaneControl == null || _taskPaneControl.ViewModel == null)
		{
			return;
		}
		_syncingTaskPane = true;
		try
		{
			_taskPaneControl.ViewModel.SelectedPrintLayout = mode;
		}
		finally
		{
			_syncingTaskPane = false;
		}
	}

	private void SetTaskPaneStatus(string message)
	{
		try
		{
			if (_taskPaneControl != null && _taskPaneControl.ViewModel != null)
			{
				_taskPaneControl.ViewModel.StatusMessage = message;
			}
		}
		catch
		{
		}
	}

	public void TaskPane_StandardizeSelection()
	{
		bool flag = false;
		try
		{
			if (_wordApp != null && _wordApp.Documents.Count > 0)
			{
				Selection selection = _wordApp.Selection;
				flag = selection != null && selection.Type == WdSelectionType.wdSelectionNormal && selection.Start != selection.End;
			}
		}
		catch
		{
		}
		if (!flag)
		{
			SetTaskPaneStatus("Hãy bôi đen vùng văn bản cần chuẩn hóa.");
			if (_wordApp != null)
			{
				try
				{
					if (_wordApp.Documents.Count > 0)
					{
						ShowModernInfo("Hãy bôi đen vùng văn bản cần chuẩn hóa rồi bấm lại.", "OFFICEFORMAT");
						return;
					}
				}
				catch
				{
				}
			}
		}
		Ribbon_ManualFormatBody(null);
	}

	public int Ribbon_GetDefaultOrgLevel(IRibbonControl control)
	{
		return _selectedOrgLevelIndex;
	}

	public void Ribbon_SelectOrgLevel(IRibbonControl control, string selectedId, int selectedIndex)
	{
		_selectedOrgLevelIndex = selectedIndex;
	}

	public int Ribbon_GetDefaultDocType(IRibbonControl control)
	{
		return _selectedDocTypeIndex;
	}

	public void Ribbon_SelectDocType(IRibbonControl control, string selectedId, int selectedIndex)
	{
		_selectedDocTypeIndex = selectedIndex;
	}

	public int Ribbon_GetDefaultPrintLayout(IRibbonControl control)
	{
		try
		{
			RestoreDocumentPrintLayout((_wordApp != null && _wordApp.Documents.Count > 0) ? _wordApp.ActiveDocument : null);
		}
		catch (Exception error)
		{
			LogLifecycle("GetPrintLayout", error);
		}
		return _selectedPrintLayoutIndex;
	}

	public void Ribbon_SelectPrintLayout(IRibbonControl control, string selectedId, int selectedIndex)
	{
		if (EnsureActiveDocument(out var doc))
		{
			SelectDocumentPrintLayout(doc, (selectedIndex == 1) ? PrintLayoutMode.TwoSidedMirror : PrintLayoutMode.SingleSided);
		}
	}

	public void Ribbon_AuditDocument(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			string report = null;
			int issues = 0;
			int incomplete = 0;
			if (RunLongOperationWithNotices("OFFICEFORMAT - Rà Soát Thể Thức", "Đang rà soát thể thức văn bản theo quy chuẩn...", () =>
			{
				report = WordToolEngine.AuditDocument((Document)doc, _selectedRulePack, out issues, out incomplete);
			}) && !SuppressDialogs)
			{
				EnsureWpfApplication();
				AuditReportDialog.ShowReport(report, issues, GetWordHwnd(), incomplete);
			}
		}
	}

	public void Ribbon_ConvertOldFontsToUnicode(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc) && RunLongOperationWithNotices("OFFICEFORMAT - Chuyển Đổi Phông Chữ", "Đang quét các bảng mã TCVN3, VNI và chuyển đổi sang Times New Roman...", () =>
		{
			WordToolEngine.ConvertOldFontsToUnicode((Document)doc);
		}))
		{
			ShowModernSuccess("Đã chuyển đổi toàn bộ phông chữ sang chuẩn Times New Roman (Unicode) thành công!", "OFFICEFORMAT");
		}
	}

	public void Ribbon_DeleteAuditComments(IRibbonControl control)
	{
		DeleteDocumentComments(onlyVFormatAudit: true);
	}

	public void Ribbon_DeleteAllComments(IRibbonControl control)
	{
		DeleteDocumentComments(onlyVFormatAudit: false);
	}

	private void DeleteDocumentComments(bool onlyVFormatAudit)
	{
		if (!EnsureActiveDocument(out var doc))
		{
			return;
		}
		string title = (onlyVFormatAudit ? "Xóa nhận xét rà soát OfficeFormat" : "Xóa tất cả nhận xét");
		try
		{
			if (doc.Comments.Count == 0)
			{
				ShowModernInfo("Tài liệu không có nhận xét nào để xóa.", title);
			}
			else
			{
				if (!onlyVFormatAudit && !ShowModernConfirm($"Xóa tất cả {doc.Comments.Count} nhận xét trong tài liệu này, gồm cả nhận xét do người dùng thêm?\r\n\r\nCó thể nhấn Ctrl+Z để hoàn tác sau khi xóa.", title, "Xóa tất cả", "Giữ lại"))
				{
					return;
				}
				int removed = 0;
				int failed = 0;
				int keptWithReplies = 0;
				if (RunLongOperationWithNotices(title, "Đang xóa nhận xét...", () =>
				{
					removed = WordToolEngine.DeleteDocumentComments((Document)doc, onlyVFormatAudit, out failed, out keptWithReplies);
				}))
				{
					string text = ((keptWithReplies > 0) ? $"\r\nGiữ lại {keptWithReplies} nhận xét rà soát vì có phần trả lời của người dùng; hãy xóa thủ công nếu không cần." : "");
					if (failed > 0)
					{
						ShowModernWarning($"Đã xóa {removed} nhận xét. Có {failed} nhận xét chưa xóa được.\r\nCó thể nhấn Ctrl+Z để hoàn tác phần đã xóa." + text, title);
					}
					else if (removed == 0)
					{
						ShowModernInfo(((keptWithReplies > 0) ? "Chưa xóa nhận xét rà soát nào." : "Không có nhận xét rà soát OfficeFormat nào để xóa.") + text, title);
					}
					else
					{
						ShowModernSuccess($"Đã xóa {removed} nhận xét. Có thể nhấn Ctrl+Z để hoàn tác." + text, title);
					}
				}
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể xóa nhận xét: " + ex.Message, title);
		}
	}

	public void Ribbon_UndoStandardize(IRibbonControl control)
	{
		if (!EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			Document document = doc;
			object Times = Missing.Value;
			document.Undo(ref Times);
		}
		catch
		{
			ShowModernInfo("Không có thao tác nào để hoàn tác.");
		}
	}

	private void RunManualComponent(string caption, Func<Document, bool, FormattingResult> run, string successMessage, string templateQuestion)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			AutoDetectCurrentDocumentSilently(doc);
		}
		catch
		{
		}
		try
		{
			FormattingResult res = null;
			if (RunLongOperationWithNotices(caption, "Đang nhận diện và chuẩn hóa thành phần văn bản...", () =>
			{
				res = run(doc, arg2: false);
			}) && (res.Success || res.Error != null || !res.CanInsertTemplate || templateQuestion == null || SuppressDialogs || (ShowModernConfirm(res.Message + "\r\n\r\n" + templateQuestion, caption, "Chèn khối mẫu", "Giữ nguyên") && RunLongOperationWithNotices(caption, "Đang chèn khối mẫu theo quy chuẩn...", () =>
			{
				res = run(doc, arg2: true);
			}))))
			{
				if (_taskPaneControl != null && _taskPaneControl.ViewModel != null)
				{
					_taskPaneControl.ViewModel.StatusMessage = res.Message;
				}
				if (res.Success)
				{
					string text;
					if (string.IsNullOrEmpty(res.HeaderStatus))
					{
						text = ((!string.IsNullOrEmpty(res.SignatureStatus)) ? res.SignatureStatus : null);
					}
					else
					{
						text = res.HeaderStatus;
					}
					ShowModernSuccess((text != null) ? (successMessage + "\r\n" + text) : successMessage, VersionLabel);
				}
				else if (res.Error != null)
				{
					ShowModernError(caption + ": " + res.Message, VersionLabel);
				}
				else
				{
					ShowModernInfo(res.Message, caption);
				}
			}
		}
		catch (Exception ex)
		{
			ShowModernError(caption + ": " + ex.Message, VersionLabel);
		}
	}

	public void Ribbon_ManualFormatHeader(IRibbonControl control)
	{
		RunManualComponent("OFFICEFORMAT - Tiêu đề Đầu trang", (Document d, bool allowInsert) => _engine.ManualFormatHeader((Document)d, _selectedRulePack, allowInsert), "Đã chuẩn hóa Khung Tiêu đề & Quốc hiệu thành công!", "Bạn có muốn chèn khối Tiêu đề mẫu theo quy chuẩn đang chọn để điền thông tin không?");
	}

	public void Ribbon_ManualFormatTitle(IRibbonControl control)
	{
		RunManualComponent("OFFICEFORMAT - Tên loại & Trích yếu", (Document d, bool allowInsert) => _engine.ManualFormatTitle((Document)d, _selectedRulePack), "Đã chuẩn hóa Tên loại văn bản & Trích yếu thành công!", null);
	}

	public void Ribbon_ManualFormatBody(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			AutoDetectCurrentDocumentSilently(doc);
		}
		catch
		{
		}
		Selection sel = null;
		try
		{
			sel = _wordApp.Selection;
		}
		catch
		{
		}
		bool flag = false;
		try
		{
			flag = sel != null && sel.Type == WdSelectionType.wdSelectionNormal;
		}
		catch
		{
		}
		if (flag)
		{
			FormattingResult selectedResult = null;
			if (!RunLongOperationWithNotices("OFFICEFORMAT - Chuẩn Hóa Vùng Chọn", "Đang chuẩn hóa các đoạn trong vùng chọn...", () =>
			{
				selectedResult = _engine.StandardizeSelection((Selection)sel, _selectedRulePack, includeAllRoles: false);
			}))
			{
				return;
			}
			if (!selectedResult.Success)
			{
				ShowModernError("Không thể chuẩn hóa vùng chọn: " + selectedResult.Message, "OFFICEFORMAT");
				return;
			}
			if (selectedResult.ParagraphsProcessed == 0)
			{
				ShowModernInfo("Vùng chọn không chứa đoạn thân bài phù hợp. Các thành phần thể thức, bảng biểu, đầu/cuối trang được giữ nguyên để tránh áp dụng nhầm.", "OFFICEFORMAT");
				return;
			}
			try
			{
				_wordApp.StatusBar = $"OfficeFormat: Đã chuẩn hóa vùng chọn ({selectedResult.ParagraphsProcessed} đoạn) thành công!";
			}
			catch
			{
			}
			SetTaskPaneStatus("Đã chuẩn hóa vùng chọn.");
			return;
		}
		FormattingResult res = null;
		if (!RunLongOperationWithNotices("OFFICEFORMAT - Chuẩn Hóa Đoạn Văn", "Đang căn đều 2 lề, thụt lề đầu dòng và chuẩn hóa giãn đoạn...", () =>
		{
			res = _engine.ManualFormatBody((Document)doc, _selectedRulePack);
		}))
		{
			return;
		}
		if (!res.Success)
		{
			if (res.Error == null)
			{
				ShowModernInfo(res.Message, "OFFICEFORMAT - Đề mục & Nội dung");
			}
			else
			{
				ShowModernError("Lỗi chuẩn hóa thân bài: " + res.Message, "OFFICEFORMAT");
			}
			return;
		}
		try
		{
			_wordApp.StatusBar = $"OfficeFormat: Đã chuẩn hóa Đề mục & Thân bài thành công ({res.ParagraphsProcessed} đoạn)! Căn đều 2 lề, thụt dòng 1cm.";
		}
		catch
		{
		}
		SetTaskPaneStatus($"Đã chuẩn hóa {res.ParagraphsProcessed} đoạn thân bài.");
	}

	public void Ribbon_ManualFormatSignature(IRibbonControl control)
	{
		RunManualComponent("OFFICEFORMAT - Chữ ký & Nơi nhận", (Document d, bool allowInsert) => _engine.ManualFormatSignature((Document)d, _selectedRulePack, allowInsert), "Đã chuẩn hóa Khung Chữ ký & Nơi nhận thành công!", "Bạn có muốn chèn khối Nơi nhận & Chữ ký mẫu theo quy chuẩn đang chọn để điền thông tin không?");
	}

	public void Ribbon_PageSetup(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			if (RunLongOperationWithNotices("OFFICEFORMAT - Căn Lề Trang A4", "Đang thiết lập khổ A4 và lề trang theo quy chuẩn...", () =>
			{
				WordToolEngine.ApplyPageSetupA4((Document)doc, _selectedRulePack, _selectedPrintLayout);
			}))
			{
				string shortName = RulePackTraits.GetShortName(_selectedRulePack);
				string arg = ((_selectedPrintLayout == PrintLayoutMode.TwoSidedMirror) ? "In 2 mặt đối xứng (Mirror Margins lề gáy mặt trước - mặt sau)" : "Gửi file điện tử (1 mặt - Lề đơn phẳng chuẩn)");
				string message = $"Đã chuẩn hóa căn lề trang A4 thành công!\n\n- Quy chuẩn: {shortName}\n- Bố cục: {arg}";
				ShowModernSuccess(message, "OFFICEFORMAT - Căn Lề Trang A4");
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi khi căn lề trang: " + ex.Message, "OFFICEFORMAT");
		}
	}

	public void Ribbon_OpenPageMarginDialog(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			EnsureWpfApplication();
			PageMarginDialog pageMarginDialog = new PageMarginDialog(_selectedRulePack, _selectedPrintLayout);
			try
			{
				WdOrientation wdOrientation = WdOrientation.wdOrientPortrait;
				if (_wordApp != null && _wordApp.Selection != null && _wordApp.Selection.PageSetup != null)
				{
					wdOrientation = _wordApp.Selection.PageSetup.Orientation;
				}
				else if (doc != null && doc.PageSetup != null)
				{
					wdOrientation = doc.PageSetup.Orientation;
				}
				pageMarginDialog.ViewModel.IsPortrait = wdOrientation != WdOrientation.wdOrientLandscape;
			}
			catch
			{
			}
			pageMarginDialog.OnApplyMargins = (float top, float bot, float left, float right, float hdr, float ftr, bool isLand, bool curSec, bool mirror) =>
			{
				if (RunLongOperationWithNotices("OFFICEFORMAT - Căn Lề Trang", "Đang áp dụng lề trang...", () =>
				{
					WordToolEngine.ApplyCustomPageMargins((Document)doc, top, bot, left, right, hdr, ftr, isLand, curSec, mirror);
				}))
				{
					SelectDocumentPrintLayout(doc, mirror ? PrintLayoutMode.TwoSidedMirror : PrintLayoutMode.SingleSided);
					ShowModernSuccess("Đã thiết lập và căn lề trang A4 thành công!", "OFFICEFORMAT");
				}
			};
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(pageMarginDialog).Owner = mainWindowHandle;
			}
			pageMarginDialog.ShowDialog();
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi khi mở hộp thoại căn lề trang: " + ex.Message, "OFFICEFORMAT");
		}
	}

	public void Ribbon_InsertLandscapePage(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertLandscapePage((Microsoft.Office.Interop.Word.Application)_wordApp);
		}
	}

	public void Ribbon_InsertPortraitPage(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertPortraitPage((Microsoft.Office.Interop.Word.Application)_wordApp);
		}
	}

	public void Ribbon_DeleteBlankPages(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			int deleted = 0;
			if (RunLongOperationWithNotices("OFFICEFORMAT - Xóa Trang Trắng", "Đang tìm và xóa trang/dòng trống thừa cuối văn bản...", () =>
			{
				deleted = WordToolEngine.DeleteBlankPages((Document)doc);
			}))
			{
				ShowModernInfo((deleted > 0) ? $"Đã xóa {deleted} trang/dòng trống thừa vô nghĩa!" : "Không tìm thấy trang trắng thừa nào.", "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_InsertPageNumbers(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		try
		{
			EnsureWpfApplication();
			PageNumberDialog pageNumberDialog = new PageNumberDialog();
			pageNumberDialog.OnApplyPageNumbering = (int fromPg, int startNum, bool hideFirst, bool isHdr) =>
			{
				if (RunLongOperationWithNotices("OFFICEFORMAT - Đánh Số Trang", "Đang đánh số trang theo quy chuẩn...", () =>
				{
					WordToolEngine.ApplyCustomPageNumbering((Document)doc, fromPg, startNum, hideFirst, isHdr);
				}))
				{
					string message = ((fromPg != 1) ? $"Đã đánh số trang: Từ trang {fromPg} (bắt đầu là số {startNum}) đến hết văn bản!" : (hideFirst ? "Đã đánh số trang chuẩn Quy định (Căn giữa lề trên, Times New Roman, tự động ẩn trang 1)!" : $"Đã đánh số trang: Từ trang 1 (bắt đầu là số {startNum})!"));
					ShowModernSuccess(message, "OFFICEFORMAT");
				}
			};
			pageNumberDialog.OnRemovePageNumbers = () =>
			{
				if (RunLongOperationWithNotices("OFFICEFORMAT - Xóa Số Trang", "Đang xóa số trang trong tài liệu...", () =>
				{
					WordToolEngine.RemoveAllPageNumbers((Document)doc);
				}))
				{
					ShowModernSuccess("Đã xóa sạch toàn bộ số trang trong tài liệu!", "OFFICEFORMAT");
				}
			};
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(pageNumberDialog).Owner = mainWindowHandle;
			}
			pageNumberDialog.ShowDialog();
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi khi mở hộp thoại đánh số trang: " + ex.Message, "OFFICEFORMAT");
		}
	}

	#region In an, PDF & Ma QR

	public void Ribbon_PrintOddPages(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				PrintEngine.PrintOddPages((Document)doc);
				ShowModernSuccess("Đã gửi lệnh in toàn bộ TRANG LẺ đến máy in!", "OFFICEFORMAT - In Ấn");
			}
			catch (Exception ex)
			{
				ShowModernError($"Lỗi khi in trang lẻ: {ex.Message}", "Lỗi In Ấn");
			}
		}
	}

	public void Ribbon_PrintEvenPages(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				PrintEngine.PrintEvenPages((Document)doc);
				ShowModernSuccess("Đã gửi lệnh in toàn bộ TRANG CHẴN đến máy in!", "OFFICEFORMAT - In Ấn");
			}
			catch (Exception ex)
			{
				ShowModernError($"Lỗi khi in trang chẵn: {ex.Message}", "Lỗi In Ấn");
			}
		}
	}

	public void Ribbon_OpenPrintDialog(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				EnsureWpfApplication();
				var dlg = new PrintDuplexDialog((Document)doc);
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở hộp thoại in ấn: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_OpenLatPdf(IRibbonControl control)
	{
		try
		{
			EnsureWpfApplication();
			var dlg = new LatPdfDialog();
			IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
			if (handle != IntPtr.Zero)
			{
				new WindowInteropHelper(dlg).Owner = handle;
			}
			dlg.ShowDialog();
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi khi mở công cụ Lật & Sắp xếp PDF: " + ex.Message, "OFFICEFORMAT");
		}
	}

	public void Ribbon_ConvertPdfToWord(IRibbonControl control)
	{
		if (EnsureLicenseActive())
		{
			try
			{
				EnsureWpfApplication();
				var dlg = new PdfConverterDialog((Microsoft.Office.Interop.Word.Application)_wordApp);
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở công cụ chuyển đổi PDF: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_ConvertPdfScanOcr(IRibbonControl control)
	{
		if (EnsureLicenseActive())
		{
			try
			{
				EnsureWpfApplication();
				var dlg = new PdfConverterDialog((Microsoft.Office.Interop.Word.Application)_wordApp);
				dlg.EnableOcr = true;
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở công cụ PDF Scan OCR: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_BatchPdfToWord(IRibbonControl control)
	{
		Ribbon_ConvertPdfToWord(control);
	}

	public void Ribbon_ExportCurrentToPdf(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			var d = (Document)doc;
			string docPath = "";
			try { docPath = d.FullName; } catch { }
			string defaultPdf = !string.IsNullOrEmpty(docPath) && File.Exists(docPath)
				? Path.ChangeExtension(docPath, ".pdf")
				: Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "VanBan.pdf");

			var sfd = new Microsoft.Win32.SaveFileDialog
			{
				Filter = "Tệp PDF (*.pdf)|*.pdf",
				FileName = Path.GetFileName(defaultPdf),
				Title = "Xuất tài liệu hiện tại sang PDF chuẩn"
			};

			if (sfd.ShowDialog() == true)
			{
				try
				{
					d.ExportAsFixedFormat(sfd.FileName, WdExportFormat.wdExportFormatPDF, OpenAfterExport: true);
					ShowModernSuccess($"Đã xuất tài liệu sang PDF thành công:\n{sfd.FileName}", "OFFICEFORMAT - Xuất PDF");
				}
				catch (Exception ex)
				{
					ShowModernError($"Lỗi khi xuất PDF: {ex.Message}", "Lỗi xuất PDF");
				}
			}
		}
	}

	public void Ribbon_CreateDocQr(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				EnsureWpfApplication();
				var dlg = new QrCodeDialog((Document)doc);
				dlg.SelectedTabIndex = 0;
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở hộp thoại tạo mã QR văn bản: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_CreateSelectionQr(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				EnsureWpfApplication();
				string selText = "";
				try
				{
					selText = _wordApp?.Selection?.Text ?? "";
				}
				catch { }
				var dlg = new QrCodeDialog((Document)doc, selText);
				dlg.SelectedTabIndex = 1;
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở hộp thoại tạo mã QR từ vùng chọn: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_CreateVietQr(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				EnsureWpfApplication();
				var dlg = new QrCodeDialog((Document)doc);
				dlg.SelectedTabIndex = 2;
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở hộp thoại tạo mã VietQR: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	public void Ribbon_OpenQrManager(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			try
			{
				EnsureWpfApplication();
				string selText = "";
				try
				{
					selText = _wordApp?.Selection?.Text ?? "";
				}
				catch { }
				var dlg = new QrCodeDialog((Document)doc, selText);
				IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
				if (handle != IntPtr.Zero)
				{
					new WindowInteropHelper(dlg).Owner = handle;
				}
				dlg.ShowDialog();
			}
			catch (Exception ex)
			{
				ShowModernError("Lỗi khi mở hộp thoại quản lý mã QR: " + ex.Message, "OFFICEFORMAT");
			}
		}
	}

	#endregion


	public void Ribbon_InsertQuocHieuAuto(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
			userPersonalization.ActiveRulePack = _selectedRulePack;
			_engine.RuleEngine.SetActiveProfile(_selectedRulePack);
			_engine.InsertAgencyHeader((Document)doc, userPersonalization.GetSelectedAgency());
			SetTaskPaneStatus("Đã chèn khối Tiêu đề Quốc hiệu & Cơ quan ban hành.");
		}
	}

	public void Ribbon_InsertQuocHieuCustom(IRibbonControl control)
	{
		Ribbon_InsertQuocHieuAuto(control);
	}

	public void Ribbon_InsertAdminAgency1(IRibbonControl control)
	{
		InsertAdminByIndex(0);
	}

	public void Ribbon_InsertAdminAgency2(IRibbonControl control)
	{
		InsertAdminByIndex(1);
	}

	public void Ribbon_InsertAdminAgency3(IRibbonControl control)
	{
		InsertAdminByIndex(2);
	}

	public void Ribbon_InsertAdminAgency4(IRibbonControl control)
	{
		InsertAdminByIndex(3);
	}

	public void Ribbon_InsertAdminAgency5(IRibbonControl control)
	{
		InsertAdminByIndex(4);
	}

	public void Ribbon_InsertAdminAgency6(IRibbonControl control)
	{
		InsertAdminByIndex(5);
	}

	private void InsertAdminByIndex(int index)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var _))
		{
			return;
		}
		UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
		if (userPersonalization.AgenciesAdmin != null && index < userPersonalization.AgenciesAdmin.Count)
		{
			AgencyProfile agencyProfile = userPersonalization.AgenciesAdmin[index];
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, agencyProfile.ParentAgency, agencyProfile.AgencyName, "Số: .../" + agencyProfile.CodePrefix, agencyProfile.Location + ", ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			return;
		}
		switch (index)
		{
		case 0:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "", "BỘ NỘI VỤ", "Số: .../BNV", "Hà Nội, ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			break;
		case 1:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ NỘI VỤ", "CỤC VĂN THƯ VÀ LƯU TRỮ NHÀ NƯỚC", "Số: .../VTLT", "Hà Nội, ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			break;
		case 2:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "UBND TỈNH...", "SỞ KHOA HỌC VÀ CÔNG NGHỆ", "Số: .../SKHCN", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			break;
		case 3:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "UBND TỈNH...", "UBND PHƯỜNG...", "Số: .../UBND", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			break;
		case 4:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "UBND TỈNH...", "UBND ĐẶC KHU...", "Số: .../UBND", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			break;
		default:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "UBND TỈNH...", "VĂN PHÒNG ỦY BAN NHÂN DÂN", "Số: .../VP", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
			break;
		}
	}

	public void Ribbon_InsertPartyAgency1(IRibbonControl control)
	{
		InsertPartyByIndex(0);
	}

	public void Ribbon_InsertPartyAgency2(IRibbonControl control)
	{
		InsertPartyByIndex(1);
	}

	public void Ribbon_InsertPartyAgency3(IRibbonControl control)
	{
		InsertPartyByIndex(2);
	}

	public void Ribbon_InsertPartyAgency4(IRibbonControl control)
	{
		InsertPartyByIndex(3);
	}

	private void InsertPartyByIndex(int index)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var _))
		{
			return;
		}
		UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
		if (userPersonalization.AgenciesParty != null && index < userPersonalization.AgenciesParty.Count)
		{
			AgencyProfile agencyProfile = userPersonalization.AgenciesParty[index];
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, agencyProfile.ParentAgency, agencyProfile.AgencyName, "Số     -" + agencyProfile.CodePrefix, agencyProfile.Location + ", ngày ... tháng ... năm 20...", isParty: true, isMilitary: false);
			return;
		}
		switch (index)
		{
		case 0:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "TỈNH ỦY...", "ĐẢNG ỦY CÁC CƠ QUAN ĐẢNG TỈNH", "Số     -NQ/DU", "..., ngày ... tháng ... năm 20...", isParty: true, isMilitary: false);
			break;
		case 1:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "TỈNH ỦY...", "BAN THƯỜNG VỤ", "Số     -NQ/TU", "..., ngày ... tháng ... năm 20...", isParty: true, isMilitary: false);
			break;
		case 2:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "TỈNH ỦY...", "ĐẢNG ỦY XÃ/PHƯỜNG...", "Số     -NQ/DU", "..., ngày ... tháng ... năm 20...", isParty: true, isMilitary: false);
			break;
		default:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "ĐẢNG ỦY XÃ/PHƯỜNG...", "CHI BỘ...", "Số     -NQ/CB", "..., ngày ... tháng ... năm 20...", isParty: true, isMilitary: false);
			break;
		}
	}

	public void Ribbon_InsertMilitaryAgency1(IRibbonControl control)
	{
		InsertMilByIndex(0);
	}

	public void Ribbon_InsertMilitaryAgency2(IRibbonControl control)
	{
		InsertMilByIndex(1);
	}

	public void Ribbon_InsertMilitaryAgency3(IRibbonControl control)
	{
		InsertMilByIndex(2);
	}

	public void Ribbon_InsertMilitaryAgency4(IRibbonControl control)
	{
		InsertMilByIndex(3);
	}

	public void Ribbon_InsertMilitaryAgency5(IRibbonControl control)
	{
		InsertMilByIndex(4);
	}

	public void Ribbon_InsertMilitaryAgency6(IRibbonControl control)
	{
		InsertMilByIndex(5);
	}

	private void InsertMilByIndex(int index)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var _))
		{
			return;
		}
		UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
		if (userPersonalization.AgenciesMilitary != null && index < userPersonalization.AgenciesMilitary.Count)
		{
			AgencyProfile agencyProfile = userPersonalization.AgenciesMilitary[index];
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, agencyProfile.ParentAgency, agencyProfile.AgencyName, "Số: .../" + agencyProfile.CodePrefix, agencyProfile.Location + ", ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			return;
		}
		switch (index)
		{
		case 0:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ QUỐC PHÒNG", "BỘ TỔNG THAM MƯU", "Số: .../TM", "Hà Nội, ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			break;
		case 1:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ QUỐC PHÒNG", "TỔNG CỤC CHÍNH TRỊ", "Số: .../CT", "Hà Nội, ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			break;
		case 2:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ QUỐC PHÒNG", "BỘ TƯ LỆNH QUÂN KHU...", "Số: .../QK", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			break;
		case 3:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ QUỐC PHÒNG", "BỘ TƯ LỆNH QUÂN ĐOÀN...", "Số: .../QĐ", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			break;
		case 4:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "QUÂN KHU...", "BỘ CHỈ HUY QUÂN SỰ TỈNH...", "Số: .../CHQS", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			break;
		default:
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ CHỈ HUY QUÂN SỰ...", "BAN CHỈ HUY QUÂN SỰ...", "Số: .../QS", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
			break;
		}
	}

	public void Ribbon_InsertQuocHieuND30(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "CƠ QUAN CHỦ QUẢN", "TÊN CƠ QUAN BAN HÀNH", "Số: .../...", "..., ngày ... tháng ... năm 20...", isParty: false, isMilitary: false);
		}
	}

	public void Ribbon_InsertDangHeader(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "TỈNH ỦY...", "BAN THƯỜNG VỤ", "Số     -NQ/TU", "..., ngày ... tháng ... năm 20...", isParty: true, isMilitary: false);
		}
	}

	public void Ribbon_InsertMilitaryHeader(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertCustomHeader((Selection)_wordApp.Selection, "BỘ QUỐC PHÒNG", "BỘ TỔNG THAM MƯU", "Số: .../TM-...", "Hà Nội, ngày ... tháng ... năm 20...", isParty: false, isMilitary: true);
		}
	}

	public void Ribbon_InsertDoanHeader(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, "TỈNH ĐOÀN...", "BCH ĐOÀN...", "Số:     -.../ĐTN", "..., ngày ... tháng ... năm 20...");
		}
	}

	public void Ribbon_InsertYouthAgency1(IRibbonControl control)
	{
		InsertYouthByIndex(0);
	}

	public void Ribbon_InsertYouthAgency2(IRibbonControl control)
	{
		InsertYouthByIndex(1);
	}

	public void Ribbon_InsertYouthAgency3(IRibbonControl control)
	{
		InsertYouthByIndex(2);
	}

	public void Ribbon_InsertYouthAgency4(IRibbonControl control)
	{
		InsertYouthByIndex(3);
	}

	public void Ribbon_InsertYouthAgency5(IRibbonControl control)
	{
		InsertYouthByIndex(4);
	}

	public void Ribbon_InsertYouthAgency6(IRibbonControl control)
	{
		InsertYouthByIndex(5);
	}

	public void Ribbon_InsertYouthAgency7(IRibbonControl control)
	{
		InsertYouthByIndex(6);
	}

	public void Ribbon_InsertYouthAgency8(IRibbonControl control)
	{
		InsertYouthByIndex(7);
	}

	private static string BuildYouthNumberPlaceholder(string codePrefix)
	{
		string text = (string.IsNullOrEmpty(codePrefix) ? "ĐTN" : codePrefix.Trim());
		if (string.Equals(text, "ĐH", StringComparison.OrdinalIgnoreCase))
		{
			return "Số:     /ĐH";
		}
		return "Số:     -.../" + text;
	}

	private void InsertYouthByIndex(int index)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var _))
		{
			return;
		}
		UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
		if (userPersonalization.AgenciesYouth != null && index < userPersonalization.AgenciesYouth.Count)
		{
			AgencyProfile agencyProfile = userPersonalization.AgenciesYouth[index];
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, agencyProfile.ParentAgency, agencyProfile.AgencyName, BuildYouthNumberPlaceholder(agencyProfile.CodePrefix), (string.IsNullOrEmpty(agencyProfile.Location) ? "..." : agencyProfile.Location) + ", ngày ... tháng ... năm 20...");
			return;
		}
		switch (index)
		{
		case 0:
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, "", "BAN CHẤP HÀNH TRUNG ƯƠNG", "Số:     -.../TWĐTN-VP", "Hà Nội, ngày ... tháng ... năm 20...");
			break;
		case 1:
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, "", "BCH ĐOÀN TỈNH...", "Số:     -.../TĐTN", "..., ngày ... tháng ... năm 20...");
			break;
		case 2:
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, "TỈNH ĐOÀN...", "BCH ĐOÀN PHƯỜNG...", "Số:     -.../ĐTN", "..., ngày ... tháng ... năm 20...");
			break;
		case 3:
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, "TỈNH ĐOÀN...", "BCH ĐOÀN XÃ...", "Số:     -.../ĐTN", "..., ngày ... tháng ... năm 20...");
			break;
		default:
			WordToolEngine.InsertYouthUnionHeader((Selection)_wordApp.Selection, "ĐOÀN XÃ...", "CHI ĐOÀN...", "Số:     -.../CĐ", "..., ngày ... tháng ... năm 20...");
			break;
		}
	}

	public void Ribbon_InsertYouthSig1(IRibbonControl control)
	{
		InsertYouthSigByIndex(0);
	}

	public void Ribbon_InsertYouthSig2(IRibbonControl control)
	{
		InsertYouthSigByIndex(1);
	}

	public void Ribbon_InsertYouthSig3(IRibbonControl control)
	{
		InsertYouthSigByIndex(2);
	}

	public void Ribbon_InsertYouthSig4(IRibbonControl control)
	{
		InsertYouthSigByIndex(3);
	}

	public void Ribbon_InsertYouthSig5(IRibbonControl control)
	{
		InsertYouthSigByIndex(4);
	}

	public void Ribbon_InsertYouthSig6(IRibbonControl control)
	{
		InsertYouthSigByIndex(5);
	}

	public void Ribbon_InsertYouthSig7(IRibbonControl control)
	{
		InsertYouthSigByIndex(6);
	}

	private void InsertYouthSigByIndex(int index)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
		List<SignerProfile> signersYouth = userPersonalization.SignersYouth;
		if (signersYouth == null || signersYouth.Count == 0)
		{
			signersYouth = UserPersonalization.CreateDefault().SignersYouth;
		}
		if (signersYouth != null && signersYouth.Count != 0)
		{
			SignerProfile signerProfile = signersYouth[Math.Min(Math.Max(index, 0), signersYouth.Count - 1)];
			RecipientProfile recipient = new RecipientProfile(signerProfile.Id + "_rec", signerProfile.Alias ?? "Nơi nhận", string.IsNullOrEmpty(signerProfile.RecipientList) ? "- Như trên;\n- Lưu VP." : signerProfile.RecipientList, RulePackType.HD29_TWDTN);
			_engine.RuleEngine.SetActiveProfile(RulePackType.HD29_TWDTN);
			try
			{
				_engine.InsertSignerBlock((Document)doc, signerProfile, recipient);
			}
			finally
			{
				_engine.RuleEngine.SetActiveProfile(_selectedRulePack);
			}
			SetTaskPaneStatus("Đã chèn khối Nơi nhận & Ký tên văn bản Đoàn (HD 29).");
		}
	}

	public void Ribbon_InsertPhuLucHeader(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertDocTitle((Selection)_wordApp.Selection, "Phụ lục I", "Kèm theo Quyết định số: .../... ngày .../.../20...");
		}
	}

	public void Ribbon_InsertBanSaoHeader(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertDocTitle((Selection)_wordApp.Selection, "SAO Y", "(Bản sao hợp lệ theo Mẫu 3.1 NĐ 30/2020/NĐ-CP)");
		}
	}

	public void Ribbon_InsertPersonalMottoHeader(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertPersonalMottoHeader((Selection)_wordApp.Selection);
		}
	}

	private void InsertDocTitleHelper(string title, string summary)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertDocTitle((Selection)_wordApp.Selection, title, summary);
		}
	}

	public void Ribbon_InsertTitleCongVan(IRibbonControl control)
	{
		InsertDocTitleHelper("V/v ...", "Kính gửi: ...");
	}

	public void Ribbon_InsertTitleQuyetDinh(IRibbonControl control)
	{
		InsertDocTitleHelper("QUYẾT ĐỊNH", "Về việc ...");
	}

	public void Ribbon_InsertTitleQuyetDinhQuyChe(IRibbonControl control)
	{
		InsertDocTitleHelper("QUYẾT ĐỊNH", "Ban hành (Phê duyệt) Quy chế ...");
	}

	public void Ribbon_InsertTitleNghiQuyet(IRibbonControl control)
	{
		InsertDocTitleHelper("NGHỊ QUYẾT", "Về việc ...");
	}

	public void Ribbon_InsertTitleToTrinh(IRibbonControl control)
	{
		InsertDocTitleHelper("TỜ TRÌNH", "Về việc ...");
	}

	public void Ribbon_InsertTitleKeHoach(IRibbonControl control)
	{
		InsertDocTitleHelper("KẾ HOẠCH", "Thực hiện ...");
	}

	public void Ribbon_InsertTitleBaoCao(IRibbonControl control)
	{
		InsertDocTitleHelper("BÁO CÁO", "Kết quả ...");
	}

	public void Ribbon_InsertTitleThongBao(IRibbonControl control)
	{
		InsertDocTitleHelper("THÔNG BÁO", "Về việc ...");
	}

	public void Ribbon_InsertTitleBienBan(IRibbonControl control)
	{
		InsertDocTitleHelper("BIÊN BẢN", "Hoạt động ...");
	}

	public void Ribbon_InsertTitleCongDien(IRibbonControl control)
	{
		InsertDocTitleHelper("CÔNG ĐIỆN", "Về việc ...");
	}

	public void Ribbon_InsertTitleGiayMoi(IRibbonControl control)
	{
		InsertDocTitleHelper("GIẤY MỜI", "Dự cuộc họp ...");
	}

	public void Ribbon_InsertTitleGiayGioiThieu(IRibbonControl control)
	{
		InsertDocTitleHelper("GIẤY GIỚI THIỆU", "");
	}

	public void Ribbon_InsertTitleGiayNghiPhep(IRibbonControl control)
	{
		InsertDocTitleHelper("GIẤY NGHỈ PHÉP", "");
	}

	public void Ribbon_InsertSigAuto(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
			userPersonalization.ActiveRulePack = _selectedRulePack;
			_engine.RuleEngine.SetActiveProfile(_selectedRulePack);
			_engine.InsertSignerBlock((Document)doc, userPersonalization.GetSelectedSigner(), userPersonalization.GetSelectedRecipient());
			SetTaskPaneStatus("Đã chèn khối Nơi nhận & Ký tên cuối văn bản.");
		}
	}

	public void Ribbon_InsertSigCustom(IRibbonControl control)
	{
		Ribbon_InsertSigAuto(control);
	}

	public void Ribbon_InsertSigSaved1(IRibbonControl control)
	{
		InsertSavedSigByIndex(0);
	}

	public void Ribbon_InsertSigSaved2(IRibbonControl control)
	{
		InsertSavedSigByIndex(1);
	}

	public void Ribbon_InsertSigSaved3(IRibbonControl control)
	{
		InsertSavedSigByIndex(2);
	}

	public void Ribbon_InsertSigSaved4(IRibbonControl control)
	{
		InsertSavedSigByIndex(3);
	}

	public void Ribbon_InsertSigSaved5(IRibbonControl control)
	{
		InsertSavedSigByIndex(4);
	}

	public void Ribbon_InsertSigSaved6(IRibbonControl control)
	{
		InsertSavedSigByIndex(5);
	}

	public void Ribbon_InsertSigSaved7(IRibbonControl control)
	{
		InsertSavedSigByIndex(6);
	}

	public void Ribbon_InsertSigSaved8(IRibbonControl control)
	{
		InsertSavedSigByIndex(7);
	}

	public void Ribbon_InsertSigSaved9(IRibbonControl control)
	{
		InsertSavedSigByIndex(8);
	}

	public void Ribbon_InsertSigSaved10(IRibbonControl control)
	{
		InsertSavedSigByIndex(9);
	}

	public void Ribbon_InsertSigSaved11(IRibbonControl control)
	{
		InsertSavedSigByIndex(10);
	}

	private void InsertSavedSigByIndex(int index)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var _))
		{
			return;
		}
		UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
		string text = "GIÁM ĐỐC";
		string text2 = "Nguyễn Văn A";
		bool isPartySig = _selectedRulePack == RulePackType.HD05_VPTW;
		if (userPersonalization.Signers != null && index < userPersonalization.Signers.Count)
		{
			text = ((!string.IsNullOrEmpty(userPersonalization.Signers[index].AuthorityTitle)) ? userPersonalization.Signers[index].AuthorityTitle : userPersonalization.Signers[index].Position);
			text2 = userPersonalization.Signers[index].FullName;
			if (userPersonalization.Signers[index].RulePack == RulePackType.HD05_VPTW)
			{
				isPartySig = true;
			}
		}
		else
		{
			switch (index)
			{
			case 0:
				text = "BỘ TRƯỞNG";
				text2 = "Nguyễn Văn A";
				break;
			case 1:
				text = "KT. BỘ TRƯỞNG\nTHỨ TRƯỞNG";
				text2 = "Trần Văn B";
				break;
			case 2:
				text = "CỤC TRƯỞNG";
				text2 = "Lê Văn C";
				break;
			case 3:
				text = "GIÁM ĐỐC";
				text2 = "Nguyễn Văn A";
				break;
			case 4:
				text = "TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH";
				text2 = "Trần Văn B";
				break;
			case 5:
				text = "T/M BAN CHẤP HÀNH\nBÍ THƯ";
				text2 = "Lê Văn C";
				isPartySig = true;
				break;
			case 6:
				text = "T/M BAN THƯỜNG VỤ\nPHÓ BÍ THƯ";
				text2 = "Nguyễn Văn A";
				isPartySig = true;
				break;
			case 7:
				text = "K/T CHÁNH VĂN PHÒNG\nPHÓ CHÁNH VĂN PHÒNG";
				text2 = "Lê Khánh Toàn";
				isPartySig = true;
				break;
			case 8:
				text = "T/L BAN THƯỜNG VỤ\nCHÁNH VĂN PHÒNG";
				text2 = "Hoàng Văn E";
				isPartySig = true;
				break;
			case 9:
				text = "T/M CHI BỘ\nBÍ THƯ";
				text2 = "Phạm Văn D";
				isPartySig = true;
				break;
			default:
				text = "TL. CHỦ TỊCH\nCHÁNH VĂN PHÒNG";
				text2 = "Hoàng Văn E";
				break;
			}
		}
		WordToolEngine.InsertCustomSignature((Selection)_wordApp.Selection, isPartySig ? "- Như trên;\n- Lưu: VP." : "- Như trên;\n- Lưu: VT.", text, text2, isPartySig);
	}

	public void Ribbon_InsertRecSaved1(IRibbonControl control)
	{
		InsertSavedRecByIndex(0);
	}

	public void Ribbon_InsertRecSaved2(IRibbonControl control)
	{
		InsertSavedRecByIndex(1);
	}

	public void Ribbon_InsertRecSaved3(IRibbonControl control)
	{
		InsertSavedRecByIndex(2);
	}

	public void Ribbon_InsertRecSaved4(IRibbonControl control)
	{
		InsertSavedRecByIndex(3);
	}

	public void Ribbon_InsertRecSaved5(IRibbonControl control)
	{
		InsertSavedRecByIndex(4);
	}

	private void InsertSavedRecByIndex(int index)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
			string text = "- Như trên;\n- Lưu: VT.";
			text = ((userPersonalization.Recipients != null && index < userPersonalization.Recipients.Count) ? userPersonalization.Recipients[index].Content : (index switch
			{
				0 => "- Như Điều ...;\n- Thủ tướng, các Phó Thủ tướng Chính phủ;\n- Lưu: VT, PC.", 
				1 => "- Như kính gửi;\n- Giám đốc, các Phó Giám đốc;\n- Lưu: VT, TH.", 
				2 => "- Như Điều 3;\n- Thường trực Tỉnh ủy, HĐND tỉnh, UBND tỉnh;\n- Lưu: VT, TCCB.", 
				3 => "- Đảng ủy cấp trên (để b/c);\n- Các đồng chí đảng viên trong Chi bộ;\n- Lưu: Chi bộ.", 
				_ => "- Như Điều ...;\n- Thủ trưởng Bộ Quốc phòng (để b/c);\n- Lưu: VT, ...;", 
			}));
			bool isPartyRec = _selectedRulePack == RulePackType.HD05_VPTW || index == 3 || (userPersonalization.Recipients != null && index < userPersonalization.Recipients.Count && userPersonalization.Recipients[index].RulePack == RulePackType.HD05_VPTW);
			bool isYouthRec = !isPartyRec && (_selectedRulePack == RulePackType.HD29_TWDTN || (userPersonalization.Recipients != null && index < userPersonalization.Recipients.Count && userPersonalization.Recipients[index].RulePack == RulePackType.HD29_TWDTN));
			Selection selection = _wordApp.Selection;
			if (selection != null)
			{
				Range recRange = selection.Range;
				recRange.Text = "Nơi nhận:\r\n" + text.Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n";
				recRange.Font.Name = "Times New Roman";
				recRange.Font.Size = (isPartyRec ? 12f : 11f);
				recRange.Font.Bold = 0;
				recRange.Font.Italic = 0;
				recRange.Font.Underline = WdUnderline.wdUnderlineNone;
				recRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
				recRange.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
				recRange.ParagraphFormat.SpaceBefore = 0f;
				recRange.ParagraphFormat.SpaceAfter = 0f;
				recRange.ParagraphFormat.FirstLineIndent = 0f;
				if (recRange.Paragraphs.Count > 0)
				{
					Paragraph hdrPara = recRange.Paragraphs[1];
					hdrPara.Range.Font.Size = (isPartyRec ? 14f : (isYouthRec ? 13f : 12f));
					hdrPara.Range.Font.Bold = (isPartyRec ? 0 : (-1));
					hdrPara.Range.Font.Italic = ((!isPartyRec && !isYouthRec) ? (-1) : 0);
					hdrPara.Range.Font.Underline = (isPartyRec ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
					for (int pIdx = 2; pIdx <= recRange.Paragraphs.Count; pIdx++)
					{
						Paragraph itemP = recRange.Paragraphs[pIdx];
						itemP.Range.Font.Size = (isPartyRec ? 12f : 11f);
						itemP.Range.Font.Bold = 0;
						itemP.Range.Font.Italic = 0;
						itemP.Range.Font.Underline = WdUnderline.wdUnderlineNone;
					}
				}
			}
		}
	}

	private void InsertStandardSigHelper(string recipients, string title, string name, bool? isParty = null)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			bool partyMode = isParty ?? (_selectedRulePack == RulePackType.HD05_VPTW || (title ?? "").Contains("CHI BỘ") || (title ?? "").Contains("BAN CHẤP HÀNH") || (title ?? "").Contains("BAN THƯỜNG VỤ") || (title ?? "").Contains("BÍ THƯ") || (title ?? "").Contains("T/M") || (title ?? "").Contains("K/T") || (title ?? "").Contains("T/L"));
			WordToolEngine.InsertCustomSignature((Selection)_wordApp.Selection, recipients, title, name, partyMode);
		}
	}

	public void Ribbon_InsertSigChiBo(IRibbonControl control)
	{
		InsertStandardSigHelper("- Đảng ủy phường (b/c);\n- Lưu: Chi bộ.", "T/M CHI BỘ\r\nBÍ THƯ", "Nguyễn Văn A", isParty: true);
	}

	public void Ribbon_InsertSigChiBoPho(IRibbonControl control)
	{
		InsertStandardSigHelper("- Đảng ủy phường (b/c);\n- Lưu: Chi bộ.", "T/M CHI BỘ\r\nPHÓ BÍ THƯ", "Trần Văn B", isParty: true);
	}

	public void Ribbon_InsertSigDangUyXa(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như kính gửi;\n- Lưu: VP.", "T/M BAN CHẤP HÀNH\r\nBÍ THƯ", "Nguyễn Văn A", isParty: true);
	}

	public void Ribbon_InsertSigBTV(IRibbonControl control)
	{
		InsertStandardSigHelper("- Thường trực Tỉnh ủy;\n- Lưu: VP.", "T/M BAN THƯỜNG VỤ\r\nPHÓ BÍ THƯ", "Nguyễn Văn B", isParty: true);
	}

	public void Ribbon_InsertSigBCH(IRibbonControl control)
	{
		InsertStandardSigHelper("- Ban Thường vụ;\n- Lưu: VP.", "T/M BAN CHẤP HÀNH\r\nBÍ THƯ", "Nguyễn Văn A", isParty: true);
	}

	public void Ribbon_InsertSigUBNDXa(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như trên;\n- Lưu: VT.", "TM. ỦY BAN NHÂN DÂN\r\nCHỦ TỊCH", "Trần Văn B");
	}

	public void Ribbon_InsertSigUBND(IRibbonControl control)
	{
		InsertStandardSigHelper("- Thường trực Tỉnh ủy, HĐND tỉnh;\n- Lưu: VT, TH.", "TM. ỦY BAN NHÂN DÂN\r\nCHỦ TỊCH", "Nguyễn Văn A");
	}

	public void Ribbon_InsertSigThuTruong(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như trên;\n- Lưu: VT.", "GIÁM ĐỐC", "Nguyễn Văn A");
	}

	public void Ribbon_InsertSigKT(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như trên;\n- Lưu: VT.", "KT. CHỦ TỊCH\r\nPHÓ CHỦ TỊCH", "Nguyễn Văn C");
	}

	public void Ribbon_InsertSigTL(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như trên;\n- Lưu: VT.", "TL. CHỦ TỊCH\r\nCHÁNH VĂN PHÒNG", "Trần Văn D");
	}

	public void Ribbon_InsertSigTUQ(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như trên;\n- Lưu: VT.", "TUQ. CHỦ TỊCH\r\nTRƯỞNG PHÒNG", "Lê Văn E");
	}

	public void Ribbon_InsertSigQuyen(IRibbonControl control)
	{
		InsertStandardSigHelper("- Như trên;\n- Lưu: VT.", "Q. CHỦ TỊCH", "Nguyễn Văn F");
	}

	private void InsertCalloutHelper(string text, bool isRed, float width = 36f, float height = 8.5f)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var _))
		{
			WordToolEngine.InsertCalloutBox((Microsoft.Office.Interop.Word.Application)_wordApp, text, isRed, width, height);
		}
	}

	public void Ribbon_InsertDoKhanHoaToc(IRibbonControl control)
	{
		InsertCalloutHelper("HỎA TỐC", isRed: true);
	}

	public void Ribbon_InsertDoKhanThuongKhan(IRibbonControl control)
	{
		InsertCalloutHelper("THƯỢNG KHẨN", isRed: true);
	}

	public void Ribbon_InsertDoKhanKhan(IRibbonControl control)
	{
		InsertCalloutHelper("KHẨN", isRed: true);
	}

	public void Ribbon_InsertDoMatMat(IRibbonControl control)
	{
		InsertCalloutHelper("MẬT", isRed: true);
	}

	public void Ribbon_InsertDoMatToiMat(IRibbonControl control)
	{
		InsertCalloutHelper("TỐI MẬT", isRed: true);
	}

	public void Ribbon_InsertDoMatTuyetMat(IRibbonControl control)
	{
		InsertCalloutHelper("TUYỆT MẬT", isRed: true);
	}

	public void Ribbon_InsertPhamViLuuHanh(IRibbonControl control)
	{
		InsertCalloutHelper("LƯU HÀNH NỘI BỘ", isRed: false, 52f);
	}

	public void Ribbon_InsertPhamViXemXongTraLai(IRibbonControl control)
	{
		InsertCalloutHelper("XEM XONG TRẢ LẠI", isRed: false, 52f);
	}

	public void Ribbon_InsertFooterAgencyInfo(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			WordToolEngine.InsertFooterAgencyInfo((Document)doc, (Selection)_wordApp.Selection);
		}
	}

	public void Ribbon_FormatTables(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		int expectedSteps = 0;
		try
		{
			expectedSteps = doc.Tables.Count;
		}
		catch
		{
		}
		FormattingResult res = null;
		if (RunLongOperationWithNotices("OFFICEFORMAT - Chuẩn Hóa Bảng Biểu", "Đang chuẩn hóa bảng biểu theo quy chuẩn...", () =>
		{
			res = _engine.FixDocumentTables((Document)doc, _selectedRulePack);
		}, expectedSteps))
		{
			if (res.Success)
			{
				ShowModernSuccess($"Đã chuẩn hóa thành công {res.TablesProcessed} bảng biểu theo quy chuẩn!");
			}
			else
			{
				ShowModernError("Không thể chuẩn hóa bảng biểu: " + res.Message);
			}
		}
	}

	public void Ribbon_CleanSpaces(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc))
		{
			Selection cleanSel = null;
			try
			{
				cleanSel = _wordApp.Selection;
			}
			catch
			{
			}
			if (RunLongOperationWithNotices("OFFICEFORMAT - Xóa Khoảng Trắng Lỗi", "Đang dọn khoảng trắng thừa và dấu cách sai vị trí...", () =>
			{
				WordToolEngine.CleanSpacingErrors((Document)doc, (Selection)cleanSel);
			}, 169))
			{
				ShowModernSuccess("Đã làm sạch khoảng trắng lỗi và dấu cách bất hợp lý thành công!");
			}
		}
	}

	public void Ribbon_RemoveToneSelection(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var _))
		{
			return;
		}
		Selection sel = _wordApp.Selection;
		if (sel == null || sel.Type != WdSelectionType.wdSelectionNormal)
		{
			ShowModernInfo("Vui lòng bôi đen đoạn văn bản cần bỏ dấu tiếng Việt!");
			return;
		}
		RunLongOperationWithNotices("OFFICEFORMAT - Bỏ Dấu Tiếng Việt", "Đang chuyển đoạn bôi đen sang tiếng Việt không dấu...", () =>
		{
			WordToolEngine.RemoveToneSelection((Selection)sel);
		});
	}

	public void Ribbon_FixToneMarks(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc) && RunLongOperationWithNotices("OFFICEFORMAT - Sửa Dấu Thanh Kiểu Mới", "Đang chuẩn hóa dấu thanh kiểu mới (hòa, tòa, thủy...)...", () =>
		{
			WordToolEngine.StandardizeToneMarks((Document)doc);
		}))
		{
			ShowModernSuccess("Đã chuẩn hóa dấu thanh toàn văn bản sang kiểu mới (hòa, tòa, thủy, ủy, hoàn, quản, Quý...) thành công!");
		}
	}

	public void Ribbon_FixToneMarksOld(IRibbonControl control)
	{
		if (EnsureLicenseActive() && EnsureActiveDocument(out var doc) && RunLongOperationWithNotices("OFFICEFORMAT - Chuyển Dấu Thanh Kiểu Cũ", "Đang chuyển dấu thanh sang kiểu cũ (hoà, toà, thuỷ...)...", () =>
		{
			WordToolEngine.ConvertToTraditionalToneMarks((Document)doc);
		}))
		{
			ShowModernSuccess("Đã chuyển đổi dấu thanh toàn văn bản sang kiểu cũ (hoà, toà, thuỷ, uỷ, hoàn, quản, Quý...) thành công!");
		}
	}

	public void Ribbon_NumberToVndWords(IRibbonControl control)
	{
		if (!EnsureLicenseActive() || !EnsureActiveDocument(out var doc))
		{
			return;
		}
		Selection selection = null;
		try
		{
			selection = _wordApp.Selection;
		}
		catch
		{
		}
		if (selection == null || selection.Type != WdSelectionType.wdSelectionNormal || selection.Start == selection.End)
		{
			ShowModernInfo("Vui lòng bôi đen số tiền cần chuyển (ví dụ 2345 hoặc 2.345.000) rồi bấm lại.", "OFFICEFORMAT - Chuyển số tiền sang chữ");
			return;
		}
		try
		{
			Range range = selection.Range;
			string text = range.Text ?? string.Empty;
			int i;
			for (i = 0; i < text.Length && " \t\r\n\a\v\f\u00a0".IndexOf(text[i]) >= 0; i++)
			{
			}
			int j;
			for (j = 0; j < text.Length - i && " \t\r\n\a\v\f\u00a0".IndexOf(text[text.Length - 1 - j]) >= 0; j++)
			{
			}
			if (i + j >= text.Length)
			{
				ShowModernInfo("Vui lòng bôi đen số tiền cần chuyển (ví dụ 2345 hoặc 2.345.000) rồi bấm lại.", "OFFICEFORMAT - Chuyển số tiền sang chữ");
				return;
			}
			string text2 = text.Substring(i, text.Length - i - j);
			VndReadingOptions options = null;
			try
			{
				options = _profileRepo.LoadPersonalization().GetVndReadingOptions();
			}
			catch
			{
			}
			if (!VndAmountReader.TryConvert(text2, options, out var words, out var error))
			{
				string text3 = ((text2.Length > 60) ? (text2.Substring(0, 60) + "...") : text2);
				ShowModernInfo(error + "\r\n\r\nNội dung đã chọn: \"" + text3 + "\"", "OFFICEFORMAT - Chuyển số tiền sang chữ");
				return;
			}
			int num = range.Start + i;
			int num2 = range.End - j;
			using (new SafeUndoScope((Microsoft.Office.Interop.Word.Application)_wordApp, "OfficeFormat Chuyển số tiền sang chữ"))
			{
				Document document = doc;
				object Start = num;
				object End = num2;
				Range range2 = document.Range(ref Start, ref End);
				range2.Text = words;
				try
				{
					selection.SetRange(range2.Start, range2.End);
				}
				catch
				{
				}
			}
			try
			{
				_wordApp.StatusBar = "OfficeFormat: " + words;
			}
			catch
			{
			}
			SetTaskPaneStatus("Đã chuyển số tiền sang chữ.");
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể chuyển số tiền sang chữ: " + ex.Message, "OFFICEFORMAT - Chuyển số tiền sang chữ");
		}
	}

	public void Ribbon_OpenTemplateByTag(IRibbonControl control)
	{
		string text = control?.Tag;
		if (!string.IsNullOrEmpty(text))
		{
			OpenTemplateHelper(text);
		}
	}

	private void OpenTemplateHelper(string templateFileName)
	{
		if (EnsureLicenseActive() && !WordToolEngine.OpenTemplate((Microsoft.Office.Interop.Word.Application)_wordApp, templateFileName, out var errorMessage))
		{
			ShowModernWarning(errorMessage);
		}
	}

	public string GetMyTemplatesMenuContent(IRibbonControl control)
	{
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.Append("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
		stringBuilder.Append("<button id=\"btnDynAddUserTpl\" label=\"➕ Thêm biểu mẫu mới vào Thư viện...\" getImage=\"Ribbon_GetIcon\" onAction=\"Ribbon_QuickImportUserTemplate\" />");
		stringBuilder.Append("<button id=\"btnDynOpenMgr\" label=\"⚙\ufe0f Quản lý &amp;&amp; Chỉnh sửa biểu mẫu...\" getImage=\"Ribbon_GetIcon\" onAction=\"Ribbon_OpenMyTemplatesManager\" />");
		stringBuilder.Append("<menuSeparator id=\"sepDynUserTpl\" />");
		try
		{
			List<UserTemplateItem> allTemplates = UserTemplateService.Instance.GetAllTemplates();
			if (allTemplates != null && allTemplates.Count > 0)
			{
				int num = 0;
				foreach (UserTemplateItem item in allTemplates)
				{
					string text = EscapeXml(item.Name);
					string text2 = ((!string.IsNullOrEmpty(item.Description)) ? (" supertip=\"" + EscapeXml(item.Description) + "\"") : "");
					string text3 = SanitizeRibbonId(item.Id);
					stringBuilder.Append($"<button id=\"btnUserTpl_{num}_{text3}\" label=\"{text}\" tag=\"{item.Id}\" getImage=\"Ribbon_GetIcon\" onAction=\"Ribbon_OpenUserTemplate\"{text2} />");
					num++;
				}
			}
			else
			{
				stringBuilder.Append("<button id=\"btnNoUserTpl\" label=\"(Chưa có biểu mẫu nào - Bấm để thêm)\" onAction=\"Ribbon_QuickImportUserTemplate\" getImage=\"Ribbon_GetIcon\" />");
			}
		}
		catch
		{
			stringBuilder.Append("<button id=\"btnErrUserTpl\" label=\"(Không thể nạp danh sách biểu mẫu)\" enabled=\"false\" />");
		}
		stringBuilder.Append("</menu>");
		return stringBuilder.ToString();
	}

	public void Ribbon_OpenMyTemplatesManager(IRibbonControl control)
	{
		ShowMyTemplatesManagerDialog();
	}

	public void Ribbon_QuickImportUserTemplate(IRibbonControl control)
	{
		try
		{
			EnsureWpfApplication();
			AddUserTemplateDialog addUserTemplateDialog = new AddUserTemplateDialog();
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(addUserTemplateDialog).Owner = mainWindowHandle;
			}
			if (addUserTemplateDialog.ShowDialog() == true)
			{
				RefreshUserTemplatesRibbon();
				ShowModernSuccess($"Đã lưu biểu mẫu '{addUserTemplateDialog.CreatedTemplateItem.Name}' vào Thư viện cá nhân của bạn!");
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở hộp thoại thêm biểu mẫu: " + ex.Message);
		}
	}

	public void Ribbon_OpenUserTemplate(IRibbonControl control)
	{
		if (control == null)
		{
			return;
		}
		string text = control.Tag;
		if (string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(control.Id) && control.Id.StartsWith("btnUserTpl_"))
		{
			string text2 = control.Id.Substring("btnUserTpl_".Length);
			string[] array = text2.Split(new char[1] { '_' }, 2);
			string templateId = ((array.Length == 2) ? array[1] : array[0]);
			UserTemplateItem templateById = UserTemplateService.Instance.GetTemplateById(templateId);
			int result;
			if (templateById != null)
			{
				text = templateById.Id;
			}
			else if (int.TryParse(array[0], out result))
			{
				List<UserTemplateItem> allTemplates = UserTemplateService.Instance.GetAllTemplates();
				if (result >= 0 && result < allTemplates.Count)
				{
					text = allTemplates[result].Id;
				}
			}
		}
		if (!string.IsNullOrEmpty(text))
		{
			OpenUserTemplateById(text);
		}
		else
		{
			ShowMyTemplatesManagerDialog();
		}
	}

	public void OpenUserTemplateById(string templateId)
	{
		if (!EnsureLicenseActive())
		{
			return;
		}
		UserTemplateItem templateById = UserTemplateService.Instance.GetTemplateById(templateId);
		if (templateById == null)
		{
			ShowModernWarning("Không tìm thấy thông tin biểu mẫu được chọn trong thư viện.");
			return;
		}
		string text = UserTemplateService.Instance.GetTemplateFullPath(templateById);
		if (string.IsNullOrEmpty(text) || !File.Exists(text) || new FileInfo(text).Length < 500)
		{
			text = UserTemplateService.Instance.EnsureValidTemplateFile(templateById);
		}
		if (!WordToolEngine.OpenTemplate((Microsoft.Office.Interop.Word.Application)_wordApp, text, out var errorMessage))
		{
			ShowModernWarning(errorMessage);
		}
	}

	public void ShowTemplateLibraryDialog(string initialCategoryName = null)
	{
		try
		{
			EnsureWpfApplication();
			TemplateLibraryWindow templateLibraryWindow = new TemplateLibraryWindow(initialCategoryName);
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(templateLibraryWindow).Owner = mainWindowHandle;
			}
			templateLibraryWindow.OpenTemplateCallback = (TemplateItem item) =>
			{
				if (item != null)
				{
					if (item.IsUserTemplate)
					{
						OpenUserTemplateById(item.Id);
					}
					else
					{
						OpenTemplateHelper(item.File);
					}
				}
			};
			templateLibraryWindow.ShowDialog();
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở Thư viện biểu mẫu: " + ex.Message);
		}
	}

	public void Ribbon_OpenTemplateLibrary(IRibbonControl control)
	{
		ShowTemplateLibraryDialog();
	}

	public void Ribbon_OpenTDKTLibrary(IRibbonControl control)
	{
		ShowTemplateLibraryDialog("Thi đua – Khen thưởng");
	}

	public void ShowMyTemplatesManagerDialog()
	{
		try
		{
			EnsureWpfApplication();
			MyTemplatesManagerWindow myTemplatesManagerWindow = new MyTemplatesManagerWindow();
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(myTemplatesManagerWindow).Owner = mainWindowHandle;
			}
			myTemplatesManagerWindow.OpenTemplateCallback = (UserTemplateItem item) =>
			{
				if (item != null)
				{
					OpenUserTemplateById(item.Id);
				}
			};
			myTemplatesManagerWindow.TemplatesChangedCallback = () =>
			{
				RefreshUserTemplatesRibbon();
			};
			myTemplatesManagerWindow.ShowDialog();
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở cửa sổ Quản lý biểu mẫu: " + ex.Message);
		}
	}

	public void RefreshUserTemplatesRibbon()
	{
		try
		{
			if (_ribbon != null)
			{
				_ribbon.InvalidateControl("dynMyTemplates");
			}
		}
		catch
		{
		}
	}

	private static string EscapeXml(string str)
	{
		if (string.IsNullOrEmpty(str))
		{
			return "";
		}
		return str.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
			.Replace("\"", "&quot;")
			.Replace("'", "&apos;");
	}

	public void Ribbon_InsertTplCongVan(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Cong_van_hanh_chinh.docx");
	}

	public void Ribbon_InsertTplQuyetDinh(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Quyet_dinh.docx");
	}

	public void Ribbon_InsertTplQuyetDinhQuyChe(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Quyet_dinh_quy_che.docx");
	}

	public void Ribbon_InsertTplNghiQuyetCaBiet(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Nghi_quyet_ca_biet.docx");
	}

	public void Ribbon_InsertTplChiThi(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Chi_thi.docx");
	}

	public void Ribbon_InsertTplHuongDan(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Huong_dan.docx");
	}

	public void Ribbon_InsertTplQuyDinh(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Quy_dinh.docx");
	}

	public void Ribbon_InsertTplDeAn(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_De_an.docx");
	}

	public void Ribbon_InsertTplToTrinh(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_To_trinh.docx");
	}

	public void Ribbon_InsertTplKeHoach(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Ke_hoach.docx");
	}

	public void Ribbon_InsertTplBaoCao(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Bao_cao.docx");
	}

	public void Ribbon_InsertTplThongBao(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Thong_bao.docx");
	}

	public void Ribbon_InsertTplBienBan(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Bien_ban.docx");
	}

	public void Ribbon_InsertTplCongDien(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Cong_dien.docx");
	}

	public void Ribbon_InsertTplGiayMoi(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Giay_moi.docx");
	}

	public void Ribbon_InsertTplGiayGioiThieu(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Giay_gioi_thieu.docx");
	}

	public void Ribbon_InsertTplGiayNghiPhep(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Giay_nghi_phep.docx");
	}

	public void Ribbon_InsertTplGiayUyQuyen(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Giay_uy_quyen.docx");
	}

	public void Ribbon_InsertTplPhieuChuyen(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Phieu_chuyen.docx");
	}

	public void Ribbon_InsertTplPhuLuc(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Phu_luc_van_ban.docx");
	}

	public void Ribbon_InsertTplBanSao(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Ban_sao_van_ban.docx");
	}

	public void Ribbon_InsertTplXaCV(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Cong_van.docx");
	}

	public void Ribbon_InsertTplXaQD(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Quyet_dinh.docx");
	}

	public void Ribbon_InsertTplXaTT(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_To_trinh.docx");
	}

	public void Ribbon_InsertTplXaBC(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Bao_cao.docx");
	}

	public void Ribbon_InsertTplXaKH(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Ke_hoach.docx");
	}

	public void Ribbon_InsertTplXaTB(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Thong_bao.docx");
	}

	public void Ribbon_InsertTplXaBB(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Bien_ban.docx");
	}

	public void Ribbon_InsertTplXaGM(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Giay_moi.docx");
	}

	public void Ribbon_InsertTplXaGGT(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Hanh_chinh_Xa_Phuong_Giay_gioi_thieu.docx");
	}

	public void Ribbon_InsertTplDangNQ(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Nghi_quyet.docx");
	}

	public void Ribbon_InsertTplDangQD(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Quyet_dinh.docx");
	}

	public void Ribbon_InsertTplDangCT(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chi_thi.docx");
	}

	public void Ribbon_InsertTplDangKL(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Ket_luan.docx");
	}

	public void Ribbon_InsertTplDangQDi(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Quy_dinh.docx");
	}

	public void Ribbon_InsertTplDangCTr(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chuong_trinh.docx");
	}

	public void Ribbon_InsertTplDangKH(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Ke_hoach.docx");
	}

	public void Ribbon_InsertTplDangBC(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Bao_cao.docx");
	}

	public void Ribbon_InsertTplDangTB(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Thong_bao.docx");
	}

	public void Ribbon_InsertTplDangTTr(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_To_trinh.docx");
	}

	public void Ribbon_InsertTplDangTT(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Thong_tri.docx");
	}

	public void Ribbon_InsertTplDangCV(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Cong_van.docx");
	}

	public void Ribbon_InsertTplDangBB(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Bien_ban.docx");
	}

	public void Ribbon_InsertTplDangGM(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Giay_moi.docx");
	}

	public void Ribbon_InsertTplDangGGT(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Giay_gioi_thieu.docx");
	}

	public void Ribbon_InsertTplDangPhuLuc(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Phu_luc.docx");
	}

	public void Ribbon_InsertTplDangBanSao(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Ban_sao.docx");
	}

	public void Ribbon_InsertTplChiBoNQ(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chi_bo_Nghi_quyet.docx");
	}

	public void Ribbon_InsertTplChiBoBC(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chi_bo_Bao_cao.docx");
	}

	public void Ribbon_InsertTplChiBoKH(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chi_bo_Ke_hoach.docx");
	}

	public void Ribbon_InsertTplChiBoBB(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chi_bo_Bien_ban_sinh_hoat.docx");
	}

	public void Ribbon_InsertTplChiBoCV(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Dang_Chi_bo_Cong_van.docx");
	}

	public void Ribbon_InsertTplVBHN(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Van_ban_hop_nhat_VBHN.docx");
	}

	public void Ribbon_InsertTplGiaoAn5512(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Giao_an_KHBD_5512.docx");
	}

	public void Ribbon_InsertTplDonNghiPhep(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_xin_nghi_phep.docx");
	}

	public void Ribbon_InsertTplDonDeNghi(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_de_nghi.docx");
	}

	public void Ribbon_InsertTplDonKhieuNai(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_khieu_nai.docx");
	}

	public void Ribbon_InsertTplDonKienNghi(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_kien_nghi_phan_anh.docx");
	}

	public void Ribbon_InsertTplDonToCao(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_to_cao.docx");
	}

	public void Ribbon_InsertTplBanTuongTrinh(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Ban_tuong_trinh.docx");
	}

	public void Ribbon_InsertTplBanCamKet(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Ban_cam_ket.docx");
	}

	public void Ribbon_InsertTplGiayUyQuyenCaNhan(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Giay_uy_quyen_ca_nhan.docx");
	}

	public void Ribbon_InsertTplDonXinThoiViec(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_xin_thoi_viec.docx");
	}

	public void Ribbon_InsertTplDonTuChung(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_Don_tu_chung.docx");
	}

	public void Ribbon_InsertTplND152_01(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_01_To_trinh_khen_thuong.docx");
	}

	public void Ribbon_InsertTplND152_02(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_02_Bao_cao_thanh_tich_Co_thi_dua_Chinh_phu.docx");
	}

	public void Ribbon_InsertTplND152_03(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_03_Bao_cao_thanh_tich_CSTDTOANQUOC.docx");
	}

	public void Ribbon_InsertTplND152_04(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_04_Bao_cao_thanh_tich_tap_the_cong_trang.docx");
	}

	public void Ribbon_InsertTplND152_05(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_05_Bao_cao_thanh_tich_ca_nhan.docx");
	}

	public void Ribbon_InsertTplND152_06(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_06_Bao_cao_tom_tat_khen_thuong_dot_xuat.docx");
	}

	public void Ribbon_InsertTplND152_07(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_07_Bao_cao_tom_tat_khen_thuong_cong_hien.docx");
	}

	public void Ribbon_InsertTplND152_08(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_08_Bao_cao_thanh_tich_chuyen_de.docx");
	}

	public void Ribbon_InsertTplND152_09(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_09_Bao_cao_thanh_tich_Anh_hung_tap_the.docx");
	}

	public void Ribbon_InsertTplND152_10(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_10_Bao_cao_thanh_tich_Anh_hung_ca_nhan.docx");
	}

	public void Ribbon_InsertTplND152_11(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_11_Bao_cao_thanh_tich_doi_ngoai_to_chuc.docx");
	}

	public void Ribbon_InsertTplND152_12(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_12_Bao_cao_thanh_tich_doi_ngoai_ca_nhan.docx");
	}

	public void Ribbon_InsertTplND152_13(IRibbonControl control)
	{
		OpenTemplateHelper("Mau_ND152_13_Xac_nhan_sang_kien_kinh_nghiem.docx");
	}

	public void Ribbon_SelectHeaderDateSeparate(IRibbonControl control)
	{
		try
		{
			UserPersonalization userPersonalization = _profileRepo.LoadPersonalization() ?? new UserPersonalization();
			userPersonalization.HeaderDateInTopCell = false;
			_profileRepo.SavePersonalization(userPersonalization);
			if (_engine != null && _engine.RuleEngine != null)
			{
				_engine.RuleEngine.UpdatePersonalization(userPersonalization);
			}
			ShowModernInfo("Đã chuyển sang: Kiểu 1 - Địa danh ở 1 ô riêng (Mẫu chuẩn 2 dòng ngang hàng với Số ký hiệu).\r\nKhi chèn hoặc chuẩn hóa Tiêu đề đầu trang, Địa danh sẽ được đặt ở ô riêng biệt phía dưới.", "Bố Cục Tiêu Đề Đầu Trang");
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi khi thiết lập vị trí Địa danh: " + ex.Message, "OfficeFormat");
		}
	}

	public void Ribbon_SelectHeaderDateTopCell(IRibbonControl control)
	{
		try
		{
			UserPersonalization userPersonalization = _profileRepo.LoadPersonalization() ?? new UserPersonalization();
			userPersonalization.HeaderDateInTopCell = true;
			_profileRepo.SavePersonalization(userPersonalization);
			if (_engine != null && _engine.RuleEngine != null)
			{
				_engine.RuleEngine.UpdatePersonalization(userPersonalization);
			}
			ShowModernInfo("Đã chuyển sang: Kiểu 2 - Địa danh nằm chung trong ô trên (ô dưới trống).\r\nKhi chèn hoặc chuẩn hóa Tiêu đề đầu trang, Địa danh sẽ được đặt chung ô trên với Quốc hiệu/Tên Đảng.", "Bố Cục Tiêu Đề Đầu Trang");
		}
		catch (Exception ex)
		{
			ShowModernError("Lỗi khi thiết lập vị trí Địa danh: " + ex.Message, "OfficeFormat");
		}
	}

	public void Ribbon_OpenPersonalize(IRibbonControl control)
	{
		ShowPersonalizationDialog();
	}


	public void Ribbon_ShowContactInfo(IRibbonControl control)
	{
		try
		{
			string msg = "THÔNG TIN TÁC GIẢ & HỖ TRỢ KỸ THUẬT OFFICEFORMAT\n\n" +
			             "👤 Tác giả phát triển: Hà Xuân Vũ\n" +
			             "📞 Số điện thoại / Zalo: 0865333079\n\n" +
			             "Mọi yêu cầu hỗ trợ kỹ thuật, hướng dẫn sử dụng hoặc đóng góp ý kiến nâng cấp phần mềm, xin vui lòng liên hệ qua SĐT / Zalo trên.\n\n" +
			             "(Số điện thoại 0865333079 đã được tự động sao chép vào bộ nhớ tạm của bạn)";
			try { System.Windows.Clipboard.SetText("0865333079"); } catch { }
			ShowModernInfo(msg, "OFFICEFORMAT - THÔNG TIN TÁC GIẢ & HỖ TRỢ");
		}
		catch
		{
			System.Windows.Forms.MessageBox.Show("Tác giả: Hà Xuân Vũ\nSĐT / Zalo: 0865333079", "OFFICEFORMAT - HỖ TRỢ");
		}
	}

	public void Ribbon_ShowVFormatInfo(IRibbonControl control)
	{
		ShowLicenseDialog();
	}

	public void Ribbon_CheckUpdate(IRibbonControl control)
	{
		string fileVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
		if (!System.Version.TryParse(fileVersion, out var result))
		{
			result = Assembly.GetExecutingAssembly().GetName().Version;
		}
		string text = ((result == null) ? "Không xác định" : result.ToString(3));
		try
		{
			if (string.IsNullOrWhiteSpace(TafinexClient.UpdateVersionUrl))
			{
				ShowModernInfo($"Bạn đang sử dụng OfficeFormat phiên bản v{text} (Phiên bản đầu tiên).\r\nHiện tại chưa cấu hình tên miền (Domain) máy chủ cập nhật.", "OfficeFormat " + text + " - Kiểm Tra Cập Nhật");
				return;
			}
			Dictionary<string, object> data = new TafinexClient().CheckVersion();
			if (!UpdateManifestVerifier.TryVerify(data, out var manifest, out var error))
			{
				ShowModernWarning("Không thể xác thực thông tin cập nhật từ máy chủ: " + error, "OfficeFormat " + text + " - Kiểm Tra Cập Nhật");
				return;
			}
			if (manifest.Version <= result)
			{
				ShowModernInfo($"Bạn đang dùng phiên bản mới nhất ({manifest.VersionText}). Không cần tải hoặc cập nhật lại.", "OfficeFormat " + manifest.VersionText + " - Đã Mới Nhất");
				return;
			}
			string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			string text2 = Path.Combine(directoryName, "VFormat.Updater.exe");
			if (!File.Exists(text2))
			{
				text2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VFormat.Updater.exe");
			}
			if (File.Exists(text2))
			{
				if (ShowModernConfirm($"Đã có OfficeFormat {manifest.VersionText} (bạn đang dùng {text}).\r\n\r\nĐể cập nhật, Word sẽ được yêu cầu đóng tất cả cửa sổ. Word vẫn hỏi lưu với tài liệu chưa lưu; hãy lưu công việc trước khi tiếp tục.\r\n\r\nCập nhật ngay?", "OfficeFormat " + manifest.VersionText + " - Cập Nhật", "Cập nhật ngay", "Để sau"))
				{
					int id = Process.GetCurrentProcess().Id;
					string fileName = UpdaterLauncher.PrepareTemporaryCopy(text2);
					ProcessStartInfo processStartInfo = new ProcessStartInfo();
					processStartInfo.FileName = fileName;
					processStartInfo.Arguments = $"--pid {id} --target \"{directoryName}\" --url \"{manifest.UpdateUrl}\" --sha256 \"{manifest.Sha256}\" --version \"{manifest.VersionText}\" --restart-word";
					processStartInfo.UseShellExecute = true;
					Process.Start(processStartInfo);
				}
			}
			else
			{
				ShowModernWarning("Không tìm thấy trình cập nhật VFormat.Updater.exe. Vui lòng tải bộ cài mới nhất từ trang chủ.", "OfficeFormat " + text + " - Kiểm Tra Cập Nhật");
			}
		}
		catch (Exception ex)
		{
			Trace.WriteLine("Ribbon_CheckUpdate Error: " + ex.Message);
			ShowModernError("Không thể kiểm tra cập nhật: " + ex.Message, "OfficeFormat " + text);
		}
	}

	public void ShowPersonalizationDialog()
	{
		try
		{
			EnsureWpfApplication();
			PersonalizationDialog personalizationDialog = new PersonalizationDialog();
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(personalizationDialog).Owner = mainWindowHandle;
			}
			personalizationDialog.ShowDialog();
			try
			{
				UserPersonalization userPersonalization = _profileRepo.LoadPersonalization();
				if (userPersonalization != null && _engine != null && _engine.RuleEngine != null)
				{
					_engine.RuleEngine.UpdatePersonalization(userPersonalization);
				}
			}
			catch
			{
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở hộp thoại Cá nhân hóa: " + ex.Message, "OfficeFormat");
		}
	}

	public void ShowLicenseDialog()
	{
		try
		{
			EnsureWpfApplication();
			LicenseDialog licenseDialog = new LicenseDialog();
			IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
			if (mainWindowHandle != IntPtr.Zero)
			{
				new WindowInteropHelper(licenseDialog).Owner = mainWindowHandle;
			}
			licenseDialog.ShowDialog();
			if (_taskPaneControl != null && _taskPaneControl.ViewModel != null)
			{
				_taskPaneControl.ViewModel.RefreshLicenseStatus();
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở hộp thoại Bản quyền: " + ex.Message);
		}
	}

	public void ShowSpellCheckDialog()
	{
		try
		{
			if (EnsureActiveDocument(out var doc))
			{
				EnsureWpfApplication();
				RulePackType rulePack = _selectedRulePack;
				string documentKey = GetDocumentKey(doc);
				if (!string.IsNullOrEmpty(documentKey) && _docSelectedRules.ContainsKey(documentKey))
				{
					rulePack = _docSelectedRules[documentKey];
				}
				SpellCheckWindow spellCheckWindow = new SpellCheckWindow((Microsoft.Office.Interop.Word.Application)_wordApp, (Document)doc, rulePack);
				IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
				if (mainWindowHandle != IntPtr.Zero)
				{
					new WindowInteropHelper(spellCheckWindow).Owner = mainWindowHandle;
				}
				spellCheckWindow.Show();
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở cửa sổ Kiểm tra chính tả: " + ex.Message);
		}
	}

	public void Ribbon_OpenSpellCheck(IRibbonControl control)
	{
		ShowSpellCheckDialog();
	}

	public void ShowSpacingCheckDialog()
	{
		try
		{
			if (EnsureActiveDocument(out var doc))
			{
				EnsureWpfApplication();
				RulePackType rulePack = _selectedRulePack;
				string documentKey = GetDocumentKey(doc);
				if (!string.IsNullOrEmpty(documentKey) && _docSelectedRules.ContainsKey(documentKey))
				{
					rulePack = _docSelectedRules[documentKey];
				}
				SpellCheckWindow spacingCheckWindow = new SpellCheckWindow((Microsoft.Office.Interop.Word.Application)_wordApp, (Document)doc, rulePack, spacingCheckMode: true);
				IntPtr mainWindowHandle = Process.GetCurrentProcess().MainWindowHandle;
				if (mainWindowHandle != IntPtr.Zero)
				{
					new WindowInteropHelper(spacingCheckWindow).Owner = mainWindowHandle;
				}
				spacingCheckWindow.Show();
			}
		}
		catch (Exception ex)
		{
			ShowModernError("Không thể mở cửa sổ Kiểm tra khoảng trắng & dấu cách thừa: " + ex.Message);
		}
	}

	public void Ribbon_OpenSpacingCheck(IRibbonControl control)
	{
		ShowSpacingCheckDialog();
	}

	[ComRegisterFunction]
	public static void RegisterFunction(Type t)
	{
		AddinRegistrar.RegisterPerUser(Assembly.GetExecutingAssembly().Location);
	}

	[ComUnregisterFunction]
	public static void UnregisterFunction(Type t)
	{
		AddinRegistrar.UnregisterPerUser();
	}

	public void Ribbon_InsertTplTT68_01(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M01_Nghi_quyet_QPPL.docx");
	}

	public void Ribbon_InsertTplTT68_02(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M02_Nghi_quyet_Ca_biet.docx");
	}

	public void Ribbon_InsertTplTT68_03a(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M03a_Quyet_dinh_truc_tiep_QPPL.docx");
	}

	public void Ribbon_InsertTplTT68_03b(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M03b_Quyet_dinh_truc_tiep_Ca_biet.docx");
	}

	public void Ribbon_InsertTplTT68_04a(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M04a_Quyet_dinh_ban_hanh_quy_che_QPPL.docx");
	}

	public void Ribbon_InsertTplTT68_04b(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M04b_Quyet_dinh_ban_hanh_quy_che_Ca_biet.docx");
	}

	public void Ribbon_InsertTplTT68_05(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M05_Chi_thi_QPPL.docx");
	}

	public void Ribbon_InsertTplTT68_06(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M06_Chi_thi_Ca_biet.docx");
	}

	public void Ribbon_InsertTplTT68_07(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M07_Quy_che_Quy_dinh.docx");
	}

	public void Ribbon_InsertTplTT68_08(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M08_Thong_tu.docx");
	}

	public void Ribbon_InsertTplTT68_09(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M09_Thong_tu_lien_tich.docx");
	}

	public void Ribbon_InsertTplTT68_10(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M10_Van_ban_co_ten_loai_Ke_hoach.docx");
	}

	public void Ribbon_InsertTplTT68_11(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M11_Cong_van_hanh_chinh.docx");
	}

	public void Ribbon_InsertTplTT68_12(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M12_Cong_dien_Hoa_toc.docx");
	}

	public void Ribbon_InsertTplTT68_13a(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M13a_Giay_moi_van_ban.docx");
	}

	public void Ribbon_InsertTplTT68_13b(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M13b_Giay_moi_phieu.docx");
	}

	public void Ribbon_InsertTplTT68_14(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M14_Giay_gioi_thieu.docx");
	}

	public void Ribbon_InsertTplTT68_15a(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M15a_Giay_nghi_phep.docx");
	}

	public void Ribbon_InsertTplTT68_15b(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M15b_Giay_nghi_phep_mau_thuong.docx");
	}

	public void Ribbon_InsertTplTT68_16(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M16_Phieu_gui.docx");
	}

	public void Ribbon_InsertTplTT68_17(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M17_Phieu_chuyen.docx");
	}

	public void Ribbon_InsertTplTT68_18(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M18_Phieu_bao.docx");
	}

	public void Ribbon_InsertTplTT68_19(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M19_Thu_cong.docx");
	}

	public void Ribbon_InsertTplTT68_20(IRibbonControl control)
	{
		OpenTemplateHelper("TT68_BQP/M20_Ban_sao_van_ban.docx");
	}

	private static string SanitizeRibbonId(string id)
	{
		if (string.IsNullOrEmpty(id))
		{
			return "item";
		}
		StringBuilder stringBuilder = new StringBuilder();
		foreach (char c in id)
		{
			if (char.IsLetterOrDigit(c) || c == '_')
			{
				stringBuilder.Append(c);
			}
		}
		if (stringBuilder.Length <= 0)
		{
			return "item";
		}
		return stringBuilder.ToString();
	}
}
