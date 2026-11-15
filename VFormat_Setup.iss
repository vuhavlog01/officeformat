; Script Inno Setup 7 - Đóng gói bộ cài đặt OfficeFormat (Add-in for Microsoft Word)
#define MyAppName "OfficeFormat"
#define MyAppVersion "1.0"
#define MyAppPublisher "Hà Xuân Vũ (0865.333.079)"
#define AppGuid "{{7F28B490-A833-4F85-B2A8-3E2A398246E1}"
#define TaskPaneGuid "{{6B68585D-7BF3-4C0E-9214-B9FD286E9CD5}"

[Setup]
AppId={#AppGuid}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppComments=Tiện ích chuẩn hóa thể thức văn bản hành chính Việt Nam theo Nghị định 30/2020/NĐ-CP & Hướng dẫn 05-HD/VPTW
DefaultDirName={localappdata}\Programs\OfficeFormat
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=D:\PM\ADD_IN\v-format\Output
OutputBaseFilename=OfficeFormat-1.0-Setup
SetupIconFile=D:\PM\ADD_IN\v-format\App_VFormat_2.2.9\favicon.ico
UninstallDisplayIcon={app}\favicon.ico
VersionInfoVersion=1.0.0.0
VersionInfoTextVersion=1.0
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
CloseApplications=yes
CloseApplicationsFilter=WINWORD.EXE

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "D:\PM\ADD_IN\v-format\App_VFormat_2.2.9\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Cài đặt & Khắc phục lỗi OfficeFormat"; Filename: "{app}\CaiDat_VFormat.bat"; IconFilename: "{app}\favicon.ico"
Name: "{group}\Gỡ bỏ cài đặt {#MyAppName}"; Filename: "{uninstallexe}"

[Registry]
; Đăng ký Add-in trực tiếp vào Microsoft Word của người dùng hiện tại (HKCU)
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VFormat.Addin"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeFormat 1.0"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: string; ValueName: "Description"; ValueData: "OfficeFormat 1.0 - Chuẩn hóa thể thức văn bản hành chính, Đảng bộ và Quân đội"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: 0

; Đăng ký cho Word 32-bit trên Windows 64-bit
Root: HKCU; Subkey: "Software\Wow6432Node\Microsoft\Office\Word\Addins\VFormat.Addin"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU; Subkey: "Software\Wow6432Node\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeFormat 1.0"; Check: IsWin64
Root: HKCU; Subkey: "Software\Wow6432Node\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: string; ValueName: "Description"; ValueData: "OfficeFormat 1.0 - Chuẩn hóa thể thức văn bản hành chính, Đảng bộ và Quân đội"; Check: IsWin64
Root: HKCU; Subkey: "Software\Wow6432Node\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3; Check: IsWin64
Root: HKCU; Subkey: "Software\Wow6432Node\Microsoft\Office\Word\Addins\VFormat.Addin"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: 0; Check: IsWin64

; Đăng ký COM Classes trong HKCU\Software\Classes
Root: HKCU; Subkey: "Software\Classes\VFormat.Addin"; Flags: uninsdeletekey; ValueType: string; ValueData: "VFormat.Addin.Addin"
Root: HKCU; Subkey: "Software\Classes\VFormat.Addin\CLSID"; Flags: uninsdeletekey; ValueType: string; ValueData: "{#AppGuid}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}"; Flags: uninsdeletekey; ValueType: string; ValueData: "VFormat.Addin.Addin"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\ProgId"; Flags: uninsdeletekey; ValueType: string; ValueData: "VFormat.Addin"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "Class"; ValueData: "VFormat.Addin.Addin"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "Assembly"; ValueData: "VFormat.Addin, Version=2.0.1.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/VFormat.Addin.dll"

; Đăng ký COM Class cho TaskPaneControl
Root: HKCU; Subkey: "Software\Classes\VFormat.Addin.TaskPaneControl"; Flags: uninsdeletekey; ValueType: string; ValueData: "VFormat.Addin.TaskPaneControl"
Root: HKCU; Subkey: "Software\Classes\VFormat.Addin.TaskPaneControl\CLSID"; Flags: uninsdeletekey; ValueType: string; ValueData: "{#TaskPaneGuid}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}"; Flags: uninsdeletekey; ValueType: string; ValueData: "VFormat.Addin.TaskPaneControl"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\ProgId"; Flags: uninsdeletekey; ValueType: string; ValueData: "VFormat.Addin.TaskPaneControl"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "Class"; ValueData: "VFormat.Addin.TaskPaneControl"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "Assembly"; ValueData: "VFormat.Addin, Version=2.0.1.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#TaskPaneGuid}\InprocServer32"; Flags: uninsdeletekey; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/VFormat.Addin.dll"

[Run]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File ""{app}\CaiDat_VFormat.ps1"""; Flags: runhidden
Filename: "winword.exe"; Description: "Mở Microsoft Word ngay để trải nghiệm OfficeFormat"; Flags: postinstall nowait skipifsilent unchecked

[UninstallRun]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File ""{app}\GoCaiDat_VFormat.ps1"""; Flags: runhidden; RunOnceId: "UninstallOfficeFormat"
