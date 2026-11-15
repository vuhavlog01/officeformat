[Setup]
AppName=Tien Ich Dinh Dang Van Ban Word
AppVersion=1.0
DefaultDirName={autopf}\WordAddIn1
DefaultGroupName=WordAddIn1
OutputDir=Output
OutputBaseFilename=CaiDat_VFormat_WordAddIn
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest

[Files]
; Copy toàn bộ các file trong thư mục Release vào thư mục cài đặt
Source: "D:\PM\ADD_IN\v-format\WordAddIn1\bin\Release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Registry]
; Đăng ký Add-in trực tiếp vào Microsoft Word của người dùng hiện tại (HKCU)
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordAddIn1"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordAddIn1"; ValueType: string; ValueName: "Description"; ValueData: "Tien ich dinh dang van ban"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordAddIn1"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Tien ich Dinh Dang VFormat"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordAddIn1"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\WordAddIn1"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\WordAddIn1.vsto|vstolocal"