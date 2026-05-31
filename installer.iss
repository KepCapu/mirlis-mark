[Setup]
AppId={{B7D8F56C-9C62-4F62-9B2D-4E6E7A1A8C11}
AppName=Mark
AppVersion=1.0.0
AppVerName=Mark 1.0.0
AppPublisher=LDN Tech
AppCopyright=© 2026 LDN Tech
DefaultDirName={commonpf}\Mark
ArchitecturesInstallIn64BitMode=x64compatible
DefaultGroupName=Mark
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=MarkSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
SetupIconFile=assets\mark_app.ico
UninstallDisplayIcon={app}\Mark.exe
VersionInfoCompany=LDN Tech
VersionInfoDescription=Mark — Система маркировки
VersionInfoProductName=Mark
VersionInfoVersion=1.0.0.0
VersionInfoCopyright=© 2026 LDN Tech

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе";

[Files]
Source: "dist\Mark\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "data_sources\products.xlsx"; DestDir: "{app}\data_sources"; Flags: ignoreversion
Source: "docs\Руководство_пользователя.pdf"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Mark"; Filename: "{app}\Mark.exe"
Name: "{group}\Инструкция"; Filename: "{app}\Руководство_пользователя.pdf"
Name: "{autodesktop}\Mark"; Filename: "{app}\Mark.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Mark.exe"; Description: "Запустить Mark"; Flags: nowait postinstall skipifsilent