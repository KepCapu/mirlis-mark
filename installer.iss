[Setup]
AppId={{B7D8F56C-9C62-4F62-9B2D-4E6E7A1A8C11}
AppName=Mirlis Mark
AppVersion=1.0.0
AppPublisher=KepCapu

DefaultDirName={commonpf}\Mirlis Mark
ArchitecturesInstallIn64BitMode=x64compatible
DefaultGroupName=Mirlis Mark
DisableProgramGroupPage=yes

OutputDir=installer
OutputBaseFilename=MirlisMarkSetup

Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

SetupIconFile=assets\mark_app.ico
UninstallDisplayIcon={app}\MirlisMark.exe

VersionInfoCompany=KepCapu
VersionInfoDescription=Mirlis Mark
VersionInfoProductName=Mirlis Mark
VersionInfoVersion=1.0.0.0
VersionInfoCopyright=KepCapu


[Tasks]
Name: "desktopicon"; Description: "Create desktop shortcut";


[Files]
Source: "dist\MirlisMark.exe"; DestDir: "{app}"; Flags: ignoreversion


[Icons]
Name: "{group}\Mirlis Mark"; Filename: "{app}\MirlisMark.exe"
Name: "{autodesktop}\Mirlis Mark"; Filename: "{app}\MirlisMark.exe"; Tasks: desktopicon


[Run]
Filename: "{app}\MirlisMark.exe"; Description: "Launch Mirlis Mark"; Flags: nowait postinstall skipifsilent