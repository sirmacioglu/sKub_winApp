[Setup]
AppId={{82RB-LEGION-SKUB-2026}}
AppName=sKub
AppVersion=1.0
AppPublisher=Arif Sýrmacýoðlu
AppPublisherURL=https://srmc.tr
DefaultDirName={autopf}\sKub
DefaultGroupName=sKub
AllowNoIcons=yes
OutputDir={userdesktop}
OutputBaseFilename=sKub_v1.0_Kurulum
SetupIconFile={#SourcePath}\sKub.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; BU SATIR ÇOK ÖNEMLÝ: Dosyalarý x86 yerine doðrudan "Program Files" içine atar
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin 

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#SourcePath}\dist\sKub.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\sKub.ico"; DestDir: "{app}"; Flags: ignoreversion
; commonpf artýk x86'ya deðil, doðrudan C:\Program Files klasörüne gider
Source: "{#SourcePath}\wkhtmltopdf\*"; DestDir: "{commonpf}\wkhtmltopdf"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\sKub"; Filename: "{app}\sKub.exe"; IconFilename: "{app}\sKub.ico"
Name: "{autodesktop}\sKub"; Filename: "{app}\sKub.exe"; Tasks: desktopicon; IconFilename: "{app}\sKub.ico"

[Run]
Filename: "{app}\sKub.exe"; Description: "{cm:LaunchProgram,sKub}"; Flags: nowait postinstall skipifsilent