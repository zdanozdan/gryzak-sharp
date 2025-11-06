; Inno Setup Script dla Gryzak - Menedżer Zamówień
; Aby użyć tego skryptu, zainstaluj Inno Setup z https://innosetup.com/

#define MyAppName "Gryzak"
#define MyAppVersion "1.5.0"
#define MyAppPublisher "Mikran sp. z o.o."
#define MyAppExeName "Gryzak.exe"
#define MyAppId "{{B8F3D4A1-2E5C-4F9A-8B6D-1C3E5F7A9B2C}"

[Setup]
; AppId={{{AppId}}
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=
OutputDir=installer
OutputBaseFilename=Gryzak-Setup-{#MyAppVersion}
SetupIconFile=gryzak.iconset\1-96da4866.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Uwaga: Foldery publish\win-x64 i publish\win-x86 muszą istnieć przed kompilacją instalatora
; Uruchom: .\publish.ps1 aby opublikować obie wersje, potem .\create-installer.ps1 aby utworzyć instalator
; Automatyczne wykrywanie architektury: x64 dla 64-bitowych systemów, x86 dla 32-bitowych
Source: "publish\win-x64\*"; DestDir: "{app}"; Check: Is64BitInstallMode; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "publish\win-x86\*"; DestDir: "{app}"; Check: not Is64BitInstallMode; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
  // Sprawdzenie czy pliki istnieją jest wykonywane podczas kompilacji
  // przez skrypt create-installer.ps1 przed uruchomieniem Inno Setup
end;

