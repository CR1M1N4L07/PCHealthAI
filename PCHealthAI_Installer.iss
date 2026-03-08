; PCHealthAI_Installer.iss
; Inno Setup 6 script for PC Health AI
; Created by CRiMiNAL

[Setup]
; Identity
AppId={{A3F7B2C1-4D8E-4F9A-B0C2-D1E2F3A4B5C6}
AppName=PC Health AI
AppVersion=1.0.0
AppPublisher=CRiMiNAL
AppPublisherURL=https://github.com/CRiMiNAL/PCHealthAI
AppSupportURL=https://github.com/CRiMiNAL/PCHealthAI/issues
AppUpdatesURL=https://github.com/CRiMiNAL/PCHealthAI/releases

; Install location
DefaultDirName={autopf}\PCHealthAI
DefaultGroupName=PC Health AI
AllowNoIcons=yes

; Output
OutputDir=installer_output
OutputBaseFilename=PCHealthAI-Setup-v1.0.0

; Compression
Compression=lzma2/ultra64
SolidCompression=yes

; Appearance
WizardStyle=modern
WizardSizePercent=120

; UAC and privileges
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=commandline

; Misc
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main exe (built by BUILD_ALL.bat)
Source: "dist\PCHealthAI.exe"; DestDir: "{app}"; Flags: ignoreversion

; App icon
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

; Config (empty - user fills in Settings)
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Icons]
Name: "{group}\PC Health AI"; Filename: "{app}\PCHealthAI.exe"
Name: "{group}\Uninstall PC Health AI"; Filename: "{uninstallexe}"
Name: "{autodesktop}\PC Health AI"; Filename: "{app}\PCHealthAI.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PCHealthAI.exe"; Description: "{cm:LaunchProgram,PC Health AI}"; Flags: nowait postinstall skipifsilent

[Code]
// Pre-install Windows version check
function InitializeSetup(): Boolean;
var
  Version: TWindowsVersion;
begin
  GetWindowsVersionEx(Version);
  if Version.Major < 10 then
  begin
    MsgBox('PC Health AI requires Windows 10 or later.', mbError, MB_OK);
    Result := False;
  end
  else
    Result := True;
end;

// Post-install reminder to enter API key
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssDone then
  begin
    MsgBox(
      'PC Health AI installed successfully!' + #13#10 + #13#10 +
      'To enable AI features:' + #13#10 +
      '1. Open PC Health AI' + #13#10 +
      '2. Click the Settings icon (top-right)' + #13#10 +
      '3. Enter your Anthropic API key' + #13#10 +
      '4. Click Save' + #13#10 + #13#10 +
      'Get a free API key at: console.anthropic.com',
      mbInformation, MB_OK
    );
  end;
end;
