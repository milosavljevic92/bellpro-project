#define MyAppName "BellPro"
#define MyAppVersion "3.3.0"
#define MyAppPublisher "Tecomatic"
#define MyAppURL "https://tecomatic.rs"
#define MyAppExeName "BellPro.exe"

[Setup]
AppId={{53BA1E13-DC8C-4FBD-A955-802F4953DCD6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\Tecomatic\{#MyAppName}
DefaultGroupName=Tecomatic\{#MyAppName}
AllowNoIcons=yes

SetupIconFile=.\source\icon.ico
OutputDir=.\setup\
OutputBaseFilename=BellPro_setup_3.3.0_x86
Compression=lzma
SolidCompression=yes  
LicenseFile=.\source\license.txt
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: ".\source\vbRuntime\VB6_Full_Runtime_Setup.exe"; Flags: dontcopy
Source: ".\source\dll\mscomm32.ocx"; DestDir: "{app}"; Flags: restartreplace ignoreversion regserver 32bit
Source: ".\source\dll\msdatgrd.ocx"; DestDir: "{app}"; Flags: restartreplace ignoreversion regserver 32bit
Source: ".\source\BellPro.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[RUN]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram}"; Flags: nowait postinstall skipifsilent

[Code]
function IsRuntimeInstalled: Boolean;
begin
  Result := False; 
end;

function PrepareToInstall(var NeedsRestart: Boolean): string;
var
  ExitCode: Integer;
begin
  ExtractTemporaryFile('VB6_Full_Runtime_Setup.exe');
  if not Exec(ExpandConstant('{tmp}\VB6_Full_Runtime_Setup.exe'), '/q', '', SW_SHOW, ewWaitUntilTerminated, ExitCode) then
    Result := 'Failed to install VB Runtime.';
end;

[UninstallDelete]
Type: filesandordirs; Name: "{app}"