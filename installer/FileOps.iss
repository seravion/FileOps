; Inno Setup script for FileOps
#define MyAppName "FileOps"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "FileOps Team"
#define MyAppExeName "fileops.exe"

[Setup]
AppId={{4E6D0158-9FC6-4F59-BF8E-72F5F8D69372}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=FileOps-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "addpath"; Description: "Add FileOps to system PATH"; GroupDescription: "Additional tasks:"; Flags: unchecked

[Files]
Source: "..\dist\fileops.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Run FileOps"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if WizardIsTaskSelected('addpath') then
    begin
      Log('Selected PATH task. Add path manually or via enterprise policy if required.');
    end;
  end;
end;
