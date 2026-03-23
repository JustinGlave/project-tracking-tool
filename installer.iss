#define MyAppName "Project Tracking Tool"
#define MyAppPublisher "ATS Inc."
#define MyAppExeName "ProjectTrackingTool.exe"
#define MyAppVersion GetFileVersion("dist\ProjectTrackingTool\ProjectTrackingTool.exe")

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/JustinGlave/project-tracking-tool
AppSupportURL=https://github.com/JustinGlave/project-tracking-tool/issues
AppUpdatesURL=https://github.com/JustinGlave/project-tracking-tool/releases

; Install to LocalAppData so no admin rights are needed and the auto-updater works
DefaultDirName={localappdata}\ATS Inc\Project Tracking Tool
DefaultGroupName=ATS Inc\Project Tracking Tool
DisableProgramGroupPage=yes

; Output
OutputDir=dist
OutputBaseFilename=ProjectTrackingToolSetup
SetupIconFile=PTT_Normal.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

; No admin required (LocalAppData is user-writable)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline

; Compression
Compression=lzma2/ultra64
SolidCompression=yes

; Wizard appearance
WizardStyle=modern
WizardResizable=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Include the entire PyInstaller output folder
Source: "dist\ProjectTrackingTool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Desktop shortcut
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
; Start Menu
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
; Offer to launch after install
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up any files the app creates in its folder (logs, temp files)
Type: filesandordirs; Name: "{app}"

[Code]
// Ask user if they want to keep their data on uninstall
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  DataDir: String;
  MsgResult: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    DataDir := ExpandConstant('{userappdata}\ATS Inc\Project Tracking Tool');
    if DirExists(DataDir) then
    begin
      MsgResult := MsgBox(
        'Do you want to delete your project data?' + #13#10 +
        '(jobs, tasks, notes, and change orders)' + #13#10#13#10 +
        'Click Yes to delete all data, or No to keep it.',
        mbConfirmation, MB_YESNO
      );
      if MsgResult = IDYES then
        DelTree(DataDir, True, True, True);
    end;
  end;
end;
