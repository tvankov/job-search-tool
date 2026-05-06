#define AppName      "Job Search Tool"
#define AppVersion   "1.0"
#define AppPublisher "Todor Vankov"
#define AppURL       "https://www.todorvankov.com"
#define AppExeName   "JobSearchTool.exe"

[Setup]
AppId={{B3F2A1C4-7E89-4D56-9A23-1F6B8C0D5E72}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=LICENSE.txt
OutputDir=installer_output
OutputBaseFilename=JobSearchTool_Setup_v{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";   Description: "Create a desktop shortcut"
Name: "startmenuicon"; Description: "Create a Start Menu shortcut"

[Files]
Source: "dist\{#AppExeName}";   DestDir: "{app}"; Flags: ignoreversion
Source: "job_search_auto.py";   DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; Tasks: desktopicon
Name: "{autostartmenu}\{#AppName}";   Filename: "{app}\{#AppExeName}"; Tasks: startmenuicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\jobs"
Type: files;          Name: "{app}\config.json"
Type: files;          Name: "{app}\job_search.log"
