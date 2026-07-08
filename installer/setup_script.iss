; Seevo Teacher Helper 安装脚本
; 使用 Inno Setup 编译：ISCC.exe setup_script.iss

#define MyAppName "Seevo Teacher Helper"
#define MyAppVersion "2.1"
#define MyAppPublisher "SeevoHelper"
#define MyAppExeName "SeevoTeacherHelper.exe"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\SeevoTeacherHelper
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=SeevoTeacherHelper_Setup
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "D:\Second installation\InnoSetup\Default.isl"

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Code]
#include "check_empty_folder.iss"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "运行 Seevo Teacher Helper"; Flags: postinstall nowait skipifsilent
