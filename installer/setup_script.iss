; Seevo Teacher Helper 安装脚本
; 使用 Inno Setup 编译：ISCC.exe setup_script.iss

#define MyAppName "Seevo Teacher Helper"
#define MyAppVersion "1.0"
#define MyAppPublisher "SeevoHelper"
#define MyAppExeName "SeevoTeacherHelper.exe"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName=D:\Second installation\SeevoTeacherHelper
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=.
OutputBaseFilename=SeevoTeacherHelper_Setup
Compression=lzma
SolidCompression=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "D:\Second installation\InnoSetup\Default.isl"

[Files]
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[Code]
#emit 'procedure EmptyFolderCheck;'
#emit 'var'
#emit '  AppDir: string;'
#emit 'begin'
#emit '  AppDir := ExpandConstant(''\{app\}'');'
#emit '  if DirExists(AppDir) then begin'
#emit '    if MsgBox(''The selected folder is not empty.'' + #13#10 +'
#emit '      ''It is recommended to choose an empty folder.'' + #13#10 +'
#emit '      #13#10 + ''Install to this folder anyway?'','
#emit '      mbConfirmation, MB_YESNO) = IDNO then begin'
#emit '      Abort;'
#emit '    end;'
#emit '  end;'
#emit 'end;'
#emit ''
#emit 'function NextButtonClick(CurPageID: Integer): Boolean;'
#emit 'begin'
#emit '  Result := True;'
#emit '  if CurPageID = wpSelectDir then begin'
#emit '    EmptyFolderCheck;'
#emit '  end;'
#emit 'end;'

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "运行 Seevo Teacher Helper"; Flags: postinstall nowait skipifsilent
