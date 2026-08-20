; 星易查 Windows 安装包脚本（Inno Setup 6）
; 构建（需先跑完 PyInstaller，产物在 dist\星易查\）：
;   ISCC.exe packaging\星易查.iss
; 产物：dist\星易查-Setup.exe
;
; 中文界面语言包（ChineseSimplified.isl，官方未随 Inno Setup 附带）由
; CI / build_windows.bat 下载到本目录；缺失时自动回退英文界面，不影响安装功能。

#define MyAppName "星易查"
#define MyAppFullName "星易查 - 围串标风险识别分析系统"
#define MyAppVersion "1.0.0"
#define MyAppExeName "星易查.exe"

[Setup]
AppId={{9B3A9CB7-DAC7-44F4-B99D-A739014B2B5C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=星易查
UninstallDisplayName={#MyAppFullName}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=星易查-Setup
SetupIconFile=star.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
; 避免杀软/兼容性提示：明确声明无签名，不请求在线检查
DisableReadyMemo=no

[Languages]
#if FileExists(AddBackslash(SourcePath) + "ChineseSimplified.isl")
Name: "chinese"; MessagesFile: "ChineseSimplified.isl"
#else
Name: "default"; MessagesFile: "compiler:Default.isl"
#endif

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; GroupDescription: "附加任务:"

[Files]
Source: "..\dist\星易查\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  // 安装完成后提示 LibreOffice（.doc 正文解析的可选依赖）
  if (CurStep = ssPostInstall) and
     not FileExists(ExpandConstant('{autopf64}\LibreOffice\program\soffice.exe')) and
     not FileExists(ExpandConstant('{autopf32}\LibreOffice\program\soffice.exe')) then
    MsgBox(
      '提示：未检测到 LibreOffice。' + #13#10#13#10 +
      '解析 .doc 格式标书的【正文】时需要 LibreOffice（免费，官网 libreoffice.org 可下载）。' + #13#10 +
      '未安装时将自动跳过 .doc 正文提取，.doc 元数据比对不受影响；' + #13#10 +
      'docx / pdf / 扫描件OCR / xlsx / txt 等功能完全不受影响。' + #13#10#13#10 +
      '需要时可稍后自行安装 LibreOffice，安装后无需重新安装本软件，重启即可生效。',
      mbInformation, MB_OK);
end;
