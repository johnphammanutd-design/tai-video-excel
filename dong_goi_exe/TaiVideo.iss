; Bộ cài "HT CREATIVE AGENT" (tên cũ: "Tải Video Excel") - biên dịch bằng Inno Setup 6 (DONG_GOI.bat tự gọi)
; AppId và thư mục cài GIỮ NGUYÊN như bản cũ để cài đè lên bản "Tải Video Excel" đang có trên máy.
#define TenApp "HT CREATIVE AGENT"
#define TenCu "Tải Video Excel"
#define PhienBan "1.16"

[Setup]
AppId={{6F3B5C0E-7B1D-4C55-9E21-0A4D8E2F71C3}
AppName={#TenApp}
AppVersion={#PhienBan}
AppPublisher=HT Creative
DefaultDirName={localappdata}\Programs\TaiVideoExcel
DefaultGroupName={#TenApp}
; Không dùng lại tên nhóm Start Menu của bản cũ (nếu không, biểu tượng mới vẫn nằm trong nhóm tên cũ)
UsePreviousGroup=no
DisableProgramGroupPage=yes
; Cài cho từng người dùng, KHÔNG cần quyền quản trị.
; Thư mục cài phải ghi được để tool tự cập nhật yt-dlp và lưu cookie.
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=HTCreativeAgent_Setup
SetupIconFile=..\icon.ico
UninstallDisplayIcon={app}\app\icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Tool chạy bằng pythonw (không có cửa sổ riêng) nên Windows không tự đóng được -> bộ cài tự tắt (xem [Code])
CloseApplications=no

[Tasks]
Name: "desktopicon"; Description: "Tạo biểu tượng trên Desktop"; GroupDescription: "Biểu tượng:"

[InstallDelete]
; Xóa biểu tượng mang tên cũ khi cài đè lên bản "Tải Video Excel"
Type: files; Name: "{userdesktop}\{#TenCu}.lnk"
Type: filesandordirs; Name: "{userprograms}\{#TenCu}"
Type: files; Name: "{app}\Mo Tai Video.bat"

[Files]
Source: "build\TaiVideo\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#TenApp}"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\app\giao_dien.py"""; WorkingDir: "{app}\app"; IconFilename: "{app}\app\icon.ico"
Name: "{group}\Hướng dẫn"; Filename: "{app}\app\HUONG_DAN.txt"
Name: "{userdesktop}\{#TenApp}"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\app\giao_dien.py"""; WorkingDir: "{app}\app"; IconFilename: "{app}\app\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\app\giao_dien.py"""; WorkingDir: "{app}\app"; Description: "Mở {#TenApp}"; Flags: nowait postinstall skipifsilent

[Code]
// Tắt bản tool đang chạy từ thư mục cài đặt (kể cả khi cửa sổ đã đóng nhưng tool chưa kịp thoát),
// nếu không bộ cài báo "unable to automatically close all applications".
procedure TatToolDangChay();
var
  R: Integer;
begin
  Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -ExecutionPolicy Bypass -Command "Get-Process | Where-Object { $_.Path -like ''' +
    ExpandConstant('{app}') + '\*'' } | Stop-Process -Force"',
    '', SW_HIDE, ewWaitUntilTerminated, R);
  Sleep(800);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  TatToolDangChay();
  Result := '';
end;

function InitializeUninstall(): Boolean;
begin
  TatToolDangChay();
  Result := True;
end;

[UninstallDelete]
; Xóa cả thư viện đã tự cập nhật, file tạm và cookie của người dùng
Type: filesandordirs; Name: "{app}"
