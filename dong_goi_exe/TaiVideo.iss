; Bộ cài "Tải Video Excel" - biên dịch bằng Inno Setup 6 (DONG_GOI.bat tự gọi)
#define TenApp "Tải Video Excel"
#define PhienBan "1.10"

[Setup]
AppId={{6F3B5C0E-7B1D-4C55-9E21-0A4D8E2F71C3}
AppName={#TenApp}
AppVersion={#PhienBan}
AppPublisher=HT Creative
DefaultDirName={localappdata}\Programs\TaiVideoExcel
DefaultGroupName={#TenApp}
DisableProgramGroupPage=yes
; Cài cho từng người dùng, KHÔNG cần quyền quản trị.
; Thư mục cài phải ghi được để tool tự cập nhật yt-dlp và lưu cookie.
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=TaiVideo_Setup
SetupIconFile=..\icon.ico
UninstallDisplayIcon={app}\app\icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Tạo biểu tượng trên Desktop"; GroupDescription: "Biểu tượng:"

[Files]
Source: "build\TaiVideo\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#TenApp}"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\app\giao_dien.py"""; WorkingDir: "{app}\app"; IconFilename: "{app}\app\icon.ico"
Name: "{group}\Hướng dẫn"; Filename: "{app}\app\HUONG_DAN.txt"
Name: "{userdesktop}\{#TenApp}"; Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\app\giao_dien.py"""; WorkingDir: "{app}\app"; IconFilename: "{app}\app\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\python\pythonw.exe"; Parameters: """{app}\app\giao_dien.py"""; WorkingDir: "{app}\app"; Description: "Mở {#TenApp}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Xóa cả thư viện đã tự cập nhật, file tạm và cookie của người dùng
Type: filesandordirs; Name: "{app}"
