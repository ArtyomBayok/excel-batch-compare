; VBA key : 9D089266-3BC1-4175-B0FE-BA080F13925E
; Remove-Item -Path hklm:\SOFTWARE\Classes\0C89A0D8CC6CB
; Remove-Item -Path hklm:\SOFTWARE\ExcelCompare
; Remove-Item -Path "C:\Documents and Settings\All Users\Application Data\0C89A0D8CC6CB"

#define MyAppName "ExcelBatchCompare"
#define MyAppLongName "Excel Batch Compare"
#define MyAppVersion GetFileVersion(".\comp-lib\bin\Release\compare-lib.dll")
#define MyAppPublisher "Florent BREHERET"
;#define MyAppURL "http://code.google.com/p/excel-compare/"
     
[Setup]
AppId={#MyAppName}
PrivilegesRequired=poweruser
AppName={#MyAppLongName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
VersionInfoVersion={#MyAppVersion}
VersionInfoTextVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
;AppPublisherURL={#MyAppURL}
;AppSupportURL={#MyAppURL}
;AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppLongName}
DisableProgramGroupPage=yes
LicenseFile=.\License.txt
;InfoBeforeFile=.\ClassLibrary1\bin\Release\Info.txt
OutputDir="."
OutputBaseFilename=ExcelBatchCompareSetup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "{commonappdata}\{code:GetId}";  Flags: uninsneveruninstall

[Files]
Source: ".\references\isxdl.dll"; DestDir: {tmp}; Flags: deleteafterinstall
Source: ".\comp-lib\bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\comp-exe\bin\Release\*.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\comp-runner\bin\Release\*.exe"; DestDir: "{app}"; Flags: ignoreversion
;Source: ".\help\help.chm"; DestDir: "{app}"; Flags: ignoreversion
Source: ".\License.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Excel Compare"; Filename: "{app}\ExcelBatchCompare.exe"; WorkingDir: "{app}";
;Name: "{group}\Help"; Filename: "{app}\help.chm"; WorkingDir: "{app}";
;Name: "{group}\Project Home Page"; Filename: "http://code.google.com/p/excel-compare/"; WorkingDir: "{app}";
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

[Registry]
Root: HKLM; Subkey: "SOFTWARE\Classes\0{code:GetId}"; ValueType: string; ValueName: "t"; ValueData: {code:GetId} ; Permissions: everyone-modify; Flags: createvalueifdoesntexist uninsdeletekeyifempty
Root: HKLM; Subkey: "SOFTWARE\Classes\0{code:GetId}"; ValueType: string; ValueName: "l"; ValueData: "" ; Permissions: everyone-modify; Flags: createvalueifdoesntexist uninsdeletekeyifempty
Root: HKLM; Subkey: "SOFTWARE\ExcelCompare"; ValueType: string; ValueName: "InstallDate"; ValueData: {code:GetDate}; Flags: createvalueifdoesntexist uninsdeletekeyifempty
Root: HKLM; Subkey: "SOFTWARE\ExcelCompare"; ValueType: string; ValueName: "LicenseKey"; ValueData: ""; Flags: createvalueifdoesntexist uninsdeletekeyifempty
Root: HKLM; Subkey: "SOFTWARE\ExcelCompare"; ValueType: string; ValueName: "HostId"; ValueData: {code:GetId}; Flags: createvalueifdoesntexist uninsdeletekeyifempty

[Run] 
Filename: "{app}\{#MyAppName}.exe"; Flags: shellexec postinstall; Description: "Launch application"
                                                                                                                                                                                                                     
[UninstallDelete]
Type: filesandordirs; Name: "{app}"


[Code]

//---------------------------------------------------------------------------------------
// Download and install an exe
//---------------------------------------------------------------------------------------
procedure isxdl_AddFile(URL, Filename: PChar);
external 'isxdl_AddFile@files:isxdl.dll stdcall';
function isxdl_DownloadFiles(hWnd: Integer): Integer;
external 'isxdl_DownloadFiles@files:isxdl.dll stdcall';
function isxdl_SetOption(Option, Value: PChar): Integer;
external 'isxdl_SetOption@files:isxdl.dll stdcall';

function InstallSoftware( url: string; file: string; arguments: string; info: string; description: string ): Boolean;
  var dotnetRedistPath: string; hWnd: Integer; ResultCode: Integer; 
  begin
    dotnetRedistPath := ExpandConstant('{tmp}\' + file);
    if not FileExists(dotnetRedistPath) then begin
      isxdl_AddFile(url, dotnetRedistPath);
      hWnd := StrToInt(ExpandConstant('{wizardhwnd}'));
      isxdl_SetOption('label', info);
      isxdl_SetOption('description', description);
      if isxdl_DownloadFiles(hWnd) <> 0 then begin
        if Exec(ExpandConstant(dotnetRedistPath), arguments, '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then begin
           result := (ResultCode = 0);
        end;
      end;
    end;
  end;


//---------------------------------------------------------------------------------------
// Office PIA Installation
//---------------------------------------------------------------------------------------                                                                                                 
const o2003pia_url = 'http://download.microsoft.com/download/8/3/a/83a40b5a-5050-4940-bcc4-7943e1e59590/O2003PIA.EXE';
const o2007pia_url = 'http://download.microsoft.com/download/e/1/d/e1df4622-5f6c-4fb9-845b-38d009cc1188/PrimaryInteropAssembly.exe';
const o2010pia_url = 'http://download.microsoft.com/download/C/1/D/C1D6DBBB-700D-4669-98CF-820AC3AE8E55/PIARedist.exe';

function IsVersionInstalled(Version: String): Boolean;
  var ret: String;
  begin
    RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SOFTWARE\Classes\Excel.Application\CurVer', '', ret);
    Result := ret=Version;
  end;

function IsOffice64() : Boolean;
  begin
    if RegKeyExists(HKLM,'HKLM\Software\Microsoft\Office') and not RegKeyExists(HKLM,'Software\WOW6432Node\Microsoft\Office') then begin
       result:=true;
    end;
  end;

function IsOfficePIAInstalled(version : String) : Boolean;
  var Names: TArrayOfString; i: Integer; ret: boolean;
  begin
    result:=false;
    ret := RegGetValueNames(HKLM, 'SOFTWARE\Classes\Installer\Assemblies\Global', Names);
    if not ret then ret := RegGetValueNames(HKLM, 'SOFTWARE\WOW6432Node\Classes\Installer\Assemblies\Global', Names);
    if ret then begin  
      for i := 0 to GetArrayLength(Names)-1 do begin
        if Copy(Names[i],0,44)= ('Microsoft.Office.Interop.Excel,version="' + version) then begin
          result:=true;
          Break;
        end;
      end;
    end;
  end;

function IsOfficeInstalled(version : String) : Boolean;
  var success32: Boolean; success64: Boolean; value: Cardinal;
  begin
    success32 := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\Office\' + version + '\Common\InstallRoot', 'InstallCount', value);
    success64 := RegQueryDWordValue(HKLM, 'SOFTWARE\Wow6432Node\Microsoft\Office\' + version + '\Common\InstallRoot', 'InstallCount', value);
    result := (success32 Or success64);
  end;

function InstallOffice2003pia(checkOnly : Boolean) : Boolean;
  var  ResultCode: Integer;
  begin
    result := IsOfficeInstalled('11.0') and (not IsOfficePIAInstalled('11.0')) and (not IsOfficeInstalled('12.0')) and (not IsOfficeInstalled('14.0'));
    if result and (not checkOnly) then begin
      if InstallSoftware( o2003pia_url, 'O2003PIA.EXE', '/extract:. /quiet', 'Downloading Microsoft Office 2003 Primary Interop Assemblies', 'Please wait while Setup is downloading extra files to your computer.') then begin
        if ShellExec('', ExpandConstant('{tmp}\O2003PIA.MSI'), '', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then result:= (ResultCode <> 0);
      end;
    end;
  end;

function InstallOffice2007pia(checkOnly : boolean) : boolean;
  var  ResultCode: Integer;
  begin
    result := IsOfficeInstalled('12.0') and (not IsOfficePIAInstalled('12.0')) and (not IsOfficeInstalled('14.0'));
    if result and (not checkOnly) then begin
      if InstallSoftware( o2007pia_url, 'PrimaryInteropAssembly.exe', '/extract:. /quiet', 'Downloading Microsoft Office 2007 Primary Interop Assemblies', 'Please wait while Setup is downloading extra files to your computer.') then begin
        if ShellExec('', ExpandConstant('{tmp}\o2007pia.msi'), '', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then result:= (ResultCode <> 0);
      end;
    end;
  end;

function InstallOffice2010pia(checkOnly : boolean) : boolean;
  var  ResultCode: Integer; installPath: string;
  begin
    result := IsOfficeInstalled('14.0') and (not IsOfficePIAInstalled('14.0'));
    if result and (not checkOnly) then begin
      if InstallSoftware( o2010pia_url, 'PIARedist.exe', '/extract:. /quiet', 'Downloading Microsoft Office 2010 Primary Interop Assemblies', 'Please wait while Setup is downloading extra files to your computer.') then begin
        if ShellExec('', ExpandConstant('{tmp}\o2010pia.msi'), '', '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then result:= (ResultCode <> 0);
      end;
    end;
  end;

//---------------------------------------------------------------------------------------
// NET Framework Installation
//---------------------------------------------------------------------------------------
const dotnetRedistURLx64 = 'http://download.microsoft.com/download/c/6/e/c6e88215-0178-4c6c-b5f3-158ff77b1f38/NetFx20SP2_x64.exe';
const dotnetRedistURLx86 = 'http://download.microsoft.com/download/c/6/e/c6e88215-0178-4c6c-b5f3-158ff77b1f38/NetFx20SP2_x86.exe';

function InstallNetFramework(checkOnly : boolean) : Boolean;
  var dotnetRedistPath: string; hWnd: Integer; ResultCode: Integer; 
  begin
    result := not (RegKeyExists(HKLM,'SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727') Or RegKeyExists(HKLM,'SOFTWARE\Wow6432Node\Microsoft\NET Framework Setup\NDP\v2.0.50727'));
    if result and (not checkOnly) then begin 
      if (not IsAdminLoggedOn()) then begin
        MsgBox('Microsoft .NET Framework 2.0 needs to be installed by an Administrator', mbInformation, MB_OK);
      end else begin
         if IsWin64() then begin
            result:= InstallSoftware( dotnetRedistURLx64, 'NetFx20SP2_x64.exe', '', 'Downloading Microsoft .NET Framework 2.0', 'Please wait while Setup is downloading extra files to your computer.') = false;
         end else begin
            result:= InstallSoftware( dotnetRedistURLx86, 'NetFx20SP2_x86.exe', '', 'Downloading Microsoft .NET Framework 2.0', 'Please wait while Setup is downloading extra files to your computer.') = false;
         end;
      end
      if InstallNetFramework(true) then begin
        MsgBox(ExpandConstant('Failed to install Microsoft .NET Framework 2.0 Service Pack 2 !  '#13'Please download and install it to continue the installaton'), mbError, MB_OK);
        ShellExec('open', 'http://www.microsoft.com/en-us/download/details.aspx?id=1639','', '', SW_SHOW, ewNoWait, ResultCode);
        result:= true;
      end
    end;
  end;
                
//---------------------------------------------------------------------------------------
// Uninstall previous version
//---------------------------------------------------------------------------------------
function UnInstallPrevious(checkOnly: boolean): boolean;
  var sUnInstallString: String; iResultCode: Integer;
  begin
    result:= RegQueryStringValue(HKLM, ExpandConstant('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppName}_is1'), 'UnInstallString', sUnInstallString);
    if not result then result:= RegQueryStringValue(HKLM, ExpandConstant('SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppName}_is1'), 'UnInstallString', sUnInstallString);
    if result and not checkOnly then Begin
        Exec( RemoveQuotes(sUnInstallString), '/SILENT', '', SW_SHOW, ewWaitUntilTerminated, iResultCode) ;
        if iResultCode <> 0 then Abort();
        Sleep(1000);
    end;
  end;

//---------------------------------------------------------------------------------------
// Get machine id
//---------------------------------------------------------------------------------------
function GetId(Param: String): string;
  var org: String; host: String; date: Cardinal;
  Begin
    RegQueryStringValue(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Windows NT\CurrentVersion', 'RegisteredOrganization', org); 
    RegQueryStringValue(HKEY_LOCAL_MACHINE, 'SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName', 'ComputerName', host);
    RegQueryDWordValue(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Windows NT\CurrentVersion', 'InstallDate', date);
    Result := Uppercase(Copy(GetSHA1OfString( '3FFBBCECD32D' + org + host + IntToStr(date)),0,16));
  end;

//---------------------------------------------------------------------------------------
// Get Current date (yyyy/mm/dd)
//---------------------------------------------------------------------------------------
function GetDate(Param: String): string;
  Begin
    Result := GetDateTimeString('yyyy/mm/dd','/',#0);
  end;

//---------------------------------------------------------------------------------------
// Check the program is not used before uninstallation
//---------------------------------------------------------------------------------------
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  sInstallLib: String;
begin
  if CurUninstallStep = usUninstall  then begin
     sInstallLib := ExpandConstant('{app}\{#MyAppName}.exe' );
     if FileExists( sInstallLib ) then begin
       if not RenameFile( sInstallLib, sInstallLib ) then RaiseException(ExpandConstant('Uninstallation of {#MyAppName} is not possible as a program is currently using it.'#13'Close it and try again.'));
     end;
  end;
end;

//---------------------------------------------------------------------------------------
// Update list of todo
//---------------------------------------------------------------------------------------
function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  s: string;
begin
  s := s + 'Uninstall previous version: ' + NewLine;
  if UnInstallPrevious(true) then s := s + Space + 'Found a previous installation to uninstall' + NewLine;
  s:= s + NewLine;
  s:= s + 'Dependencies to download and install:' + NewLine;
  if InstallNetFramework(true) then s := s + Space + 'Microsoft .Net Framework 2' + NewLine;
  if InstallOffice2003pia(true) then s := s + Space + 'Office 2003 Primary Interop Assemblies (4.12 MB)' + NewLine;
  if InstallOffice2007pia(true) then s := s + Space + 'Office 2007 Primary Interop Assemblies (6.28 MB)' + NewLine;
  if InstallOffice2010pia(true) then s := s + Space + 'Office 2010 Primary Interop Assemblies (6.55 MB)' + NewLine;
  s := s + NewLine;
  s := s + MemoDirInfo + NewLine + NewLine;
  Result := s
end;

//---------------------------------------------------------------------------------------
// Uninstall previous version and install additional applications
//---------------------------------------------------------------------------------------
procedure CurStepChanged(CurStep: TSetupStep);
  begin
    if CurStep = ssInstall  then begin
      UnInstallPrevious(false);
      if InstallNetFramework(false) then Abort();
      if InstallOffice2003pia(false) then Abort();
      if InstallOffice2007pia(false) then Abort();
      if InstallOffice2010pia(false) then Abort();
    end else if CurStep = ssDone then begin

    end;
  end;





