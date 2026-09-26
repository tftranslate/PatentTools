; ===========================================
; PatentTools Setup Script for Inno Setup 7
; ===========================================

// Preprocessor variables for installer type selection
// Pass /DAllUsers=1 to ISCC.exe to build AllUsers version (requires admin)
// For User version (no UAC), pass /DIsUser=1 or omit completely
#ifdef IsUser
#define IsUserInstall
#else
#define IsAllUsersInstall
#endif

#define MyAppName "PatentTools"
#ifndef APP_VERSION
#error APP_VERSION preprocessor variable is required - pass /DAPP_VERSION=x.x.x to ISCC.exe
#endif
#define MyAppPublisher "Tobias Friedrich Ernst"
#define MyAppURL "https://github.com/tftranslate/PatentTools"

// Installer type constants (vary based on installer type)
#ifdef IsUserInstall
#define InstallPrivileges "lowest"
#define AppDataPath "{userappdata}"
#else
#define InstallPrivileges "admin"
#define AppDataPath "{commonappdata}"
#endif

[Setup]
; --- Core Identification (AppId - GENERATE ONCE, NEVER CHANGE) ---
AppId={{d87e7e79-1b46-4546-bea3-5b129489839e}
AppName={#MyAppName}
AppVersion={#APP_VERSION}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}

; --- Install Paths & Behavior ---
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=..\..\build
OutputBaseFilename=PatentTools_{#APP_VERSION}_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern

; --- Permissions & UI ---
; Privileges depend on build type (User vs AllUsers)
PrivilegesRequired={#InstallPrivileges}
ShowLanguageDialog=no
DisableDirPage=yes
DisableProgramGroupPage=yes
LicenseFile=..\..\LICENSE

; --- Uninstaller Configuration ---
; These settings ensure the app appears in Windows Apps & Features
UninstallDisplayIcon={#AppDataPath}\Microsoft\Word\STARTUP\PatentTools.dotm
UninstallDisplayName={#MyAppName}
VersionInfoVersion="{#APP_VERSION}"
Uninstallable=yes

; --- UI Settings ---
DisableWelcomePage=no
DisableFinishedPage=yes

[Messages]
; Custom messages
SetupAppTitle={#MyAppName} Installation
WelcomeLabel1=Do you want to install {#MyAppName}?
WizardReady={#MyAppName} for Microsoft Word
RestartWarning=Please save all changes in Microsoft Word before proceeding with the installation.
#ifdef IsUserInstall
InfoAfterNote=The Patent Tools ribbon will appear after a complete restart of Microsoft Word. (User installation)
#else
InfoAfterNote=The Patent Tools ribbon will appear after a complete restart of Microsoft Word. (System-wide installation)
#endif

[Files]
; Copy .dotm to Word Startup folder (path depends on installer type)
Source: "..\..\build\PatentTools.dotm"; \
DestDir: "{#AppDataPath}\Microsoft\Word\STARTUP"; \
Flags: ignoreversion restartreplace;

[Registry]
; Track installation metadata (for upgrade verification)
Root: HKCU; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; \
ValueType: string; ValueName: "Version"; ValueData: "{#APP_VERSION}"; \
Flags: uninsdeletevalue

Root: HKCU; Subkey: "Software\{#MyAppPublisher}\{#MyAppName}"; \
ValueType: string; ValueName: "StartupPath"; \
ValueData: "{#AppDataPath}\Microsoft\Word\STARTUP\"; \
Flags: uninsdeletevalue

[Icons]
; Uninstall shortcut in Start Menu
Name: "{group}\Uninstall {#MyAppName}"; \
Filename: "{uninstallexe}"

[Code]
var
  GlobalWordWasKilled: Boolean;


// --- Check whether one or more Microsoft Word processes are running ---
// Returns true if Word is confirmed running, false otherwise.
// Sets QueryFailed to true if query could not be performed.
function IsWordRunning(var QueryFailed: Boolean): Boolean;
var
  ResultCode: Integer;
  TempFile: String;
  Content: AnsiString;
  LSResult: Boolean;
begin
  QueryFailed := False;
  Result := False;
  
  TempFile := ExpandConstant('{tmp}\word_check.tmp');
  
  // /NH = no header, so output is only Word processes (or nothing)
  if Exec('cmd.exe', '/c tasklist.exe /NH /FI "IMAGENAME eq WINWORD.EXE" > "' + TempFile + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    if FileExists(TempFile) then
    begin
      // Read file content and check for WINWORD.EXE
      try
        LSResult := LoadStringFromFile(TempFile, Content);
        // If Content is not empty, Word is running
		if (Content <> '') and LSResult and (Pos('WINWORD.EXE', Content) > 0) then
		begin
		   Result := True;
		end
		else 
		begin
		   Result := False;
		end;
      except
        QueryFailed := True;
      end;
      DeleteFile(TempFile);
    end
    else
    begin
      QueryFailed := True;
    end;
  end
  else
  begin
    QueryFailed := True;
  end;
end;


// --- Get Word executable path from registry ---
function GetWordPath(): String;
var
  WordPath: String;
begin
  // Try HKLM first (most common)
  if RegQueryStringValue(HKEY_LOCAL_MACHINE,
       'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
       '', WordPath) then
  begin
    Result := WordPath;
    Log('Found Word in HKLM: ' + WordPath);
    Exit;
  end;

  // Try HKCU second
  if RegQueryStringValue(HKEY_CURRENT_USER,
          'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
          '', WordPath) then
  begin
    Result := WordPath;
    Log('Found Word in HKCU: ' + WordPath);
    Exit;
  end;

  // Try 64-bit Office path (Office 2013+)
  if RegQueryStringValue(HKEY_LOCAL_MACHINE,
          'SOFTWARE\Microsoft\Office\ClickToRun\InstallRoot\Path',
          '', WordPath) then
  begin
    Result := WordPath + 'root\Office16\WINWORD.EXE';
    Log('Found ClickToRun Office path: ' + Result);
    Exit;
  end;

  // Try 64-bit Office path (Office 2013+)
  if RegQueryStringValue(HKEY_LOCAL_MACHINE,
          'SOFTWARE\Microsoft\Office\ClickToRun',
          'InstallPath', WordPath) then
  begin
    Result := WordPath + 'root\Office16\WINWORD.EXE';
    Log('Found ClickToRun Office path: ' + Result);
    Exit;
  end;

  // Fallback to common installation paths (try each until we find one that exists)
  if FileExists(ExpandConstant('{commonpf}\Microsoft Office\root\Office16\WINWORD.EXE')) then
  begin
    Result := ExpandConstant('{commonpf}\Microsoft Office\root\Office16\WINWORD.EXE');
    Log('Found Word in Program Files: ' + Result);
    Exit;
  end;

  if FileExists(ExpandConstant('{commonpf32}\Microsoft Office\root\Office16\WINWORD.EXE')) then
  begin
    Result := ExpandConstant('{commonpf32}\Microsoft Office\root\Office16\WINWORD.EXE');
    Log('Found Word in Program Files (x86): ' + Result);
    Exit;
  end;
    
  if FileExists(ExpandConstant('{commonpf64}\Microsoft Office\root\Office16\WINWORD.EXE')) then
  begin
    Result := ExpandConstant('{commonpf64}\Microsoft Office\root\Office16\WINWORD.EXE');
    Log('Found Word in Program Files: ' + Result);
    Exit;
  end;

  // Last resort - use Windows directory (likely to fail)
  Result := ExpandConstant('{win}\WINWORD.EXE');
  Log('Fallback Word path: ' + Result);
end;



// --- Ensure Word is closed BEFORE any file operations ---
function InitializeSetup(): Boolean;
var
  MessageText: String;
  RetryCount: Integer;
  TaskKillExitCode: Integer;
  KillSuccess: Boolean;
  WordQueryFailed: Boolean;
  WordActuallyRunning: Boolean;
  QueryFailedTemp: Boolean;
begin
  Result := True;
  GlobalWordWasKilled := False;

  // Check if previous installation exists
  if RegValueExists(HKEY_CURRENT_USER,
       'Software\{#MyAppPublisher}\{#MyAppName}', 'Version') then
  begin
    MsgBox('A previous version of PatentTools is currently installed.' + #13#10 +
      'This setup will upgrade the existing installation.',
      mbInformation, MB_OK);
  end;

  // Check Word status BEFORE setup performs any file operation
  WordActuallyRunning := IsWordRunning(WordQueryFailed);
  
  // CRITICAL: If query failed OR Word is running, we must stop and ask user
  if WordQueryFailed or WordActuallyRunning then
  begin

    RetryCount := 0;
    KillSuccess := False;

    while (not KillSuccess) and (RetryCount < 3) do
    begin
	  if WordQueryFailed then
	  begin
	     MessageText := 
		   'Cannot determine if Microsoft Word is currently running.' + #13#10 + #13#10 +
		   'Please manually close Word and click OK to proceed or Cancle to abort setup.';
      end
	  else
	  begin		
        MessageText :=
          'Microsoft Word is currently running.' + #13#10 + #13#10 +
          'PatentTools cannot be updated while Word has the template open.' + #13#10 +
          'Please save all open documents.' + #13#10 + #13#10 +
          'Click OK to close Word automatically, or Cancel to abort setup.';
      end;

      case MsgBox(MessageText, mbConfirmation, MB_OKCANCEL) of
        IDCANCEL:
          begin
            Result := False;
            Exit;
          end;

        IDOK:
          begin
            Inc(RetryCount);

            if Exec(ExpandConstant('{sys}\taskkill.exe'),
                 '/IM WINWORD.EXE /F', '', SW_HIDE, ewWaitUntilTerminated, TaskKillExitCode) then
            begin
              Log(Format('taskkill attempt %d returned exit code %d.', [RetryCount, TaskKillExitCode]));
            end
            else
            begin
              Log(Format('Could not start taskkill. Windows error: %d.', [TaskKillExitCode]));
            end;

            Sleep(1500);

            // Re-check Word status (pass new variable to capture query result)
            QueryFailedTemp := False;
            if not IsWordRunning(QueryFailedTemp) then
            begin
              KillSuccess := True;
              GlobalWordWasKilled := True;
            end;
        end;
      end;
    end;

    if not KillSuccess then
    begin
      MsgBox('Microsoft Word could not be closed automatically.' + #13#10 + 'Please close all Word windows manually and start the setup again.', mbError, MB_OK);
      Result := False;
    end;
  end;
  // If we reach here: Word is confirmed NOT running (no query failure, not running)
  // Setup can proceed safely
end;

// --- Installation progress step handler ---
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  case CurStep of
    ssPostInstall:
      begin
        Log('We are in ssPostInstall now.');
        
        // Show custom dialog asking about restarting Word
        if GlobalWordWasKilled then
        begin
          if MsgBox('PatentTools has been installed successfully.' + #13#10 +
            'The Patent Tools ribbon will appear when you start Word.' + #13#10 + #13#10 +
            'Would you like to restart Microsoft Word now?',
            mbConfirmation, MB_YESNO) = IDYES then
          begin
            Sleep(500);
            Exec(GetWordPath(), '', '', SW_SHOWNORMAL, ewNoWait, ResultCode);
          end;
        end
        else
        begin
          // Word wasn't killed (wasn't running), just show success
          MsgBox('PatentTools has been installed successfully.' + #13#10 +
            'The Patent Tools ribbon will appear when you start Word.',
            mbInformation, MB_OK);
        end;
      end;
  end;
end;



// --- Custom uninstall cleanup ---
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  StartupPath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Determine startup path based on installer type
#ifdef IsUserInstall
    StartupPath := ExpandConstant('{userappdata}\Microsoft\Word\STARTUP');
#else
    StartupPath := ExpandConstant('{commonappdata}\Microsoft\Word\STARTUP');
#endif

    // Remove .dotm from startup folder
    if FileExists(StartupPath + '\PatentTools.dotm') then
      DeleteFile(StartupPath + '\PatentTools.dotm');

    // Clean registry entries completely (per-user, so always HKCU)
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER,
      'Software\{#MyAppPublisher}\{#MyAppName}');
  end;
end;

