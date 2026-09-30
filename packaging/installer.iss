; installer.iss - Installateur Windows de "Automatisation Excel AYA" (Inno Setup 6)
;
; - installation SANS droits administrateur (dans le profil de l'utilisateur)
; - raccourcis Bureau + menu Demarrer, desinstallateur dans "Applications installees"
; - mise a jour : relancer un installateur plus recent -> remplace le programme,
;   ne touche JAMAIS a la config (%APPDATA%\AYA Excel) ni aux fichiers Excel.
;
; Construit par build_windows.ps1 :  ISCC /DAppVersion=1.1.0 /DSourceDir=... installer.iss

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\dist\AYA_Excel"
#endif
#ifndef OutputDir
  #define OutputDir "..\release"
#endif

#define AppName "Automatisation Excel AYA"
#define AppExe  "AYA_Excel.exe"
#define AppAuthor "Achraf Bengarin"

[Setup]
; AppId FIXE : c'est lui qui permet la mise a jour "par-dessus". Ne jamais le changer.
AppId={{CB0C2B5A-B86D-40AE-9C36-805123AF172D}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion} - par {#AppAuthor}
AppPublisher={#AppAuthor}
AppCopyright=(c) 2026 {#AppAuthor}
VersionInfoCompany={#AppAuthor}
VersionInfoCopyright=(c) 2026 {#AppAuthor}
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\AYA Excel
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
OutputDir={#OutputDir}
OutputBaseFilename=AYA_Excel_Setup_{#AppVersion}
SetupIconFile=..\assets\aya.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
AppComments=Developpe par {#AppAuthor}
CloseApplications=yes
RestartApplications=no
; Logiciel ouvert pendant une mise a jour -> "DeleteFile code 5 Acces refuse" :
; gere dans [Code] (FermerLogiciel) AVANT de toucher aux fichiers.

[Languages]
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Creer un raccourci sur le Bureau"; GroupDescription: "Raccourcis :"

[InstallDelete]
; mise a jour propre : on enleve l'ancien moteur avant de copier le nouveau
; (uniquement le dossier du programme ; la config utilisateur est ailleurs)
Type: filesandordirs; Name: "{app}\_internal"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{userdocs}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{userdocs}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Lancer {#AppName}"; Flags: nowait postinstall skipifsilent

; La desinstallation supprime le dossier du programme seulement.
; La config (%APPDATA%\AYA Excel) et les journaux restent : ce sont des donnees utilisateur.

[Code]
{ Logiciel encore ouvert (y compris une ancienne version sans AppMutex) ?
  Sinon Windows refuse de remplacer AYA_Excel.exe : "DeleteFile a echoue ; code 5". }
function ProcessusOuvert(const Nom: String): Boolean;
var
  Locator, Service, Liste: Variant;
begin
  Result := False;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('', 'root\CIMV2', '', '');
    Liste := Service.ExecQuery(Format('SELECT ProcessId FROM Win32_Process WHERE Name="%s"', [Nom]));
    Result := Liste.Count > 0;
  except
    Log('Detection du logiciel ouvert impossible : ' + GetExceptionMessage);
  end;
end;

function LogicielOuvert: Boolean;
begin
  Result := ProcessusOuvert('AYA_Excel.exe') or ProcessusOuvert('AYA_Excel_CLI.exe');
end;

{ Ferme toutes les fenetres du logiciel. '' = OK, sinon message d'erreur. }
function FermerLogiciel: String;
var
  Code, Essai: Integer;
begin
  Result := '';
  if not LogicielOuvert then
    exit;
  Log('Automatisation Excel AYA est ouvert');
  if SuppressibleMsgBox('Automatisation Excel AYA est encore ouvert.' + #13#10#13#10 +
       'Cliquez OK pour le fermer automatiquement et continuer' + #13#10 +
       '(vos fichiers Excel ne sont pas touches).',
       mbConfirmation, MB_OKCANCEL, IDOK) <> IDOK then
  begin
    Result := 'Fermez Automatisation Excel AYA, puis recommencez.';
    exit;
  end;
  { plusieurs fenetres ouvertes possibles : on recommence jusqu'a ce que tout soit ferme }
  for Essai := 1 to 10 do
  begin
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM AYA_Excel.exe', '', SW_HIDE, ewWaitUntilTerminated, Code);
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM AYA_Excel_CLI.exe', '', SW_HIDE, ewWaitUntilTerminated, Code);
    Sleep(1000);
    if not LogicielOuvert then
    begin
      Log(Format('Logiciel ferme (essai %d)', [Essai]));
      exit;
    end;
  end;
  Result := 'Impossible de fermer Automatisation Excel AYA.' + #13#10 +
            'Fermez-le (ou redemarrez le PC), puis recommencez.';
end;

{ Installation / mise a jour : on s'arrete AVANT de toucher aux fichiers si besoin. }
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := FermerLogiciel;
end;

{ Desinstallation : meme verification. }
function InitializeUninstall: Boolean;
var
  Erreur: String;
begin
  Erreur := FermerLogiciel;
  Result := Erreur = '';
  if not Result then
    SuppressibleMsgBox(Erreur, mbError, MB_OK, IDOK);
end;
