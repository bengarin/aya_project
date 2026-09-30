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
