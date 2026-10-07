[Setup]
AppName=Simplewright
AppVersion=2.0
AppPublisher=Traphael
DefaultDirName={autopf}\Simplewright
DefaultGroupName=Simplewright
OutputDir=.
OutputBaseFilename=Simplewright_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
; Icona dell'installer
SetupIconFile=C:\Users\enzac\OneDrive\Desktop\simplewright\icon.ico

[Tasks]
Name: "desktopicon"; Description: "Crea un collegamento sul desktop"; GroupDescription: "Collegamenti aggiuntivi:"

[Files]
; 1. Copia l'intera cartella creata da PyInstaller (--onedir)
Source: "dist\simplewright\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 2. Copia l'icona dentro la cartella di installazione dell'app
Source: "C:\Users\enzac\OneDrive\Desktop\simplewright\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Run]
; 3. Opzione per avviare Simplewright al termine dell'installazione
Filename: "{app}\simplewright.exe"; Description: "Avvia Simplewright"; Flags: postinstall nowait skipifsilent

[Icons]
; 4. Collegamenti nel Menu Start e sul Desktop
Name: "{group}\Simplewright"; Filename: "{app}\simplewright.exe"; IconFilename: "{app}\icon.ico"
Name: "{group}\Disinstalla Simplewright"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Simplewright"; Filename: "{app}\simplewright.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon