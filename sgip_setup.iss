[Setup]
AppName=SGIP - Sistema de Gestion de Inventario para Plantas
AppVersion=1.0.0
AppPublisher=Corpoelec
AppPublisherURL=https://github.com/stopmjdev
DefaultDirName={pf}\SGIP
DefaultGroupName=SGIP
OutputDir=dist\installer
OutputBaseFilename=SGIP_Setup_v1.0.0_XP32
Compression=lzma2
SolidCompression=yes
; El instalador requiere Windows 7+.
; Para Windows XP: copiar SGIP.exe directamente (no necesita instalador).
ArchitecturesInstallIn64BitMode=

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el escritorio"; GroupDescription: "Iconos adicionales:"

[Files]
Source: "dist\SGIP_XP\SGIP.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\SGIP Inventario";   DestName: "SGIP"; Filename: "{app}\SGIP.exe"
Name: "{group}\Desinstalar SGIP";  Filename: "{uninstallexe}"
Name: "{commondesktop}\SGIP Inventario"; Filename: "{app}\SGIP.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\SGIP.exe"; Description: "Iniciar SGIP ahora"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
