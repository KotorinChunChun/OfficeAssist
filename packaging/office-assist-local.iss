#ifndef Payload
  #error Payload is required
#endif
#ifndef OutputDirectory
  #error OutputDirectory is required
#endif
[Setup]
AppId=KotorinToys.OfficeAssist
AppName=OfficeAssist
AppVersion=0.1.0
DefaultDirName={localappdata}\Programs\OfficeAssist
PrivilegesRequired=lowest
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
MinVersion=10.0.22000
OutputDir={#OutputDirectory}
OutputBaseFilename=OfficeAssist-0.1.0-local-setup
Compression=lzma2/fast
SolidCompression=yes
CloseApplications=no
RestartApplications=no
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\OfficeAssist.exe,0
SetupIconFile=..\src\OfficeAssist\assets\OfficeAssist.ico
WizardStyle=modern
[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
[Files]
Source: "{#Payload}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{userprograms}\OfficeAssist"; Filename: "{app}\OfficeAssist.exe"
