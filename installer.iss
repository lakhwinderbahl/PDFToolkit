[Setup]
AppName=PDF Toolkit
AppVersion=1.0
DefaultDirName={pf}\PDF Toolkit
DefaultGroupName=PDF Toolkit
OutputDir=C:\Users\lsingh\Desktop\PDFtoolkitapp\dist
OutputBaseFilename=pdftoolkit-setup
Compression=lzma
SolidCompression=yes
SetupIconFile=C:\Users\lsingh\Desktop\PDFtoolkitapp\icon.ico

[Files]
Source: "C:\Users\lsingh\Desktop\PDFtoolkitapp\dist\pdftoolkitapp.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PDF Toolkit"; Filename: "{app}\pdftoolkitapp.exe"
Name: "{commondesktop}\PDF Toolkit"; Filename: "{app}\pdftoolkitapp.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
