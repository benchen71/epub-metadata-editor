[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{C093DB2B-F7B8-4E5A-8E98-626F5486A44B}
AppName=EPubMetadataEditor
AppVerName=EPubMetadataEditor 1.2.16
VersionInfoVersion=1.2.16
AppPublisher=Ben Chenoweth
AppPublisherURL=http://code.google.com/p/epub-metadata-editor/
AppSupportURL=http://code.google.com/p/epub-metadata-editor/wiki/FAQ
AppUpdatesURL=http://code.google.com/p/epub-metadata-editor/downloads/list

; Select destination directory depending on Windows version
DefaultDirName={reg:HKCU\Software\EPubMetadataEditor,Path|{pf}\EPubMetadataEditor}

DefaultGroupName=EPubMetadataEditor
UninstallDisplayIcon={app}\EPubMetadataEditor.exe
AllowNoIcons=yes
LicenseFile=EPubMetadataEditor License.txt
OutputBaseFilename=EPubMetadataEditor Setup 1.2.16
Compression=lzma
SolidCompression=yes
ChangesAssociations=yes

[Languages]
Name: english; MessagesFile: compiler:Default.isl

[Tasks]
Name: desktopicon; Description: {cm:CreateDesktopIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: quicklaunchicon; Description: {cm:CreateQuickLaunchIcon}; GroupDescription: {cm:AdditionalIcons}; Flags: unchecked
Name: associate; Description: &Associate EPUB files; GroupDescription: Other tasks:; Flags: unchecked

[Files]
Source: EPubMetadataEditor.exe; DestDir: {app}; Flags: ignoreversion
Source: EPubMetadataEditorConsole.exe; DestDir: {app}; Flags: ignoreversion
Source: Ionic.Zip.dll; DestDir: {app}; Flags: ignoreversion
Source: EPubMetadataEditor License.txt; DestDir: {app}; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: {group}\EPubMetadataEditor; Filename: {app}\EPubMetadataEditor.exe
Name: {commondesktop}\EPubMetadataEditor; Filename: {app}\EPubMetadataEditor.exe; Tasks: desktopicon
Name: {userappdata}\Microsoft\Internet Explorer\Quick Launch\EPubMetadataEditor; Filename: {app}\EPubMetadataEditor.exe; Tasks: quicklaunchicon

[Registry]
Root: HKCU; Subkey: Software\EPubMetadataEditor; Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: Software\EPubMetadataEditor; ValueType: string; ValueName: Path; ValueData: {app}; Flags: uninsdeletekey
Root: HKCR; SubKey: .epub; ValueType: string; ValueData: EPUB File; Flags: uninsdeletekey; Tasks: associate
Root: HKCR; SubKey: EPUB File; ValueType: string; ValueData: EPUB File; Flags: uninsdeletekey; Tasks: associate
Root: HKCR; SubKey: EPUB File\Shell\Open\Command; ValueType: string; ValueData: """{app}\EPubMetadataEditor.exe"" ""%1"""; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; SubKey: EPUB File\Shell\Extract; ValueType: string; ValueData: "Extract cover"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; SubKey: EPUB File\Shell\Extract\Command; ValueType: string; ValueData: """{app}\EPubMetadataEditorConsole.exe"" ""%1"" -extract cover"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCR; Subkey: EPUB File\DefaultIcon; ValueType: string; ValueData: {app}\EPubMetadataEditor.exe,0; Flags: uninsdeletevalue; Tasks: associate

[Run]
Filename: {app}\EPubMetadataEditor.exe; Description: {cm:LaunchProgram,EPubMetadataEditor}; Flags: nowait postinstall skipifsilent
