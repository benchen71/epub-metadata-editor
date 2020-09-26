[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{C093DB2B-F7B8-4E5A-8E98-626F5486A44B}
AppName=EPUB Metadata Editor
AppVerName=EPUB Metadata Editor 1.8.3
VersionInfoVersion=1.8.3
AppPublisher=Ben Chenoweth
AppPublisherURL=https://github.com/benchen71/epub-metadata-editor
AppSupportURL=https://github.com/benchen71/epub-metadata-editor/wiki
AppUpdatesURL=https://github.com/benchen71/epub-metadata-editor/releases

; Select destination directory depending on Windows version
DefaultDirName={reg:HKCU\Software\EPubMetadataEditor,Path|{pf}\EPubMetadataEditor}

DefaultGroupName=EPubMetadataEditor
UninstallDisplayIcon={app}\EPubMetadataEditor.exe
AllowNoIcons=yes
LicenseFile=EPubMetadataEditor License.txt
OutputBaseFilename=EPubMetadataEditorInstaller
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
Source: EPubMetadataEditor License.txt; DestDir: {app}; Flags: ignoreversion
Source: btn_donateCC_LG.bmp; Flags: dontcopy
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

[Code]
procedure BitmapImageOnClick(Sender: TObject);
var
  ErrorCode: Integer;
begin
  ShellExecAsOriginalUser('open', 'https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=KC9T4JCJ2MPZG', '', '', SW_SHOWNORMAL, ewNoWait, ErrorCode);
end;

procedure CreateTheWizardPages;
var
  Page: TWizardPage;
  Label1, Label2: TLabel;
  BitmapImage: TBitmapImage;
  BitmapFileName: String;
begin
  Page := CreateCustomPage(wpInstalling, 'Installation Complete', 'Setup was completed successfully');

  Label1 := TLabel.Create(Page);
  Label1.Caption := 'Thank you for installing EPubMetadataEditor.';
  Label1.Left := 10;
  Label1.Top := 10;
  Label1.Parent := Page.Surface;

  Label2 := TLabel.Create(Page);
  Label2.Caption := 'Please consider donating, if you would like to show your appreciation for the program and support future development.  Donations are handled by PayPal.';
  Label2.Left := 10;
  Label2.Top := 40;
  Label2.Width := 400;
  Label2.WordWrap := true;
  Label2.Parent := Page.Surface;

  BitmapFileName := ExpandConstant('{tmp}\btn_donateCC_LG.bmp');
  ExtractTemporaryFile(ExtractFileName(BitmapFileName));

  BitmapImage := TBitmapImage.Create(Page);
  BitmapImage.AutoSize := True;
  BitmapImage.Bitmap.LoadFromFile(BitmapFileName);
  BitmapImage.Left := 160;
  BitmapImage.Top := 100;
  BitmapImage.Cursor := crHand;
  BitmapImage.OnClick := @BitmapImageOnClick;
  BitmapImage.Parent := Page.Surface;
end;

procedure InitializeWizard();
begin
  CreateTheWizardPages;  
end;
