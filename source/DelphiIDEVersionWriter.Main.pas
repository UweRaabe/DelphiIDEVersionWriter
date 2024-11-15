unit DelphiIDEVersionWriter.Main;

interface

type
  TWizard = class
  private
  class var
    FInstance: TWizard;
  var
    FNotifierID: Integer;
    FPluginInfoID: Integer;
    function GetDescription: string;
    function GetTitle: string;
    function GetVersion: string;
  public
    constructor Create;
    destructor Destroy; override;
    class procedure CreateInstance;
    class procedure DestroyInstance;
    property Description: string read GetDescription;
    property Title: string read GetTitle;
    property Version: string read GetVersion;
  end;

procedure Register;

implementation

uses
  System.Generics.Collections, System.SysUtils, System.StrUtils, System.IOUtils, System.Classes,
  Xml.XMLIntf,
  PlatformAPI, ToolsAPI;

procedure Register;
begin
  TWizard.CreateInstance;
end;

{$IF RTLVersion < 30; Seattle}
function GetProductVersion(const AFileName: string; var AMajor, AMinor, ABuild: Cardinal): Boolean;
var
  FileName: string;
  InfoSize, Wnd: DWORD;
  VerBuf: Pointer;
  FI: PVSFixedFileInfo;
  VerSize: DWORD;
begin
  Result := False;
  // GetFileVersionInfo modifies the filename parameter data while parsing.
  // Copy the string const into a local variable to create a writeable copy.
  FileName := AFileName;
  UniqueString(FileName);
  InfoSize := GetFileVersionInfoSize(PChar(FileName), Wnd);
  if InfoSize <> 0 then
  begin
    GetMem(VerBuf, InfoSize);
    try
      if GetFileVersionInfo(PChar(FileName), Wnd, InfoSize, VerBuf) then
        if VerQueryValue(VerBuf, '\', Pointer(FI), VerSize) then
        begin
          AMajor := HiWord(FI.dwProductVersionMS);
          AMinor := LoWord(FI.dwProductVersionMS);
          ABuild := HiWord(FI.dwProductVersionLS);
          Result:= True;
        end;
    finally
      FreeMem(VerBuf);
    end;
  end;
end;
{$IFEND}

const
  cTitle = 'Delphi IDE Version Writer';
  cVersion = 'V1.0.0';
  cCopyRight = 'Copyright© 2023 by Uwe Raabe' + sLineBreak +
               'https://www.uweraabe.de/';

resourcestring
  SDescription = 'Writes some Delphi version into the dproj file';

type
  TToolsAPI = record
  public
    class function OTAAboutBoxServices: IOTAAboutBoxServices; static;
    class function OTAModuleServices: IOTAModuleServices; static;
    class function OTAProjectFileStorage: IOTAProjectFileStorage; static;
    class function OTAServices: IOTAServices; static;
    class function OTASplashScreenServices: IOTASplashScreenServices; static;
    class function OTAVersionSKUInfoService: IOTAVersionSKUInfoService; static;
  end;

type
  TWizardNotifier = class(TNotifierObject, IOTAIDENotifier)
  private
    procedure CheckProjectFile(const FileName: string);
    procedure Initialize;
  public
    constructor Create;
    procedure AfterCompile(Succeeded: Boolean);
    procedure BeforeCompile(const Project: IOTAProject; var Cancel: Boolean);
    procedure FileNotification(NotifyCode: TOTAFileNotification; const FileName: string; var Cancel: Boolean);
    procedure WriteDelphiInfo(const AFileName: string); overload;
    procedure WriteDelphiInfo(AProject: IOTAProject); overload;
  end;

type
  TProjectStorage = class
  private
    FProject: IOTAProject;
    FRootName: string;
    function CreateProjectStorageNode: IXMLNode;
    function ForceIdentNode(const Section, Ident: string): IXMLNode;
    function ForceProjectStorageNode: IXMLNode;
    function ForceSectionNode(const Section: string): IXMLNode;
    function GetIdentNode(const Section, Ident: string): IXMLNode;
    function GetProjectStorageNode: IXMLNode;
    function GetSectionNode(const Section: string): IXMLNode;
  public
    constructor Create(const ARootName: string; AProject: IOTAProject);
    procedure MarkModified;
    function ReadAttribute(const Section, Ident: string; const Default: Integer): Integer; overload;
    function ReadAttribute(const Section, Ident, Default: string): string; overload;
    function ReadString(const Section, Ident, Default: string): string;
    procedure WriteAttribute(const Section, Ident: string; Value: Integer); overload;
    procedure WriteAttribute(const Section, Ident, Value: string); overload;
    procedure WriteString(const Section, Ident, Value: string);
    property Project: IOTAProject read FProject;
    property RootName: string read FRootName;
  end;

constructor TWizard.Create;
begin
  inherited;
  FNotifierID := TToolsAPI.OTAServices.AddNotifier(TWizardNotifier.Create);
  FPluginInfoID := TToolsAPI.OTAAboutBoxServices.AddPluginInfo(Title, Description, 0);
end;

destructor TWizard.Destroy;
begin
  if FPluginInfoID > 0 then
    TToolsAPI.OTAAboutBoxServices.RemovePluginInfo(FPluginInfoID);
  if FNotifierID > 0 then
    TToolsAPI.OTAServices.RemoveNotifier(FNotifierID);

  inherited;
end;

class procedure TWizard.CreateInstance;
begin
  FInstance := TWizard.Create;
end;

class procedure TWizard.DestroyInstance;
begin
  FInstance.Free;
end;

function TWizard.GetDescription: string;
begin
  Result := SDescription + sLineBreak + sLineBreak + cCopyRight;
end;

function TWizard.GetTitle: string;
begin
  Result := cTitle + ' ' + Version;
end;

function TWizard.GetVersion: string;
var
  build: Cardinal;
  major: Cardinal;
  minor: Cardinal;
begin
  if GetProductVersion(GetModuleName(HInstance), major, minor, build) then begin
    Result := Format('V%d.%d.%d', [major, minor, build]); // do not localize
  end
  else begin
    Result := cVersion;
  end;
end;

constructor TWizardNotifier.Create;
begin
  inherited Create;
  Initialize;
end;

procedure TWizardNotifier.AfterCompile(Succeeded: Boolean);
begin
end;

procedure TWizardNotifier.BeforeCompile(const Project: IOTAProject; var Cancel: Boolean);
begin
end;

procedure TWizardNotifier.FileNotification(NotifyCode: TOTAFileNotification; const FileName: string; var Cancel:
    Boolean);
begin
  case NotifyCode of
    ofnFileOpened: CheckProjectFile(FileName);
  end;
end;

procedure TWizardNotifier.CheckProjectFile(const FileName: string);
begin
  if MatchText(TPath.GetExtension(FileName), ['.dproj', '.cbproj']) then begin // do not localize
    WriteDelphiInfo(FileName);
  end;
end;

procedure TWizardNotifier.Initialize;
var
  I: Integer;
  module: IOTAProject;
begin
  { Handle open projects }
  if TToolsAPI.OTAModuleServices.MainProjectGroup <> nil then begin
    for I := 0 to TToolsAPI.OTAModuleServices.MainProjectGroup.ProjectCount - 1 do begin
      module := TToolsAPI.OTAModuleServices.MainProjectGroup.Projects[I];
      WriteDelphiInfo(module);
    end;
  end
  else begin
    module := TToolsAPI.OTAModuleServices.GetActiveProject;
    if module <> nil then begin
      WriteDelphiInfo(module);
    end;
  end;
end;

procedure TWizardNotifier.WriteDelphiInfo(const AFileName: string);
var
  project: IOTAProject;
begin
  project := TToolsAPI.OTAModuleServices.FindModule(AFileName) as IOTAProject;
  if project <> nil then
    WriteDelphiInfo(project)
end;

procedure TWizardNotifier.WriteDelphiInfo(AProject: IOTAProject);
const
  cProjectNodeName = 'DelphiInfo';
  cSectionNodeName = 'IDEVersion';
const
  cBool: array[Boolean] of Integer = (0, 1);
  cVersionNames: TArray<string> = [ //
    'Delphi 2007',                  // 18
    '',                             // 19
    'Delphi 2009',                  // 20
    'Delphi 2010',                  // 21
    'Delphi XE',                    // 22
    'Delphi XE2',                   // 23
    'Delphi XE3',                   // 24
    'Delphi XE4',                   // 25
    'Delphi XE5',                   // 26
    'Delphi XE6',                   // 27
    'Delphi XE7',                   // 28
    'Delphi XE8',                   // 29
    'Delphi 10 Seattle',            // 30
    'Delphi 10.1 Berlin',           // 31
    'Delphi 10.2 Tokyo',            // 32
    'Delphi 10.3 Rio',              // 33
    'Delphi 10.4%s Sydney',         // 34
    'Delphi 11%s Alexandria',       // 35
    'Delphi 12%s Athens',           // 36
    ''];
var
  lst: TStringList;
  majorVersion: Integer;
  minorVersion: Integer;
  platformId: TPlatformIds;
  skuInfo: IOTAVersionSKUInfoService;
  storage: TProjectStorage;
  subVersion: string;
  versionName: string;
begin
  majorVersion := Trunc(RTLVersion);
  minorVersion := 0;
  if majorVersion > 33 then begin
{$IF declared(RTLVersion1042) }
      minorVersion := 2;
{$IFEND}
{$IF declared(RTLVersion111) }
      minorVersion := 1;
{$IFEND}
{$IF declared(RTLVersion112) }
      minorVersion := 2;
{$IFEND}
{$IF declared(RTLVersion113) }
      minorVersion := 3;
{$IFEND}
{$IF declared(RTLVersion121) }
      minorVersion := 1;
{$IFEND}
{$IF declared(RTLVersion122) }
      minorVersion := 2;
{$IFEND}
    subVersion := '';
    if minorVersion > 0 then
      subVersion := '.' + minorVersion.ToString;
    versionName := Format(cVersionNames[majorVersion - 18], [subVersion]);
  end
  else begin
    versionName := cVersionNames[majorVersion - 18];
  end;
  skuInfo := TToolsAPI.OTAVersionSKUInfoService;
  storage := TProjectStorage.Create(cProjectNodeName, AProject);
  try
    storage.WriteAttribute(cSectionNodeName, 'NAME', versionName);
    storage.WriteAttribute(cSectionNodeName, 'VERSION', Format('%d.%d', [majorVersion, minorVersion]));
    storage.WriteAttribute(cSectionNodeName, 'SKU', skuInfo.SKU);
    storage.WriteAttribute(cSectionNodeName, 'TRIAL', cBool[skuInfo.IsProductTrial]);
    lst := TStringList.Create();
    try
      for platformId in skuInfo.Platforms do
        lst.Add(PlatformIDToName(platformId));
      storage.WriteString(cSectionNodeName, 'PLATFORMS', lst.CommaText);
    finally
      lst.Free;
    end;
  finally
    storage.Free;
  end;
end;

constructor TProjectStorage.Create(const ARootName: string; AProject: IOTAProject);
begin
  inherited Create;
  FRootName := ARootName;
  FProject := AProject;
end;

function TProjectStorage.CreateProjectStorageNode: IXMLNode;
begin
  Result := nil;
  if Project <> nil then
    Result := TToolsAPI.OTAProjectFileStorage.AddNewSection(Project, RootName, False);
end;

function TProjectStorage.ForceIdentNode(const Section, Ident: string): IXMLNode;
var
  sectionNode: IXMLNode;
begin
  Result := nil;
  sectionNode := ForceSectionNode(Section);
  if sectionNode = nil then Exit;
  Result := sectionNode.ChildNodes.FindNode(Ident);
  if Result = nil then
    Result := sectionNode.AddChild(Ident);
end;

function TProjectStorage.ForceProjectStorageNode: IXMLNode;
begin
  Result := GetProjectStorageNode;
  if Result = nil then
    Result := CreateProjectStorageNode;
end;

function TProjectStorage.ForceSectionNode(const Section: string): IXMLNode;
var
  storageNode: IXMLNode;
begin
  Result := nil;

  storageNode := ForceProjectStorageNode;
  if storageNode = nil then Exit;

  Result := storageNode.ChildNodes.FindNode(Section);
  if Result = nil then
    Result := storageNode.AddChild(Section);
end;

function TProjectStorage.GetIdentNode(const Section, Ident: string): IXMLNode;
var
  sectionNode: IXMLNode;
begin
  Result := nil;
  sectionNode := GetSectionNode(Section);
  if sectionNode <> nil then
    Result := sectionNode.ChildNodes.FindNode(Ident);
end;

function TProjectStorage.GetProjectStorageNode: IXMLNode;
begin
  Result := nil;
  if Project <> nil then
    Result := TToolsAPI.OTAProjectFileStorage.GetProjectStorageNode(Project, RootName, False);
end;

function TProjectStorage.GetSectionNode(const Section: string): IXMLNode;
var
  storageNode: IXMLNode;
begin
  Result := nil;
  storageNode := GetProjectStorageNode;
  if storageNode <> nil then
    Result := storageNode.ChildNodes.FindNode(Section);
end;

procedure TProjectStorage.MarkModified;
begin
  if Project <> nil then
    Project.MarkModified;
end;

function TProjectStorage.ReadAttribute(const Section, Ident: string; const Default: Integer): Integer;
var
  sectionNode: IXMLNode;
begin
  Result := Default;
  sectionNode := GetSectionNode(Section);
  if (sectionNode <> nil) and sectionNode.HasAttribute(Ident) then
    Result := sectionNode.Attributes[Ident];
end;

function TProjectStorage.ReadAttribute(const Section, Ident, Default: string): string;
var
  sectionNode: IXMLNode;
begin
  Result := Default;
  sectionNode := GetSectionNode(Section);
  if (sectionNode <> nil) and sectionNode.HasAttribute(Ident) then
    Result := sectionNode.Attributes[Ident];
end;

function TProjectStorage.ReadString(const Section, Ident, Default: string): string;
var
  identNode: IXMLNode;
begin
  Result := Default;
  identNode := GetIdentNode(Section, Ident);
  if identNode <> nil then
    Result := identNode.Text;
end;

procedure TProjectStorage.WriteAttribute(const Section, Ident: string; Value: Integer);
var
  sectionNode: IXMLNode;
begin
  sectionNode := ForceSectionNode(Section);
  if (sectionNode <> nil) then begin
    sectionNode.Attributes[Ident] := Value;
    MarkModified;
  end;
end;

procedure TProjectStorage.WriteAttribute(const Section, Ident, Value: string);
var
  sectionNode: IXMLNode;
begin
  sectionNode := ForceSectionNode(Section);
  if (sectionNode <> nil) then begin
    sectionNode.Attributes[Ident] := Value;
    MarkModified;
  end;
end;

procedure TProjectStorage.WriteString(const Section, Ident, Value: string);
var
  identNode: IXMLNode;
begin
  identNode := ForceIdentNode(Section, Ident);
  if identNode <> nil then begin
    identNode.Text := Value;
    MarkModified;
  end;
end;

class function TToolsAPI.OTAAboutBoxServices: IOTAAboutBoxServices;
begin
  BorlandIDEServices.GetService(IOTAAboutBoxServices, Result);
end;

class function TToolsAPI.OTAServices: IOTAServices;
begin
  BorlandIDEServices.GetService(IOTAServices, Result);
end;

class function TToolsAPI.OTAModuleServices: IOTAModuleServices;
begin
  BorlandIDEServices.GetService(IOTAModuleServices, Result);
end;

class function TToolsAPI.OTAProjectFileStorage: IOTAProjectFileStorage;
begin
  BorlandIDEServices.GetService(IOTAProjectFileStorage, Result);
end;

class function TToolsAPI.OTASplashScreenServices: IOTASplashScreenServices;
begin
  BorlandIDEServices.GetService(IOTASplashScreenServices, Result);
end;

class function TToolsAPI.OTAVersionSKUInfoService: IOTAVersionSKUInfoService;
begin
  BorlandIDEServices.GetService(IOTAVersionSKUInfoService, Result);
end;

initialization
finalization
  TWizard.DestroyInstance;
end.
