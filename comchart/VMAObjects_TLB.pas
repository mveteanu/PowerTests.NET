unit VMAObjects_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : $Revision:   1.88  $
// File generated on 29.05.2001 01:04:36 from Type Library described below.

// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
// ************************************************************************ //
// Type Lib: D:\t\VMA ASP Chart\VMAObjects.tlb (1)
// IID\LCID: {6BE22038-7AE6-46B7-AE87-FDB7B8AA70FD}\0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\System32\STDOLE2.TLB)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\STDVCL40.DLL)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls, StdVCL;

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  VMAObjectsMajorVersion = 1;
  VMAObjectsMinorVersion = 0;

  LIBID_VMAObjects: TGUID = '{6BE22038-7AE6-46B7-AE87-FDB7B8AA70FD}';

  IID_IASPChart: TGUID = '{9718CED5-0CF6-44CB-B9AE-702532DCF302}';
  CLASS_ASPChart: TGUID = '{E8906DB3-20A2-4438-85AA-40983D00F79B}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IASPChart = interface;
  IASPChartDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  ASPChart = IASPChart;


// *********************************************************************//
// Interface: IASPChart
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9718CED5-0CF6-44CB-B9AE-702532DCF302}
// *********************************************************************//
  IASPChart = interface(IDispatch)
    ['{9718CED5-0CF6-44CB-B9AE-702532DCF302}']
    procedure OnStartPage(const AScriptingContext: IUnknown); safecall;
    procedure OnEndPage; safecall;
    procedure DefineCanvas(const title: WideString; width: Integer; height: Integer); safecall;
    procedure AddBar(const name: WideString; value: Double); safecall;
    procedure GenerateChart; safecall;
    procedure About; safecall;
  end;

// *********************************************************************//
// DispIntf:  IASPChartDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {9718CED5-0CF6-44CB-B9AE-702532DCF302}
// *********************************************************************//
  IASPChartDisp = dispinterface
    ['{9718CED5-0CF6-44CB-B9AE-702532DCF302}']
    procedure OnStartPage(const AScriptingContext: IUnknown); dispid 1;
    procedure OnEndPage; dispid 2;
    procedure DefineCanvas(const title: WideString; width: Integer; height: Integer); dispid 3;
    procedure AddBar(const name: WideString; value: Double); dispid 4;
    procedure GenerateChart; dispid 5;
    procedure About; dispid 6;
  end;

// *********************************************************************//
// The Class CoASPChart provides a Create and CreateRemote method to          
// create instances of the default interface IASPChart exposed by              
// the CoClass ASPChart. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoASPChart = class
    class function Create: IASPChart;
    class function CreateRemote(const MachineName: string): IASPChart;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TASPChart
// Help String      : ASPChart Object
// Default Interface: IASPChart
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TASPChartProperties= class;
{$ENDIF}
  TASPChart = class(TOleServer)
  private
    FIntf:        IASPChart;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TASPChartProperties;
    function      GetServerProperties: TASPChartProperties;
{$ENDIF}
    function      GetDefaultInterface: IASPChart;
  protected
    procedure InitServerData; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IASPChart);
    procedure Disconnect; override;
    procedure OnStartPage(const AScriptingContext: IUnknown);
    procedure OnEndPage;
    procedure DefineCanvas(const title: WideString; width: Integer; height: Integer);
    procedure AddBar(const name: WideString; value: Double);
    procedure GenerateChart;
    procedure About;
    property  DefaultInterface: IASPChart read GetDefaultInterface;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TASPChartProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TASPChart
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TASPChartProperties = class(TPersistent)
  private
    FServer:    TASPChart;
    function    GetDefaultInterface: IASPChart;
    constructor Create(AServer: TASPChart);
  protected
  public
    property DefaultInterface: IASPChart read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

implementation

uses ComObj;

class function CoASPChart.Create: IASPChart;
begin
  Result := CreateComObject(CLASS_ASPChart) as IASPChart;
end;

class function CoASPChart.CreateRemote(const MachineName: string): IASPChart;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ASPChart) as IASPChart;
end;

procedure TASPChart.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{E8906DB3-20A2-4438-85AA-40983D00F79B}';
    IntfIID:   '{9718CED5-0CF6-44CB-B9AE-702532DCF302}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TASPChart.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IASPChart;
  end;
end;

procedure TASPChart.ConnectTo(svrIntf: IASPChart);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TASPChart.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TASPChart.GetDefaultInterface: IASPChart;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TASPChart.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TASPChartProperties.Create(Self);
{$ENDIF}
end;

destructor TASPChart.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TASPChart.GetServerProperties: TASPChartProperties;
begin
  Result := FProps;
end;
{$ENDIF}

procedure TASPChart.OnStartPage(const AScriptingContext: IUnknown);
begin
  DefaultInterface.OnStartPage(AScriptingContext);
end;

procedure TASPChart.OnEndPage;
begin
  DefaultInterface.OnEndPage;
end;

procedure TASPChart.DefineCanvas(const title: WideString; width: Integer; height: Integer);
begin
  DefaultInterface.DefineCanvas(title, width, height);
end;

procedure TASPChart.AddBar(const name: WideString; value: Double);
begin
  DefaultInterface.AddBar(name, value);
end;

procedure TASPChart.GenerateChart;
begin
  DefaultInterface.GenerateChart;
end;

procedure TASPChart.About;
begin
  DefaultInterface.About;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TASPChartProperties.Create(AServer: TASPChart);
begin
  inherited Create;
  FServer := AServer;
end;

function TASPChartProperties.GetDefaultInterface: IASPChart;
begin
  Result := FServer.DefaultInterface;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents('Servers',[TASPChart]);
end;

end.
