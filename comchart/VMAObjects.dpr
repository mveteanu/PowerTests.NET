library VMAObjects;

{%File 'ASPChart.asp'}

uses
  ComServ,
  VMAObjects_TLB in 'VMAObjects_TLB.pas',
  uaspchart in 'uaspchart.pas' {ASPChart: CoClass};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
