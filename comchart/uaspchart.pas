unit uaspchart;

interface

uses
  ComObj, ActiveX, AspTlb, VMAObjects_TLB, StdVcl,
  Classes, Windows;

Type PBarRecord = ^TBarRecord;
     TBarRecord = record
                    name  : String;
                    value : Double;
                  end;

type
  TASPChart = class(TASPObject, IASPChart)
  protected
    procedure OnEndPage; safecall;
    procedure OnStartPage(const AScriptingContext: IUnknown); safecall;
    procedure DefineCanvas(const title: WideString; width, height: Integer);
      safecall;
    procedure AddBar(const name: WideString; value: Double); safecall;
    procedure GenerateChart; safecall;
    procedure About; safecall;
  private
    ChartWidth,ChartHeight : Integer;
    ChartTitle: String;
    ChartBarList : TList;
    procedure StreamToVariant (Stream : TMemoryStream; var v : OleVariant);
  end;

implementation

uses ComServ,JPEG, Chart, Series, Graphics, Controls,TeEngine;

procedure TASPChart.OnEndPage;
begin
  inherited OnEndPage;
end;

procedure TASPChart.OnStartPage(const AScriptingContext: IUnknown);
begin
  ChartBarList:=TList.Create;
  inherited OnStartPage(AScriptingContext);
end;

procedure TASPChart.DefineCanvas(const title: WideString; width,
  height: Integer);
begin
  ChartTitle := title;
  ChartWidth := width;
  ChartHeight:= height;
end;

procedure TASPChart.AddBar(const name: WideString; value: Double);
var br:PBarRecord;
begin
  New(br);
  br^.name  := name;
  br^.value := value;
  ChartBarList.Add(br);
end;

procedure TASPChart.StreamToVariant (Stream : TMemoryStream; var v : OleVariant);
var
  p : pointer;
begin
  v := VarArrayCreate ([0, Stream.Size - 1], varByte);
  p := VarArrayLock (v);
  Stream.Position := 0;
  Stream.Read (p^, Stream.Size);
  VarArrayUnlock (v);
end;

procedure TASPChart.GenerateChart;
var
  HorizBarSeries : THorizBarSeries;
  FChart         : TChart;
  MemStream      : TMemoryStream;
  FJPEG          : TJPEGImage;
  Bitmap         : TBitmap;
  Rect           : TRect;
  i              : Integer;
  br             : PBarRecord;
  msvar          : OleVariant;
begin
  FChart := TChart.Create(nil);
  Bitmap := TBitmap.Create;
  FJPEG := TJPEGImage.Create;
  MemStream  := TMemoryStream.Create;
  try
    FChart.Color := clWhite;
    FChart.BevelOuter := bvNone;
    FChart.Legend.Visible := False;

    Rect.Left := 0;
    Rect.Top := 0;
    Rect.Right := ChartWidth;
    Rect.Bottom := ChartHeight;

    HorizBarSeries := THorizBarSeries.Create(FChart);
    HorizBarSeries.BarStyle := bsRectGradient;
    HorizBarSeries.ParentChart := FChart;
    HorizBarSeries.ShowInLegend := False;
    HorizBarSeries.Marks.Style := smsValue;
    Randomize;

    with FChart do
    begin
      SeriesList.Clear;
      for i := 0 to ChartBarList.Count-1 do
      begin
        br:=ChartBarList.Items[i];
        HorizBarSeries.AddBar(br^.value , br^.name ,Random(2147483648)); //xxx = addbar in loc de add
      end;
      SeriesList.Add(HorizBarSeries);
      with Title do
      begin
        Font.Size := 10;
        Font.Color := clBlack;
        Text.Clear;
        Text.Add(ChartTitle);
      end;
    end;
    Bitmap := FChart.TeeCreateBitmap(clWhite, Rect);
    FJPEG.Assign(Bitmap);
    FJPEG.SaveToStream(MemStream);

    StreamToVariant(MemStream, msvar);
    Response.ContentType := 'image/jpeg';
    Response.BinaryWrite(msvar);
  finally
    FChart.Free;
    FJPEG.Free;
    Bitmap.Free;
  end;
end;

procedure TASPChart.About;
var mesaj:WideString;
begin
  mesaj:='<p><b>Chart ASP object</b><br>'+
         '<a href="http://vmasoft.hypermart.net">(c) VMA software</a></p>';
  Response.Write(mesaj);
end;


initialization
  TAutoObjectFactory.Create(ComServer, TASPChart, Class_ASPChart,
    ciMultiInstance, tmApartment);
end.
