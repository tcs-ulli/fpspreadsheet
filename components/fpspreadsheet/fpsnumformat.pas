unit fpsNumFormat;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

interface

uses
  Classes, SysUtils,
  fpstypes, fpspreadsheet;

type
  { TsNumFormatList }

  TsNumFormatList = class(TFPList)
  private
    FOwnsData: Boolean;
    function GetItem(AIndex: Integer): TsNumFormatParams;
    procedure SetItem(AIndex: Integer; const AValue: TsNumFormatParams);
  protected
    FWorkbook: TsWorkbook;
    FClass: TsNumFormatParamsClass;
    procedure AddBuiltinFormats; virtual;
  public
    constructor Create(AWorkbook: TsWorkbook; AOwnsData: Boolean);
    destructor Destroy; override;
    function AddFormat(ASections: TsNumFormatSections): Integer; overload;
    function AddFormat(AFormatStr: String): Integer; overload;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function Find(ASections: TsNumFormatSections): Integer; overload;
    function Find(AFormatstr: String): Integer; overload;
    property Items[AIndex: Integer]: TsNumFormatParams read GetItem write SetItem; default;
    {@@ Workbook from which the number formats are collected in the list. It is
     mainly needed to get access to the FormatSettings for easy localization of
     some formatting strings. }
    property Workbook: TsWorkbook read FWorkbook;
  end;


function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsCurrencyFormat(ANumFormat: TsNumFormatParams): Boolean; overload;

function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsDateTimeFormat(AFormatStr: String): Boolean; overload;
function IsDateTimeFormat(ANumFormat: TsNumFormatParams): Boolean; overload;

function IsDateFormat(ANumFormat: TsNumFormatParams): Boolean;

function IsTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsTimeFormat(AFormatStr: String): Boolean; overload;
function IsTimeFormat(ANumFormat: TsNumFormatParams): Boolean; overload;

function IsTimeIntervalFormat(ANumFormat: TsNumFormatParams): Boolean;


implementation

uses
  fpsUtils, fpsNumFormatParser;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number format code is for currency,
  i.e. requires currency symbol.

  @param  AFormat   Built-in number format identifier to be checked
  @return True if AFormat is nfCurrency or nfCurrencyRed, false otherwise.
-------------------------------------------------------------------------------}
function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [nfCurrency, nfCurrencyRed];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to currency values.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkCurrency elements; false otherwise
-------------------------------------------------------------------------------}
function IsCurrencyFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkCurrency] <> []);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given number format code is for date/time values.

  @param   AFormat  Built-in number format identifier to be checked
  @return  True if AFormat is a date/time format (such as nfShortTime),
           false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [nfShortDateTime, nfShortDate, nfLongDate,
    nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM,
    nfDayMonth, nfMonthYear, nfTimeInterval];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given string with formatting codes is for date/time values.

  @param   AFormatStr   String with formatting codes to be checked.
  @return  True if AFormatStr is a date/time format string (such as 'hh:nn'),
           false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(AFormatStr: string): Boolean;
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(nil, AFormatStr);
  try
    Result := parser.IsDateTimeFormat;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to date/time values.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkDate or nfkTime elements; false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkDate, nfkTime] <> []);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to a date value.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkDate, but no nfkTime elements; false otherwise
-------------------------------------------------------------------------------}
function IsDateFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and (nfkDate in ANumFormat.Sections[0].Kind);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given built-in number format code is for time values.

  @param   AFormat  Built-in number format identifier to be checked
  @return  True if AFormat represents to a time-format, false otherwise
-------------------------------------------------------------------------------}
function IsTimeFormat(AFormat: TsNumberFormat): boolean;
begin
  Result := AFormat in [nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM,
    nfTimeInterval];
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the given string with formatting codes is for time values.

  @param   AFormatStr   String with formatting codes to be checked
  @return  True if AFormatStr represents a time-format, false otherwise
-------------------------------------------------------------------------------}
function IsTimeFormat(AFormatStr: String): Boolean;
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(nil, AFormatStr);
  try
    Result := parser.IsTimeFormat;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters apply to time values.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkTime, but no nfkDate elements; false otherwise
-------------------------------------------------------------------------------}
function IsTimeFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkTime] <> []);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified number format parameters is a time interval
  format.

  @param   ANumFormat   Number format parameters
  @return  True if Kind of the 1st format parameter section contains the
           nfkTimeInterval elements; false otherwise
-------------------------------------------------------------------------------}
function IsTimeIntervalFormat(ANumFormat: TsNumFormatParams): Boolean;
begin
  Result := (ANumFormat <> nil) and
            (ANumFormat.Sections[0].Kind * [nfkTimeInterval] <> []);
end;


{ TsNumFormatList }

constructor TsNumFormatList.Create(AWorkbook: TsWorkbook; AOwnsData: Boolean);
begin
  inherited Create;
  FClass := TsNumFormatParams;
  FWorkbook := AWorkbook;
  FOwnsData := AOwnsData;
end;

destructor TsNumFormatList.Destroy;
begin
  Clear;
  inherited;
end;

function TsNumFormatList.AddFormat(ASections: TsNumFormatSections): Integer;
var
  nfp: TsNumFormatParams;
begin
  Result := Find(ASections);
  if Result = -1 then begin
    nfp := FClass.Create;
    nfp.Sections := ASections;
    Result := inherited Add(nfp);
  end;
end;

function TsNumFormatList.AddFormat(AFormatStr: String): Integer;
var
  parser: TsNumFormatParser;
  newSections: TsNumFormatSections;
  i: Integer;
begin
  parser := TsNumFormatParser.Create(FWorkbook, AFormatStr);
  try
    SetLength(newSections, parser.ParsedSectionCount);
    for i:=0 to High(newSections) do
      newSections[i] := parser.ParsedSections[i];
    Result := AddFormat(newSections);
  finally
    parser.Free;
  end;
end;

procedure TsNumFormatList.AddBuiltinFormats;
begin
end;

procedure TsNumFormatList.Clear;
var
  i: Integer;
begin
  for i := Count-1 downto 0 do Delete(i);
  inherited;
end;

procedure TsNumFormatList.Delete(AIndex: Integer);
var
  p: TsNumFormatParams;
begin
  if FOwnsData then
  begin
    p := GetItem(AIndex);
    if p <> nil then p.Free;
  end;
  inherited Delete(AIndex);
end;

function TsNumFormatList.Find(ASections: TsNumFormatSections): Integer;
var
  nfp: TsNumFormatParams;
begin
  for Result := 0 to Count-1 do begin
    nfp := GetItem(Result);
    if nfp.SectionsEqualTo(ASections) then
      exit;
  end;
  Result := -1;
end;

function TsNumFormatList.Find(AFormatStr: String): Integer;
var
  nfp: TsNumFormatParams;
begin
  nfp := CreateNumFormatParams(FWorkbook, AFormatStr);
  if nfp = nil then
    Result := -1
  else
    Result := Find(nfp.Sections);
end;

function TsNumFormatList.GetItem(AIndex: Integer): TsNumFormatParams;
begin
  Result := TsNumFormatParams(inherited Items[AIndex]);
end;

procedure TsNumFormatList.SetItem(AIndex: Integer;
  const AValue: TsNumFormatParams);
begin
  inherited Items[AIndex] := AValue;
end;


end.
