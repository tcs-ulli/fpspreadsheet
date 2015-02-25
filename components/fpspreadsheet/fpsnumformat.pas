unit fpsNumFormat;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

interface

uses
  Classes, SysUtils,
  fpstypes, fpspreadsheet;

type
  {@@ Contents of a number format record }
  TsNumFormatData = class
  public
    {@@ Excel refers to a number format by means of the format "index". }
    Index: Integer;
    {@@ OpenDocument refers to a number format by means of the format "name". }
    Name: String;
    {@@ Identifier of a built-in number format, see TsNumberFormat }
    NumFormat: TsNumberFormat;
    {@@ String of format codes, such as '#,##0.00', or 'hh:nn'. }
    FormatString: string;
  end;

  {@@ Specialized list for number format items }
  TsCustomNumFormatList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TsNumFormatData;
    procedure SetItem(AIndex: Integer; AValue: TsNumFormatData);
  protected
    {@@ Workbook from which the number formats are collected in the list. It is
     mainly needed to get access to the FormatSettings for easy localization of some
     formatting strings. }
    FWorkbook: TsWorkbook;
    {@@ Identifies the first number format item that is written to the file. Items
     having a smaller index are not written. }
    FFirstNumFormatIndexInFile: Integer;
    {@@ Identifies the index of the next Excel number format item to be written.
     Needed for auto-creating of the user-defined Excel number format indexes }
    FNextNumFormatIndex: Integer;
    procedure AddBuiltinFormats; virtual;
    procedure RemoveFormat(AIndex: Integer);

  public
    constructor Create(AWorkbook: TsWorkbook);
    destructor Destroy; override;
    function AddFormat(AFormatIndex: Integer; AFormatName: String;
      ANumFormat: TsNumberFormat; AFormatString: String): Integer; overload;
    function AddFormat(AFormatIndex: Integer; ANumFormat: TsNumberFormat;
      AFormatString: String): Integer; overload;
    function AddFormat(AFormatName: String; ANumFormat: TsNumberFormat;
      AFormatString: String): Integer; overload;
    function AddFormat(ANumFormat: TsNumberFormat; AFormatString: String): Integer; overload;
    procedure AnalyzeAndAdd(AFormatIndex: Integer; AFormatString: String);
    procedure Clear;
    procedure ConvertAfterReading(AFormatIndex: Integer; var AFormatString: String;
      var ANumFormat: TsNumberFormat); virtual;
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat); virtual;
    procedure Delete(AIndex: Integer);
    function Find(ANumFormat: TsNumberFormat; AFormatString: String): Integer; virtual;
    function FindByFormatStr(AFormatString: String): Integer;
    function FindByIndex(AFormatIndex: Integer): Integer;
    function FindByName(AFormatName: String): Integer;
    function FormatStringForWriting(AIndex: Integer): String; virtual;
    procedure Sort;

    {@@ Workbook from which the number formats are collected in the list. It is
     mainly needed to get access to the FormatSettings for easy localization of some
     formatting strings. }
    property Workbook: TsWorkbook read FWorkbook;
    {@@ Identifies the first number format item that is written to the file. Items
     having a smaller index are not written. }
    property FirstNumFormatIndexInFile: Integer read FFirstNumFormatIndexInFile;
    {@@ Number format items contained in the list }
    property Items[AIndex: Integer]: TsNumFormatData read GetItem write SetItem; default;
  end;

function IsCurrencyFormat(AFormat: TsNumberFormat): Boolean;
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsDateTimeFormat(AFormatStr: String): Boolean; overload;
function IsTimeFormat(AFormat: TsNumberFormat): Boolean; overload;
function IsTimeFormat(AFormatStr: String): Boolean; overload;


implementation

uses
  Math,
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
  Checks whether the given number format code is for date/time values.

  @param   AFormat  Built-in number format identifier to be checked
  @return  True if AFormat is a date/time format (such as nfShortTime),
           false otherwise
-------------------------------------------------------------------------------}
function IsDateTimeFormat(AFormat: TsNumberFormat): Boolean;
begin
  Result := AFormat in [{nfFmtDateTime, }nfShortDateTime, nfShortDate, nfLongDate,
    nfShortTime, nfLongTime, nfShortTimeAM, nfLongTimeAM, nfTimeInterval];
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


{*******************************************************************************
*                       TsCustomNumFormatList                                  *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Constructor of the number format list.

  @param AWorkbook The workbook is needed to get access to its "FormatSettings"
                   for localization of some formatting strings.
-------------------------------------------------------------------------------}
constructor TsCustomNumFormatList.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  AddBuiltinFormats;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the number format list: clears the list and destroys the
  format items
-------------------------------------------------------------------------------}
destructor TsCustomNumFormatList.Destroy;
begin
  Clear;
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format described by the Excel format index, the ODF format
  name, the format string, and the built-in format identifier to the list
  and returns the index of the new item.

  @param AFormatIndex  Format index to be used by Excel
  @param AFormatName   Format name to be used by OpenDocument
  @param AFormatString String of formatting codes
  @param ANumFormat    Identifier for built-in number format
  @return              List index of the new item
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.AddFormat(AFormatIndex: Integer;
  AFormatName: String; ANumFormat: TsNumberFormat; AFormatString: String): Integer;
var
  item: TsNumFormatData;
begin
  item := TsNumFormatData.Create;
  item.Index := AFormatIndex;
  item.Name := AFormatName;
  item.NumFormat := ANumFormat;
  item.FormatString := AFormatString;
  Result := inherited Add(item);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format described by the Excel format index, the format string,
  and the built-in format identifier to the list and returns the index of
  the new item in the format list. To be used when writing an Excel file.

  @param AFormatIndex  Format index to be used by Excel
  @param ANumFormat    Identifier for built-in number format
  @param AFormatString String of formatting codes
  @return              Index of the new item in the format list
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.AddFormat(AFormatIndex: Integer;
  ANumFormat: TsNumberFormat; AFormatString: String): integer;
begin
  Result := AddFormat(AFormatIndex, '', ANumFormat, AFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format described by the ODF format name, the format string,
  and the built-in format identifier to the list and returns the index of
  the new item in the format list. To be used when writing an ODS file.

  @param AFormatName   Format name to be used by OpenDocument
  @param AFormatString String of formatting codes
  @param ANumFormat    Identifier for built-in number format
  @return              Index of the new item in the format list
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.AddFormat(AFormatName: String;
  ANumFormat: TsNumberFormat; AFormatString: String): Integer;
begin
  if (AFormatString = '') and (ANumFormat <> nfGeneral) then
  begin
    Result := 0;
    exit;
  end;
  Result := AddFormat(FNextNumFormatIndex, AFormatName, ANumFormat, AFormatString);
  inc(FNextNumFormatIndex);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format described by the format string, and the built-in
  format identifier to the format list and returns the index of the new
  item in the list. The Excel format index and ODS format name are auto-generated.

  @param ANumFormat     Identifier for built-in number format
  @param AFormatString  String of formatting codes
  @return               Index of the new item in the list
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.AddFormat(ANumFormat: TsNumberFormat;
  AFormatString: String): Integer;
begin
  Result := AddFormat('', ANumFormat, AFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Adds the builtin format items to the list. The formats must be specified in
  a way that is compatible with fpc syntax.

  Conversion of the formatstrings to the syntax used in the destination file
  can be done by calling "ConvertAfterReadung" bzw. "ConvertBeforeWriting".
  "AddBuiltInFormats" must be called before user items are added.

  Must specify FFirstNumFormatIndexInFile (BIFF5-8, e.g. don't save formats <164)
  and must initialize the index of the first user format (FNextNumFormatIndex)
  which is automatically incremented when adding user formats.

  In TsCustomNumFormatList nothing is added.
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.AddBuiltinFormats;
begin
  // must be overridden - see xlscommon as an example.
end;

{@@ ----------------------------------------------------------------------------
  Called from the reader when a format item has been read from an Excel file.
  Determines the number format type, format string etc and converts the
  format string to fpc syntax which is used directly for getting the cell text.

  @param AFormatIndex Excel index of the number format read from the file
  @param AFormatString String of formatting codes as read fromt the file.
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.AnalyzeAndAdd(AFormatIndex: Integer;
  AFormatString: String);
var
  nf: TsNumberFormat = nfGeneral;
begin
  if FindByIndex(AFormatIndex) > -1 then
    exit;

  // Analyze & convert the format string, extract infos for internal formatting
  ConvertAfterReading(AFormatIndex, AFormatString, nf);

  // Add the new item
  AddFormat(AFormatIndex, nf, AFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Clears the number format list and frees memory occupied by the format items.
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.Clear;
var
  i: Integer;
begin
  for i:=0 to Count-1 do RemoveFormat(i);
  inherited Clear;
end;

{@@ ----------------------------------------------------------------------------
  Takes the format string as it is read from the file and extracts the
  built-in number format identifier out of it for use by fpc.
  The method also converts the format string to a form that can be used
  by fpc's FormatDateTime and FormatFloat.

  The method should be overridden in a class that knows knows more about the
  details of the spreadsheet file format.

  @param AFormatIndex   Excel index of the number format read
  @param AFormatString  string of formatting codes extracted from the file data
  @param ANumFormat     identifier for built-in fpspreadsheet format extracted
                        from the file data
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.ConvertAfterReading(AFormatIndex: Integer;
  var AFormatString: String; var ANumFormat: TsNumberFormat);
var
  parser: TsNumFormatParser;
  fmt: String;
  lFormatData: TsNumFormatData;
  i: Integer;
begin
  i := FindByIndex(AFormatIndex);
  if i > 0 then
  begin
    lFormatData := Items[i];
    fmt := lFormatData.FormatString;
  end else
    fmt := AFormatString;

  // Analyzes the format string and tries to convert it to fpSpreadsheet format.
  parser := TsNumFormatParser.Create(Workbook, fmt);
  try
    if parser.Status = psOK then
    begin
      ANumFormat := parser.NumFormat;
      AFormatString := parser.FormatString[nfdDefault];
    end else
    begin
      //  Show an error here?
    end;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Is called before collecting all number formats of the spreadsheet and before
  writing them to file. Its purpose is to convert the format string as used by fpc
  to a format compatible with the spreadsheet file format.
  Nothing is changed in the TsCustomNumFormatList, the method needs to be
  overridden by a descendant class which known more about the details of the
  destination file format.

  Needs to be overridden by a class knowing more about the destination file
  format.

  @param AFormatString String of formatting codes. On input in fpc syntax. Is
                       overwritten on output by format string compatible with
                       the destination file.
  @param ANumFormat    Identifier for built-in fpspreadsheet number format
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.ConvertBeforeWriting(var AFormatString: String;
  var ANumFormat: TsNumberFormat);
begin
  Unused(AFormatString, ANumFormat);
  // nothing to do here. But see, e.g., xlscommon.TsBIFFNumFormatList
end;


{@@ ----------------------------------------------------------------------------
  Deletes a format item from the list, and makes sure that its memory is
  released.

  @param  AIndex   List index of the item to be deleted.
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.Delete(AIndex: Integer);
begin
  RemoveFormat(AIndex);
  Delete(AIndex);
end;

{@@ ----------------------------------------------------------------------------
  Seeks a format item with the given properties and returns its list index,
  or -1 if not found.

  @param ANumFormat    Built-in format identifier
  @param AFormatString String of formatting codes
  @return              Index of the format item in the format list,
                       or -1 if not found.
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.Find(ANumFormat: TsNumberFormat;
  AFormatString: String): Integer;
var
  item: TsNumFormatData;
begin
  for Result := Count-1 downto 0 do
  begin
    item := Items[Result];
    if (item <> nil) and (item.NumFormat = ANumFormat) and (item.FormatString = AFormatString)
      then exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Finds the item with the given format string and returns its index in the
  format list, or -1 if not found.

  @param  AFormatString  string of formatting codes to be searched in the list.
  @return Index of the format item in the format list, or -1 if not found.
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.FindByFormatStr(AFormatString: String): integer;
var
  item: TsNumFormatData;
begin
  { We search backwards to find user-defined items first. They usually are
    more appropriate than built-in items. }
  for Result := Count-1 downto 0 do
  begin
    item := Items[Result];
    if item.FormatString = AFormatString then
      exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Finds the item with the given Excel format index and returns its index in
  the format list, or -1 if not found.
  Is used by BIFF file formats.

  @param  AFormatIndex  Excel format index to the searched
  @return Index of the format item in the format list, or -1 if not found.
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.FindByIndex(AFormatIndex: Integer): integer;
var
  item: TsNumFormatData;
begin
  for Result := 0 to Count-1 do
  begin
    item := Items[Result];
    if item.Index = AFormatIndex then
      exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Finds the item with the given ODS format name and returns its index in
  the format list (or -1, if not found)
  To be used by OpenDocument file format.

  @param  AFormatName  Format name as used by OpenDocument to identify a
                       number format

  @return Index of the format item in the list, or -1 if not found
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.FindByName(AFormatName: String): integer;
var
  item: TsNumFormatData;
begin
  for Result := 0 to Count-1 do
  begin
    item := Items[Result];
    if item.Name = AFormatName then
      exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Determines the format string to be written into the spreadsheet file. Calls
  ConvertBeforeWriting in order to convert the fpc format strings to the dialect
  used in the file.

  @param AIndex  Index of the format item under consideration.
  @return        String of formatting codes that will be written to the file.
-------------------------------------------------------------------------------}
function TsCustomNumFormatList.FormatStringForWriting(AIndex: Integer): String;
var
  item: TsNumFormatdata;
  nf: TsNumberFormat;
begin
  item := Items[AIndex];
  if item <> nil then
  begin
    Result := item.FormatString;
    nf := item.NumFormat;
    ConvertBeforeWriting(Result, nf);
  end else
    Result := '';
end;

function TsCustomNumFormatList.GetItem(AIndex: Integer): TsNumFormatData;
begin
  Result := TsNumFormatData(inherited Items[AIndex]);
end;

{@@ ----------------------------------------------------------------------------
  Deletes the memory occupied by the formatting data, but keeps an empty item in
  the list to retain the indexes of following items.

  @param AIndex The number format item at this index will be removed.
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.RemoveFormat(AIndex: Integer);
var
  item: TsNumFormatData;
begin
  item := GetItem(AIndex);
  if item <> nil then
  begin
    item.Free;
    SetItem(AIndex, nil);
  end;
end;

procedure TsCustomNumFormatList.SetItem(AIndex: Integer; AValue: TsNumFormatData);
begin
  inherited Items[AIndex] := AValue;
end;

function CompareNumFormatData(Item1, Item2: Pointer): Integer;
begin
  Result := CompareValue(TsNumFormatData(Item1).Index, TsNumFormatData(Item2).Index);
end;

{@@ ----------------------------------------------------------------------------
  Sorts the format data items in ascending order of the Excel format indexes.
-------------------------------------------------------------------------------}
procedure TsCustomNumFormatList.Sort;
begin
  inherited Sort(@CompareNumFormatData);
end;


end.
