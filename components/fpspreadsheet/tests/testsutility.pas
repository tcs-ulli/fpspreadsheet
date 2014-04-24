unit testsutility;

{ Utility unit with general functions for the real fpspreadsheet test units,
  e.g. getting temporary files }

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpspreadsheet;

const
  TestFileBIFF8='testbiff8.xls'; //with 1904 datemode date system
  TestFileBIFF8_1899='testbiff8_1899.xls'; //with 1899/1900 datemode date system
  TestFileODF='testodf.ods'; //OpenDocument/LibreOffice with 1904 datemode date system
  TestFileODF_1899='testodf_1899.ods'; //OpenDocument/LibreOffice with 1899/1900 datemode date system
  TestFileManual='testmanual.xls'; //file name for manual checking using external spreadsheet program (Excel/LibreOffice..)
  DatesSheet = 'Dates'; //worksheet name
  FormulasSheet = 'Formulas'; //worksheet name
  ManualSheet = 'ManualTests'; //worksheet names
  NumbersSheet = 'Numbers'; //worksheet name
  StringsSheet = 'Texts'; //worksheet name

// Returns an A.. notation based on sheet, row, optional column (e.g. A1).
function CellNotation(WorkSheet: TsWorksheet; Row: integer; Column: integer=0): string;

// Returns an A notation of column based on sheet and column
function ColNotation(WorkSheet: TsWorksheet; Column:Integer): String;

// Note: using this function instead of GetWorkSheetByName for compatibility with
// older fpspreadsheet versions that don't have that function
function GetWorksheetByName(AWorkBook: TsWorkBook; AName: String): TsWorksheet;

implementation

function GetWorksheetByName(AWorkBook: TsWorkBook; AName: String): TsWorksheet;
var
  i:integer;
  Worksheets: cardinal;
begin
  Result := nil;
  if AWorkBook=nil then
    exit;

  Worksheets:=AWorkBook.GetWorksheetCount;

  try
    for i:=0 to Worksheets-1 do
    begin
      if AWorkBook.GetWorksheetByIndex(i).Name=AName then
      begin
        Result := AWorkBook.GetWorksheetByIndex(i);
        exit;
      end;
    end;
  except
    Result := nil; //e.g. Getworksheetbyindex unexpectedly gave nil
    exit;
  end;
end;

// Converts column number to A.. notation
function ColumnToLetter(Column: integer): string;
begin
  begin
    if Column < 26 then
      Result := char(Column+65)
    else
    if Column < 26*26 then
      Result := char(Column div 26 + 65) +
        char(Column mod 26 + 65)
    else
    if Column < 26*26*26 then
      Result := char(Column div (26*26) + 65) +
        char(Column mod (26*26) div 26 + 65) +
        char(Column mod (26*26*26) + 65)
    else
      Result := 'ColNotation: At most three digits supported.';
  end;
end;

function CellNotation(WorkSheet: TsWorksheet; Row: integer; Column: integer=0): string;
begin
  // From 0-based to Excel A1 notation
  if not(assigned(Worksheet)) then
    result:='CellNotation: error getting worksheet.'
  else
    result:=WorkSheet.Name+'!'+ColumnToLetter(Column)+inttostr(Row+1)
end;

function ColNotation(WorkSheet: TsWorksheet; Column:Integer): String;
begin
  if not Assigned(Worksheet) then
    Result := 'ColNotation: error getting worksheet.'
  else
    Result := WorkSheet.Name + '!' + ColumnToLetter(Column);
end;

end.

