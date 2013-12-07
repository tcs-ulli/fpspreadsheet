unit testsutility;

{ Utility unit with general functions for tests,
  e.g. getting temporary files }

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpspreadsheet;

const
  TestFileBIFF8='testbiff8.xls'; //with 1904 datemode date system
  TestFileBIFF8_1899='testbiff8_1899.xls'; //with 1899/1900 datemode date system
  TestFileManual='testmanual.xls';
  DatesSheet = 'Dates'; //worksheet name
  FormulasSheet = 'Formulas'; //worksheet name
  ManualSheet = 'ManualTests'; //worksheet names
  NumbersSheet = 'Numbers'; //worksheet name
  StringsSheet = 'Texts'; //worksheet name

// Returns an A.. notation based on row (e.g. A1).
// Useful as all test values should be put in the A column of the spreadsheet
function CellNotation(Row: integer): string;

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

function CellNotation(Row: integer): string;
begin
  // From 0-based to Excel A1 notation
  // Note: we're only testing in the A column, that's why we hardcode the value
  result:=DatesSheet+'!A'+inttostr(Row+1);
end;


end.

