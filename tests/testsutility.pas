unit testsutility;

{ Utility unit with general functions for the real fpspreadsheet test units,
  e.g. getting temporary files }

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpspreadsheet;

const
  TestFileBIFF8_1904='testbiff8_1904.xls'; //with 1904 datemode date system
  TestFileBIFF8_1899='testbiff8_1899.xls'; //with 1899/1900 datemode date system
  TestFileBIFF8=TestFileBIFF8_1899;
  TestFileODF_1904='testodf_1904.ods'; //OpenDocument/LibreOffice with 1904 datemode date system
  TestFileODF_1899='testodf_1899.ods'; //OpenDocument/LibreOffice with 1899/1900 datemode date system
  TestFileODF=TestFileODF_1899;
  TestFileOOXML_1904='testooxml_1904.xlsx'; //Excel xlsx with 1904 datemode date system
  TestFileOOXML_1899='testooxml_1899.xlsx'; //Excel xlsx with 1899/1900 datemode date system
  TestFileOOXML=TestFileOOXML_1899;
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

// Returns a notation for row bassed on sheet and row
function RowNotation(Worksheet: TsWorksheet; Row: Integer): String;

// Note: using this function instead of GetWorkSheetByName for compatibility with
// older fpspreadsheet versions that don't have that function
function GetWorksheetByName(AWorkBook: TsWorkBook; AName: String): TsWorksheet;

// Gets new empty temp file and returns the file name
// Removes any existing file by that name
// Should be called just before writing to the file as
// GetTempFileName is used which does not guarantee
// file uniqueness
function NewTempFile: String;

implementation

function NewTempFile: String;
begin
  Result := GetTempFileName;
  if FileExists(Result) then
  begin
    DeleteFile(Result);
    sleep(50); //e.g. on Windows, give file system chance to perform changes
  end;
end;

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
  if Worksheet.Name <> '' then
    result := WorkSheet.Name + '!' + ColumnToLetter(Column) + inttostr(Row+1)
  else
    Result := ColumnToLetter(Column) + IntToStr(Row + 1);
end;

function ColNotation(WorkSheet: TsWorksheet; Column:Integer): String;
begin
  if not Assigned(Worksheet) then
    Result := 'ColNotation: error getting worksheet.'
  else
  if Worksheet.Name <> '' then
    Result := WorkSheet.Name + '!' + ColumnToLetter(Column)
  else
    Result := ColumnToLetter(Column);
end;

function RowNotation(Worksheet: TsWorksheet; Row: Integer): String;
begin
  if not Assigned(Worksheet) then
    Result := 'RowNotation: error getting worksheet.'
  else
  if Worksheet.Name <> '' then
    Result := Worksheet.Name + '!' + IntToStr(Row+1)
  else
    Result := IntToStr(Row+1);
end;

end.

