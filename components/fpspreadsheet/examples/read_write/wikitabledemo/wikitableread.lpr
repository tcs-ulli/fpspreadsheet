{
wikitableread.lpr

Demonstrates how to read a wikitable (wikimedia format) file using the fpspreadsheet library
Note: the output written by wikitablewrite cannot yet be read by the
wikitableread demo.
}
program wikitableread;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8,
  fpstypes, fpspreadsheet, wikitable, fpsutils;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  InputFilename: string;
  MyDir: string;
  i: Integer;
  CurCell: PCell;

{$R *.res}

begin
  // Open the input file
  MyDir := ExtractFilePath(ParamStr(0));
  InputFileName := MyDir + 'test.wikitable_wikimedia';

  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please make sure a file exists with data in the correct format.');
    Halt;
  end;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  MyWorkbook.ReadFromFile(InputFilename, sfWikiTable_WikiMedia);

  MyWorksheet := MyWorkbook.GetFirstWorksheet;

  // Write all cells with contents to the console
  WriteLn('');
  WriteLn('Contents of the first worksheet of the file:');
  WriteLn('');

  for CurCell in MyWorkSheet.Cells do
  begin
    Write('Row: ', CurCell^.Row,
      ' Col: ', CurCell^.Col, ' Value: ',
      UTF8ToConsole(MyWorkSheet.ReadAsUTF8Text(CurCell^.Row, CurCell^.Col))
    );
    if HasFormula(CurCell) then
      WriteLn(' Formula: ', CurCell^.FormulaValue)
    else
      WriteLn;
  end;

  // Finalization
  MyWorkbook.Free;
end.

