{
excel5read.lpr

Demonstrates how to read an Excel 5.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel5read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpsTypes, fpsUtils, fpspreadsheet, xlsbiff5;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  InputFilename: string;
  MyDir: string;
  i: Integer;
  CurCell: PCell;
begin
  // Open the input file
  MyDir := ExtractFilePath(ParamStr(0));
  InputFileName := MyDir + 'test.xls';
  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run excel5write first.');
    Halt;
  end;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(InputFilename, sfExcel5);

    MyWorksheet := MyWorkbook.GetFirstWorksheet;

    // Write all cells with contents to the console
    WriteLn('');
    WriteLn('Contents of the first worksheet of the file:');
    WriteLn('');

    for CurCell in MyWorksheet.Cells do
    begin
      Write('Row: ', CurCell^.Row,
       ' Col: ', CurCell^.Col, ' Value: ',
      UTF8ToConsole(MyWorkSheet.ReadAsUTF8Text(CurCell^.Row, CurCell^.Col)));
      if HasFormula(CurCell) then
        Write(' - Formula: ', CurCell^.FormulaValue);
      WriteLn;
    end;

  finally
    // Finalization
    MyWorkbook.Free;
  end;
end.

