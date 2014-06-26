{
excel8read.lpr

Demonstrates how to read an Excel 8.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel8read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff8,
  laz_fpspreadsheet, fpsutils;

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
  InputFileName := MyDir + 'test.xls';

  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run excel8write first.');
    Halt;
  end;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFormulas := true;

  MyWorkbook.ReadFromFile(InputFilename, sfExcel8);

  MyWorksheet := MyWorkbook.GetFirstWorksheet;

  // Write all cells with contents to the console
  WriteLn('');
  WriteLn('Contents of the first worksheet of the file:');
  WriteLn('');

  CurCell := MyWorkSheet.GetFirstCell();
  for i := 0 to MyWorksheet.GetCellCount - 1 do
  begin
    Write('Row: ', CurCell^.Row,
     ' Col: ', CurCell^.Col, ' Value: ',
     UTF8ToAnsi(MyWorkSheet.ReadAsUTF8Text(CurCell^.Row,
       CurCell^.Col))
     );
    if Length(CurCell^.RPNFormulaValue) > 0 then
      WriteLn(' Formula: ', MyWorkSheet.ReadRPNFormulaAsString(CurCell))
    else
      WriteLn;
    CurCell := MyWorkSheet.GetNextCell();
  end;

  // Finalization
  MyWorkbook.Free;
end.

