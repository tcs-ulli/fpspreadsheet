{
excel5read.dpr

Demonstrates how to read an Excel 5.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel5read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff5, laz_fpspreadsheet;

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
  InputFileName := MyDir + 'test' + STR_EXCEL_EXTENSION;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.ReadFromFile(InputFilename, sfExcel5);

  MyWorksheet := MyWorkbook.GetFirstWorksheet;

  // Write all cells with contents to the console
  WriteLn('');
  WriteLn('Contents of the first worksheet of the file:');
  WriteLn('');

  for i := 0 to MyWorksheet.GetCellCount - 1 do
  begin
    CurCell := MyWorkSheet.GetCellByIndex(i);
    WriteLn('Row: ', CurCell^.Row, ' Col: ', CurCell^.Col, ' Value: ',
     MyWorkSheet.ReadAsAnsiText(CurCell^.Row, CurCell^.Col));
  end;

  // Finalization
  MyWorkbook.Free;
end.

