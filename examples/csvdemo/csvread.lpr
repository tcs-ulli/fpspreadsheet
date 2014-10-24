{
csvread.dpr

Demonstrates how to read a CSV file using the fpspreadsheet library
}

program myexcel2read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, fpscsv;

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
  InputFileName := MyDir + 'test' + STR_COMMA_SEPARATED_EXTENSION;
  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run excel2write first.');
    Halt;
  end;

  WriteLn('Opening input file ', InputFilename);

  // Tab-delimited
  CSVParams.Delimiter := #9;
  CSVParams.QuoteChar := '''';

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
  MyWorkbook.ReadFromFile(InputFilename, sfCSV);

  MyWorksheet := MyWorkbook.GetFirstWorksheet;

  // Write all cells with contents to the console
  WriteLn('');
  WriteLn('Contents of the first worksheet of the file:');
  WriteLn('');

  CurCell := MyWorkSheet.GetFirstCell();
  for i := 0 to MyWorksheet.GetCellCount - 1 do
  begin
    if HasFormula(CurCell) then
      WriteLn('Row: ', CurCell^.Row, ' Col: ', CurCell^.Col, ' Formula: ', MyWorksheet.ReadFormulaAsString(CurCell))
    else
    WriteLn('Row: ', CurCell^.Row,
      ' Col: ', CurCell^.Col,
      ' Value: ', UTF8ToAnsi(MyWorkSheet.ReadAsUTF8Text(CurCell^.Row, CurCell^.Col))
     );
    CurCell := MyWorkSheet.GetNextCell();
  end;

  // Finalization
  MyWorkbook.Free;
end.

