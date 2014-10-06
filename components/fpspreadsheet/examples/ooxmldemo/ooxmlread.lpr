{
ooxmlread.lpr

Demonstrates how to read an Excel xlsx file using the fpspreadsheet library

}

program ooxmlread;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsxooxml; //fpsallformats;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  InputFilename: String;
  MyDir: string;
  cell: PCell;
  i: Integer;

begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Open the input file
  MyDir := ExtractFilePath(ParamStr(0));
  InputFileName := MyDir + 'test.xlsx';
  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run opendocwrite first.');
    Halt;
  end;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
  MyWorkbook.ReadFromFile(InputFilename, sfOOXML);

  MyWorksheet := MyWorkbook.GetFirstWorksheet;

  // Write all cells with contents to the console
  WriteLn('');
  WriteLn('Contents of the first worksheet of the file:');
  WriteLn('');

  cell := MyWorkSheet.GetFirstCell();
  for i := 0 to MyWorksheet.GetCellCount - 1 do begin
    WriteLn('Row: ', cell^.Row,
      ' Col: ', cell^.Col, ' Value: ',
      UTF8ToAnsi(MyWorkSheet.ReadAsUTF8Text(cell^.Row, cell^.Col))
    );
    cell := MyWorkSheet.GetNextCell();
  end;

  // Finalization
  MyWorkbook.Free;
end.

