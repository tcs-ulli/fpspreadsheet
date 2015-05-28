{
test_write_colors.lpr

Demonstrates how to write the Excel-97 colors to a worksheet

AUTHOR: Wernber Pamler
}

program demo_write_colors;

{$mode delphi}{$H+}

uses
  Classes, SysUtils,
  fpsTypes, fpsutils, fpspalette, fpspreadsheet, xlsbiff8;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  palette: TsPalette;
  row: Cardinal;

const
  TestFile = 'test_colors.xls';

begin
  Writeln('Starting program.');
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet('Colors');

    // Create the palette
    palette := TsPalette.Create;
    try
      palette.AddExcelColors;

      // Write colors to worksheet
      for row := 0 to palette.Count-1 do begin
        Myworksheet.WriteBackgroundColor(row, 0, palette[row]);
        Myworksheet.WriteUTF8Text(row, 0, GetColorName(palette[row]));
        MyWorksheet.WriteFontColor(row, 0, HighContrastColor(palette[row]));
        MyWorksheet.WriteHorAlignment(row, 0, haCenter);
      end;
    finally
      palette.Free;
    end;

    MyWorksheet.WriteColWidth(0, 25);

    // Save the spreadsheet to a file
    MyWorkbook.WriteToFile(MyDir + TestFile, sfExcel8, True);

  finally
    MyWorkbook.Free;
  end;

  writeln('Finished.');
  WriteLn('Please open "'+Testfile+'" in your spreadsheet program.');
  ReadLn;
end.

