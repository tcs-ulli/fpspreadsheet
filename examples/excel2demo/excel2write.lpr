{
excel2write.dpr

Demonstrates how to write an Excel 2.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel2write;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff2, laz_fpspreadsheet;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyRPNFormula: TsRPNFormula;
  MyDir: string;
begin
  // Open the output file
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  // Write some number cells
  MyWorksheet.WriteNumber(0, 0, 1.0);
  MyWorksheet.WriteNumber(0, 1, 2.0);
  MyWorksheet.WriteNumber(0, 2, 3.0);
  MyWorksheet.WriteNumber(0, 3, 4.0);

  // Write the formula E1 = ABS(A1)
  SetLength(MyRPNFormula, 2);
  MyRPNFormula[0].ElementKind := fekCell;
  MyRPNFormula[0].Col := 0;
  MyRPNFormula[0].Row := 0;
  MyRPNFormula[1].ElementKind := fekABS;
  MyWorksheet.WriteRPNFormula(0, 4, MyRPNFormula);

  // Write the formula F1 = ROUND(A1, 0)
  SetLength(MyRPNFormula, 3);
  MyRPNFormula[0].ElementKind := fekCell;
  MyRPNFormula[0].Col := 0;
  MyRPNFormula[0].Row := 0;
  MyRPNFormula[1].ElementKind := fekNum;
  MyRPNFormula[1].DoubleValue := 0.0;
  MyRPNFormula[2].ElementKind := fekROUND;
  MyWorksheet.WriteRPNFormula(0, 5, MyRPNFormula);

  // Write some string cells
  MyWorksheet.WriteUTF8Text(1, 0, 'First');
  MyWorksheet.WriteFont(1, 0, 'Arial', 12, [fssBold, fssItalic, fssUnderline], scRed);
  MyWorksheet.WriteUTF8Text(1, 1, 'Second');
  MyWorksheet.WriteUTF8Text(1, 2, 'Third');
  MyWorksheet.WriteUTF8Text(1, 3, 'Fourth');

  // Write current date/time
  MyWorksheet.WriteDateTime(2, 0, now);

  // Write cell with background color
  MyWorksheet.WriteUTF8Text(3, 0, 'Text');
  MyWorksheet.WriteBackgroundColor(3, 0, scSilver);

  // Empty cell with background color
  MyWorksheet.WriteBackgroundColor(3, 1, scGrey);

  // Cell2 with top and bottom borders
  MyWorksheet.WriteUTF8Text(4, 0, 'Text');
  MyWorksheet.WriteBorders(4, 0, [cbNorth, cbSouth]);
  MyWorksheet.WriteBorders(4, 1, [cbNorth, cbSouth]);
  MyWorksheet.WriteBorders(4, 2, [cbNorth, cbSouth]);

  // Left, center, right aligned texts
  MyWorksheet.WriteUTF8Text(5, 0, 'L');
  MyWorksheet.WriteUTF8Text(5, 1, 'C');
  MyWorksheet.WriteUTF8Text(5, 2, 'R');
  MyWorksheet.WriteHorAlignment(5, 0, haLeft);
  MyWorksheet.WriteHorAlignment(5, 1, haCenter);
  MyWorksheet.WriteHorAlignment(5, 2, haRight);

  // Red font, italic
  MyWorksheet.WriteNumber(6, 0, 2014);
  MyWorksheet.WriteFont(6, 0, 'Calibri', 15, [fssItalic], scRed);
  MyWorksheet.WriteNumber(6, 1, 2015);
  MyWorksheet.WriteFont(6, 1, 'Times New Roman', 9, [fssUnderline], scBlue);
  MyWorksheet.WriteNumber(6, 2, 2016);
  MyWorksheet.WriteFont(6, 2, 'Courier New', 8, [], scBlue);
  MyWorksheet.WriteNumber(6, 3, 2017);
  MyWorksheet.WriteFont(6, 3, 'Arial', 18, [fssBold], scBlue);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test' + STR_EXCEL_EXTENSION, sfExcel2, true);
  MyWorkbook.Free;
end.

