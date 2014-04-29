{
excel5write.dpr

Demonstrates how to write an Excel 5.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel5write;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff5;

const
  Str_First = 'First';
  Str_Second = 'Second';
  Str_Third = 'Third';
  Str_Fourth = 'Fourth';
  Str_Worksheet1 = 'Meu Relatório';
  Str_Worksheet2 = 'My Worksheet 2';
  Str_Total = 'Total:';
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyRPNFormula: TsRPNFormula;
  MyDir: string;
  i: Integer;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet(Str_Worksheet1);

  MyWorkbook.AddFont('Calibri', 20, [], scRed);

  // Write some cells
  MyWorksheet.WriteNumber(0, 0, 1.0);// A1
  MyWorksheet.WriteVertAlignment(0, 0, vaCenter);


  MyWorksheet.WriteNumber(0, 1, 2.0);// B1
  MyWorksheet.WriteNumber(0, 2, 3.0);// C1
  MyWorksheet.WriteNumber(0, 3, 4.0);// D1

  MyWorksheet.WriteUTF8Text(4, 2, Str_Total);// C5
  MyWorksheet.WriteBorders(4, 2, [cbEast, cbNorth, cbWest, cbSouth]);
  myWorksheet.WriteFontColor(4, 2, scRed);
  MyWorksheet.WriteBackgroundColor(4, 2, scSilver);
  MyWorksheet.WriteVertAlignment(4, 2, vaTop);

  MyWorksheet.WriteNumber(4, 3, 10);         // D5

  MyWorksheet.WriteUTF8Text(4, 4, 'This is a long wrapped text.');
  MyWorksheet.WriteUsedFormatting(4, 4, [uffWordWrap]);
  MyWorksheet.WriteHorAlignment(4, 4, haCenter);

  MyWorksheet.WriteUTF8Text(4, 5, 'Stacked text');
  MyWorksheet.WriteTextRotation(4, 5, rtStacked);
  MyWorksheet.WriteHorAlignment(4, 5, haCenter);

  MyWorksheet.WriteUTF8Text(4, 6, 'CW-rotated text');
  MyWorksheet.WriteTextRotation(4, 6, rt90DegreeClockwiseRotation);

  MyWorksheet.WriteUTF8Text(4, 7, 'CCW-rotated text');
  MyWorksheet.WriteTextRotation(4, 7, rt90DegreeCounterClockwiseRotation);

  MyWorksheet.WriteUTF8Text(4, 8, 'CW-rotated text');
  MyWorksheet.WriteTextRotation(4, 8, rt90DegreeClockwiseRotation);
  MyWorksheet.WriteVertAlignment(4, 8, vaTop);
  MyWorksheet.WriteHorAlignment(4, 8, haLeft);

  MyWorksheet.WriteUTF8Text(4, 9, 'CCW-rotated text');
  MyWorksheet.WriteTextRotation(4, 9, rt90DegreeCounterClockwiseRotation);
  MyWorksheet.WriteVertAlignment(4, 9, vaTop);
  Myworksheet.WriteHorAlignment(4, 9, haRight);

  MyWorksheet.WriteUTF8Text(4, 10, 'CW-rotated text');
  MyWorksheet.WriteTextRotation(4, 10, rt90DegreeClockwiseRotation);
  MyWorksheet.WriteVertAlignment(4, 10, vaCenter);

  MyWorksheet.WriteUTF8Text(4, 11, 'CCW-rotated text');
  MyWorksheet.WriteTextRotation(4, 11, rt90DegreeCounterClockwiseRotation);
  MyWorksheet.WriteVertAlignment(4, 11, vaCenter);

  // Write current date/time
  MyWorksheet.WriteDateTime(5, 0, now);
  MyWorksheet.WriteFont(5, 0, 'Courier New', 20, [fssBold, fssItalic, fssUnderline], scBlue);

{ Uncomment this to test large XLS files
  for i := 2 to 20 do
  begin
    MyWorksheet.WriteAnsiText(i, 0, ParamStr(0));
    MyWorksheet.WriteAnsiText(i, 1, ParamStr(0));
    MyWorksheet.WriteAnsiText(i, 2, ParamStr(0));
    MyWorksheet.WriteAnsiText(i, 3, ParamStr(0));
  end;
}

  // Write the formula E1 = A1 + B1
  SetLength(MyRPNFormula, 3);
  MyRPNFormula[0].ElementKind := fekCell;
  MyRPNFormula[0].Col := 0;
  MyRPNFormula[0].Row := 0;
  MyRPNFormula[1].ElementKind := fekCell;
  MyRPNFormula[1].Col := 1;
  MyRPNFormula[1].Row := 0;
  MyRPNFormula[2].ElementKind := fekAdd;
  MyWorksheet.WriteRPNFormula(0, 4, MyRPNFormula);

  // Write the formula F1 = ABS(A1)
  SetLength(MyRPNFormula, 2);
  MyRPNFormula[0].ElementKind := fekCell;
  MyRPNFormula[0].Col := 0;
  MyRPNFormula[0].Row := 0;
  MyRPNFormula[1].ElementKind := fekABS;
  MyWorksheet.WriteRPNFormula(0, 5, MyRPNFormula);

  //MyFormula.FormulaStr := '';

  // Creates a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet(Str_Worksheet2);

  // Write some string cells
  MyWorksheet.WriteUTF8Text(0, 0, Str_First);
  MyWorksheet.WriteUTF8Text(0, 1, Str_Second);
  MyWorksheet.WriteUTF8Text(0, 2, Str_Third);
  MyWorksheet.WriteUTF8Text(0, 3, Str_Fourth);

  // Creates a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet('Colors');
  for i:=0 to MyWorkbook.GetPaletteSize-1 do begin
    MyWorksheet.WriteBlank(i, 0);
    Myworksheet.WriteBackgroundColor(i, 0, TsColor(i));
    MyWorksheet.WriteUTF8Text(i, 1, MyWorkbook.GetColorName(i));
  end;

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.xls', sfExcel5, true);
  MyWorkbook.Free;
end.

