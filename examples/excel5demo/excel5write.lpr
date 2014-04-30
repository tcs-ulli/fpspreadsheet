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
  number: Double;
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

  // Write current date/time to cells B11:B16
  MyWorksheet.WriteUTF8Text(10, 0, 'nfShortDate');
  MyWorksheet.WriteDateTime(10, 1, now, nfShortDate);
  MyWorksheet.WriteUTF8Text(11, 0, 'nfShortTime');
  MyWorksheet.WriteDateTime(11, 1, now, nfShortTime);
  MyWorksheet.WriteUTF8Text(12, 0, 'nfLongTime');
  MyWorksheet.WriteDateTime(12, 1, now, nfLongTime);
  MyWorksheet.WriteUTF8Text(13, 0, 'nfShortDateTime');
  MyWorksheet.WriteDateTime(13, 1, now, nfShortDateTime);
  MyWorksheet.WriteUTF8Text(14, 0, 'nfFmtDateTime, DM');
  MyWorksheet.WriteDateTime(14, 1, now, nfFmtDateTime, 'DM');
  MyWorksheet.WriteUTF8Text(15, 0, 'nfFmtDateTime, MY');
  MyWorksheet.WriteDateTime(15, 1, now, nfFmtDateTime, 'MY');
  MyWorksheet.WriteUTF8Text(16, 0, 'nfShortTimeAM');
  MyWorksheet.WriteDateTime(16, 1, now, nfShortTimeAM);
  MyWorksheet.WriteUTF8Text(17, 0, 'nfLongTimeAM');
  MyWorksheet.WriteDateTime(17, 1, now, nfLongTimeAM);
  MyWorksheet.WriteUTF8Text(18, 0, 'nfFmtDateTime, MS');
  MyWorksheet.WriteDateTime(18, 1, now, nfFmtDateTime, 'MS');
  MyWorksheet.WriteUTF8Text(19, 0, 'nfFmtDateTime, MSZ');
  MyWorksheet.WriteDateTime(19, 1, now, nfFmtDateTime, 'MSZ');

  // Write formatted numbers
  number := 12345.67890123456789;
  MyWorksheet.WriteUTF8Text(24, 1, '12345.67890123456789');
  MyWorksheet.WriteUTF8Text(24, 2, '-12345.67890123456789');
  MyWorksheet.WriteUTF8Text(25, 0, 'nfFixed, 0 decs');
  MyWorksheet.WriteNumber(25, 1, number, nfFixed, 0);
  MyWorksheet.WriteNumber(25, 2, -number, nfFixed, 0);
  MyWorksheet.WriteUTF8Text(26, 0, 'nfFixed, 2 decs');
  MyWorksheet.WriteNumber(26, 1, number, nfFixed, 2);
  MyWorksheet.WriteNumber(26, 2, -number, nfFixed, 2);
  MyWorksheet.WriteUTF8Text(27, 0, 'nfFixedTh, 0 decs');
  MyWorksheet.WriteNumber(27, 1, number, nfFixedTh, 0);
  MyWorksheet.WriteNumber(27, 2, -number, nfFixedTh, 0);
  MyWorksheet.WriteUTF8Text(28, 0, 'nfFixedTh, 2 decs');
  MyWorksheet.WriteNumber(28, 1, number, nfFixedTh, 2);
  MyWorksheet.WriteNumber(28, 2, -number, nfFixedTh, 2);
  MyWorksheet.WriteUTF8Text(29, 0, 'nfSci, 1 dec');
  MyWorksheet.WriteNumber(29, 1, number, nfSci);
  MyWorksheet.WriteNumber(29, 2, -number, nfSci);
  MyWorksheet.WriteNumber(29, 3, 1.0/number, nfSci);
  MyWorksheet.WriteNumber(29, 4, -1.0/number, nfSci);
  MyWorksheet.WriteUTF8Text(30, 0, 'nfExp, 2 decs');
  MyWorksheet.WriteNumber(30, 1, number, nfExp, 2);
  MyWorksheet.WriteNumber(30, 2, -number, nfExp, 2);
  MyWorksheet.WriteNumber(30, 3, 1.0/number, nfExp, 2);
  MyWorksheet.WriteNumber(30, 4, -1.0/number, nfExp, 2);

  number := 1.333333333;
  MyWorksheet.WriteUTF8Text(35, 0, 'nfPercentage, 0 decs');
  MyWorksheet.WriteNumber(35, 1, number, nfPercentage, 0);
  MyWorksheet.WriteUTF8Text(36, 0, 'nfPercentage, 2 decs');
  MyWorksheet.WriteNumber(36, 1, number, nfPercentage, 2);
  MyWorksheet.WriteUTF8Text(37, 0, 'nfTimeInterval');
  MyWorksheet.WriteDateTime(37, 1, number, nfTimeInterval);


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

