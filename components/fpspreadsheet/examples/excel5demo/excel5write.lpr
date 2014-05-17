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
  i, r: Integer;
  number: Double;
  fmt: string;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet(Str_Worksheet1);

  MyWorksheet.Options := MyWorksheet.Options + [soHasFrozenPanes];
  MyWorksheet.LeftPaneWidth := 1;
  MyWorksheet.TopPaneHeight := 2;

  { unfrozen panes not working at the moment
  MyWorksheet.LeftPaneWidth := 20*72*2; // 72 pt = inch  --> 2 inches = 5 cm }

  MyWorkbook.AddFont('Calibri', 20, [], scRed);

  // Change row height
  MyWorksheet.WriteRowHeight(0, 20);  // modify height of row 0 to 20 mm

  // Change colum widths
  MyWorksheet.WriteColWidth(0, 40);
  MyWorksheet.WriteColWidth(1, 20);
  MyWorksheet.WriteColWidth(2, 20);
  MyWorksheet.WriteColWidth(3, 15);
  MyWorksheet.WriteColWidth(4, 15);

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

  // F6 empty cell, only all thin borders
  MyWorksheet.WriteBorders(5, 5, [cbNorth, cbEast, cbSouth, cbWest]);
  MyWorksheet.WriteBorderLineStyle(5, 5, cbSouth, lsDotted);
  MyWorksheet.WriteBorderColor(5, 5, cbSouth, scRed);
  MyWorksheet.WriteBorderLineStyle(5, 5, cbNorth, lsThick);

  // H6 empty cell, only all medium borders
  MyWorksheet.WriteBorders(5, 7, [cbNorth, cbEast, cbSouth, cbWest]);
  MyWorksheet.WriteBorderColor(5, 7, cbSouth, scBlack);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbSouth, lsMedium);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbEast, lsMedium);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbWest, lsMedium);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbNorth, lsMedium);

  // J6 empty cell, only all thick borders
  MyWorksheet.WriteBorders(5, 9, [cbNorth, cbEast, cbSouth, cbWest]);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbSouth, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbEast, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbWest, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbNorth, lsThick);


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

  r:= 10;
  // Write current date/time to cells B11:B16
  MyWorksheet.WriteUTF8Text(r, 0, 'nfShortDate');
  MyWorksheet.WriteDateTime(r, 1, now, nfShortDate);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfLongDate');
  MyWorksheet.WriteDateTime(r, 1, now, nfLongDate);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfShortTime');
  MyWorksheet.WriteDateTime(r, 1, now, nfShortTime);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfLongTime');
  MyWorksheet.WriteDateTime(r, 1, now, nfLongTime);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfShortDateTime');
  MyWorksheet.WriteDateTime(r, 1, now, nfShortDateTime);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFmtDateTime, DM');
  MyWorksheet.WriteDateTime(r, 1, now, nfFmtDateTime, 'DM');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFmtDateTime, MY');
  MyWorksheet.WriteDateTime(r, 1, now, nfFmtDateTime, 'MY');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfShortTimeAM');
  MyWorksheet.WriteDateTime(r, 1, now, nfShortTimeAM);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfLongTimeAM');
  MyWorksheet.WriteDateTime(r, 1, now, nfLongTimeAM);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFmtDateTime, MS');
  MyWorksheet.WriteDateTime(r, 1, now, nfFmtDateTime, 'MS');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFmtDateTime, MSZ');
  MyWorksheet.WriteDateTime(r, 1, now, nfFmtDateTime, 'MSZ');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFmtDateTime, mm:ss.zzz');
  MyWorksheet.WriteDateTime(r, 1, now, nfFmtDateTime, 'mm:ss.zzz');

  // Write formatted numbers
  number := 12345.67890123456789;
  inc(r, 2);
  MyWorksheet.WriteUTF8Text(r, 1, '12345.67890123456789');
  MyWorksheet.WriteUTF8Text(r, 2, '-12345.67890123456789');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfGeneral');
  MyWorksheet.WriteNumber(r, 1, number, nfGeneral);
  MyWorksheet.WriteNumber(r, 2, -number, nfGeneral);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixed, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixed, 0);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 0);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixed, 1 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixed, 1);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 1);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixed, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixed, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixed, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixed, 3);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 3);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 0);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 0);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 1 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 1);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 1);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 3);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 3);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfSci, 1 dec');
  MyWorksheet.WriteNumber(r, 1, number, nfSci, 1);
  MyWorksheet.WriteNumber(r, 2, -number, nfSci, 1);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfSci, 1);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfSci, 1);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfSci, 2 dec');
  MyWorksheet.WriteNumber(r, 1, number, nfSci, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfSci, 2);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfSci, 2);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfSci, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfSci, 3 dec');
  MyWorksheet.WriteNumber(r, 1, number, nfSci, 3);
  MyWorksheet.WriteNumber(r, 2, -number, nfSci, 3);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfSci, 3);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfSci, 3);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 1 dec');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 1);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 1);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 1);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 1);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 2);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 2);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 3);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 3);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 3);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 3);

  inc(r,2);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrency, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfCurrency, 0, 'USD');
  MyWorksheet.WriteNumber(r, 2, -number, nfCurrency, 0, 'USD');
  MyWorksheet.WriteNumber(r, 3, 0.0, nfCurrency, 0, 'USD');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrencyRed, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfCurrencyRed, 0, 'USD');
  MyWorksheet.WriteNumber(r, 2, -number, nfCurrencyRed, 0, 'USD');
  MyWorksheet.WriteNumber(r, 3, 0.0, nfCurrencyRed, 0, 'USD');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrencyDash, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfCurrencyDash, 0, 'USD');
  MyWorksheet.WriteNumber(r, 2, -number, nfCurrencyDash, 0, 'USD');
  MyWorksheet.WriteNumber(r, 3, 0.0, nfCurrencyDash, 0, 'USD');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrencyDashRed, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfCurrencyDashRed, 0, 'USD');
  MyWorksheet.WriteNumber(r, 2, -number, nfCurrencyDashRed, 0, 'USD');
  MyWorksheet.WriteNumber(r, 3, 0.0, nfCurrencyDashRed, 0, 'USD');

  inc(r, 2);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, "$"#,##0_);("$"#,##0)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '"$"#,##0_);("$"#,##0)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '"$"#,##0_);("$"#,##0)');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, "$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '"$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '"$"#,##0.0_);[Red]("$"#,##0.0)');
  inc(r);
  fmt := '"€"#,##0.0_);[Red]("€"#,##0.0)';
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, '+fmt);
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, UTF8ToAnsi(fmt));
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, UTF8ToAnsi(fmt));
  inc(r);
  fmt := '[Green]"¥"#,##0.0_);[Red]-"¥"#,##0.0';
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, '+fmt);
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, UTF8ToAnsi(fmt));
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, UTF8ToAnsi(fmt));
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, _("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)');
  inc(r, 2);
  number := 1.333333333;
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 0);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 1 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 1);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 3);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, hh:mm:ss');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h:m:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'H:M:s');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, hh:mm');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'hh:mm');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h:m');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'h:m');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'h');


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

