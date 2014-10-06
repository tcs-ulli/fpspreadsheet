{
excel8write.lpr

Demonstrates how to write an Excel 8+ file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel8write;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, fpsRPN, xlsbiff8;

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
  MyDir: string;
  number: Double;
  lCell: PCell;
  lCol: TCol;
  i: Integer;
  r: Integer = 10;
  s: String;
begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.SetDefaultFont('Calibri', 9);
  MyWorkbook.UsePalette(@PALETTE_BIFF8, Length(PALETTE_BIFF8));
  MyWorkbook.FormatSettings.CurrencyFormat := 2;
  MyWorkbook.FormatSettings.NegCurrFormat := 14;
  MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];

  MyWorksheet := MyWorkbook.AddWorksheet(Str_Worksheet1);
  MyWorksheet.Options := MyWorksheet.Options - [soShowGridLines];

  number := 10000;
  {
  MyWorksheet.Options := MyWorksheet.Options + [soHasFrozenPanes];
  myWorksheet.LeftPaneWidth := 1;
  MyWorksheet.TopPaneHeight := 2;
  }

  { non-frozen panes not working, at the moment. Requires SELECTION records?
  MyWorksheet.LeftPaneWidth := 20*72*2; // 72 pt = inch  --> 2 inches = 5 cm
  }
  // Write some cells
  MyWorksheet.WriteNumber(0, 0, 1.0);        // A1
  MyWorksheet.WriteNumber(0, 1, 2.0);        // B1
  MyWorksheet.WriteNumber(0, 2, 3.0);        // C1
  MyWorksheet.WriteNumber(0, 3, 4.0);        // D1
  MyWorksheet.WriteUTF8Text(4, 2, Str_Total);// C5
  MyWorksheet.WriteNumber(4, 3, 10);         // D5

  // D6 number with background color
  MyWorksheet.WriteNumber(5, 3, 10);
  lCell := MyWorksheet.GetCell(5, 3);
  lCell^.BackgroundColor := scPurple;
  lCell^.UsedFormattingFields := [uffBackgroundColor];
  // or: MyWorksheet.WriteBackgroundColor(5, 3, scPurple);
  MyWorksheet.WriteFontColor(5, 3, scWhite);
  MyWorksheet.WriteFontSize(5, 3, 12);
  // or: MyWorksheet.WriteFont(5, 3, 'Arial', 12, [], scWhite);

  // E6 empty cell, only background color
  MyWorksheet.WriteBackgroundColor(5, 4, scYellow);

  // F6 empty cell, all borders
  MyWorksheet.WriteBorders(5, 5, [cbNorth, cbEast, cbSouth, cbWest]);
  MyWorksheet.WriteBorderStyle(5, 5, cbSouth, lsDotted, scRed);
  MyWorksheet.WriteBorderLineStyle(5, 5, cbNorth, lsThick);

  // F7, top border only, but different color
  MyWorksheet.WriteBorders(6, 5, [cbNorth]);
  MyWorksheet.WriteBorderColor(6, 5, cbNorth, scGreen);
  MyWorksheet.WriteUTF8Text(6, 5, 'top border green or red?');
  // Excel shows it to be red --> the upper border wins

  // H6 empty cell, all medium borders
  MyWorksheet.WriteBorders(5, 7, [cbNorth, cbEast, cbSouth, cbWest]);
  MyWorksheet.WriteBorderColor(5, 7, cbSouth, scBlack);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbSouth, lsMedium);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbEast, lsMedium);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbWest, lsMedium);
  MyWorksheet.WriteBorderLineStyle(5, 7, cbNorth, lsMedium);

  // J6 empty cell, all thick borders, diagonals thin&red
  MyWorksheet.WriteBorders(5, 9, [cbNorth, cbEast, cbSouth, cbWest, cbDiagUp, cbDiagDown]);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbSouth, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbEast, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbWest, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbNorth, lsThick);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbDiagUp, lsThin);
  MyWorksheet.WriteBorderLineStyle(5, 9, cbDiagDown, lsThin);
  MyWorksheet.WriteBorderColor(5, 9, cbDiagUp, scRed);

  // K6 empty cell, top border thick
  MyWorksheet.WriteBorders(5, 11, [cbNorth]);
  MyWorksheet.WriteBorderLineStyle(5, 11, cbNorth, lsThick);

  // L6 empty cell, bottom border medium
  MyWorksheet.WriteBorders(5, 12, [cbSouth]);
  MyWorksheet.WriteBorderLineStyle(5, 12, cbSouth, lsMedium);

  // M6 empty cell, top & bottom border dashed and dotted
  MyWorksheet.WriteBorders(5, 13, [cbNorth, cbSouth]);
  MyWorksheet.WriteBorderLineStyle(5, 13, cbNorth, lsDashed);
  MyWorksheet.WriteBorderLineStyle(5, 13, cbSouth, lsDotted);

  // N6 empty cell, left border: double
//  MyWorksheet.WriteBlank(5, 14);
  MyWorksheet.WriteBorders(5, 14, [cbWest]);
  MyWorksheet.WriteBorderLineStyle(5, 14, cbWest, lsDouble);

  // Word-wrapped long text in D7
  MyWorksheet.WriteUTF8Text(6, 3, 'This is a very, very, very, very long wrapped text.');
  MyWorksheet.WriteUsedFormatting(6, 3, [uffWordwrap]);

  // Cell with changed font in D8
  MyWorksheet.WriteUTF8Text(7, 3, 'This is 16pt red bold & italic Times New Roman.');
  Myworksheet.WriteFont(7, 3, 'Times New Roman', 16, [fssBold, fssItalic], scRed);

  // Cell with changed font and background in D9
  MyWorksheet.WriteUTF8Text(8, 3, 'Colors...');
  MyWorksheet.WriteFont(8, 3, 'Courier New', 12, [fssUnderline], scBlue);
  MyWorksheet.WriteBackgroundColor(8, 3, scYellow);
                                         {}
    {
  // Uncomment this to test large XLS files
  for i := 50 to 1000 do
  begin
//    MyWorksheet.WriteUTF8Text(i, 0, ParamStr(0));
//    MyWorksheet.WriteUTF8Text(i, 1, ParamStr(0));
//    MyWorksheet.WriteUTF8Text(i, 2, ParamStr(0));
    MyWorksheet.WriteUTF8Text(i, 3, ParamStr(0));
  end;
     }

  // Write the string formula E1 = A1 + B1 ...
  MyWorksheet.WriteFormula(0, 4, 'A1+B1');
  // ... and the rpn formula E2 = A1 + B1
  MyWorksheet.WriteRPNFormula(1, 4, CreateRPNFormula(
    RPNCellValue('A1',
    RPNCellValue('B1',
    RPNFunc(fekAdd,
    nil)))));

  // Write the formula F1 = ABS(A1) as string formula ...
  MyWorksheet.WriteFormula(0, 5, 'ABS(A1)');
  // ... and F2 = ABS(A1) as rpn formula
  MyWorksheet.WriteRPNFormula(1, 5, CreateRPNFormula(
    RPNCellValue('A1',
    RPNFunc('ABS',
    nil))));

  // Write a string formula to G1 = "A" & "B" ...
  MyWorksheet.WriteFormula(0, 6, '"A"&"B"');
  // ... and again as rpn formula
  MyWorksheet.WriteRPNFormula(1, 6, CreateRPNFormula(
    RPNString('A',
    RPNSTring('B',
    RPNFunc(fekConcat,
    nil)))));

  r := 10;
  MyWorksheet.WriteUTF8Text(r, 0, 'Writing current date/time:');
  inc(r, 2);
  // Write current date/time to cells B11:B16
  MyWorksheet.WriteUTF8Text(r, 0, '(default format)');
  MyWorksheet.WriteDateTime(r, 1, now);
  inc(r);
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
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, ''dd/mmm''');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'dd/mmm');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, ''MMM/YY''');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'mmm/yy');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfShortTimeAM');
  MyWorksheet.WriteDateTime(r, 1, now, nfShortTimeAM);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfLongTimeAM');
  MyWorksheet.WriteDateTime(r, 1, now, nfLongTimeAM);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, ''mm:ss''');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'mm:ss');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, ''mm:ss.z''');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'mm:ss.z');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, ''mm:ss.zzz''');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'mm:ss.zzz');

  // Write formatted numbers
  s := '31415.9265359';
  val(s, number, i);

  inc(r, 2);
  MyWorksheet.WriteUTF8Text(r, 0, 'The number '+s+' is displayed in various formats:');
  inc(r,2);
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
  inc(r,2);

  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 0 dec');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 0);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 0);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 0);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 0);
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
  MyWorksheet.WriteCurrency(r, 1, number, nfCurrency, 0, 'USD');
  MyWorksheet.WriteCurrency(r, 2, -number, nfCurrency, 0, 'USD');
  MyWorksheet.WriteCurrency(r, 3, 0.0, nfCurrency, 0, 'USD');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrencyRed, 0 decs');
  MyWorksheet.WriteCurrency(r, 1, number, nfCurrencyRed, 0, 'USD');
  MyWorksheet.WriteCurrency(r, 2, -number, nfCurrencyRed, 0, 'USD');
  MyWorksheet.WriteCurrency(r, 3, 0.0, nfCurrencyRed, 0, 'USD');

  inc(r,2);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, "EUR "#,##0_);("EUR "#,##0)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '"EUR "#,##0_);("EUR "#,##0)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '"EUR "#,##0_);("EUR "#,##0)');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, "$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '"$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '"$"#,##0.0_);[Red]("$"#,##0.0)');

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
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval Default=[h]:mm:ss');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h:m:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'h:m:s');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [h]:m:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[h]:m:s');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [h]:n:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[h]:n:s');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [hh]:mm');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[hh]:mm');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [hh]:nn');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[hh]:nn');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval [h]:m');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[h]:m');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [h]:n');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[h]:n');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [h]');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[h]');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [m]:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[m]:s');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, m:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'm:s');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [mm]:ss');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[mm]:ss');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, [ss]');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, '[ss]');

  // Set width of columns 0, 1 and 5
  MyWorksheet.WriteColWidth(0, 30);
  lCol.Width := 25;
  MyWorksheet.WriteColInfo(1, lCol);
  MyWorksheet.WriteColWidth(2, 15);
  MyWorksheet.WriteColWidth(3, 15);
  MyWorksheet.WriteColWidth(4, 15);
  lCol.Width := 5;
  MyWorksheet.WriteColInfo(5, lCol);

  // Set height of rows 0
  MyWorksheet.WriteRowHeight(0, 5);  // 5 lines

  // Creates a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet(Str_Worksheet2);

  // Write some string cells
  MyWorksheet.WriteUTF8Text(0, 0, Str_First);
  MyWorksheet.WriteUTF8Text(0, 1, Str_Second);
  MyWorksheet.WriteUTF8Text(0, 2, Str_Third);
  MyWorksheet.WriteUTF8Text(0, 3, Str_Fourth);
  MyWorksheet.WriteTextRotation(0, 0, rt90DegreeClockwiseRotation);
  MyWorksheet.WriteUsedFormatting(0, 1, [uffBold]);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.xls', sfExcel8, true);
  MyWorkbook.Free;
end.

