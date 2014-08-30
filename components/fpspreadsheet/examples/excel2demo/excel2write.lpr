{
excel2write.lpr

Demonstrates how to write an Excel 2.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel2write;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpspreadsheet, xlsbiff2;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyRPNFormula: TsRPNFormula;
  MyDir: string;
  number: Double;
  lCol: TCol;
  lRow: TRow;
  r: Integer;
  fmt: String;
begin
  // Open the output file
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  //MyWorksheet.WriteColWidth(0, 5);
  //MyWorksheet.WriteColWidth(1, 30);

  MyWorksheet.WriteRowHeight(0, 3);  // 3 lines

  // Turn off grid lines and hide headers
  //MyWorksheet.Options := MyWorksheet.Options - [soShowGridLines, soShowHeaders];

{    -- currently not working
  //MyWorksheet.Options := MyWorksheet.Options + [soHasFrozenPanes];
  MyWorksheet.LeftPaneWidth := 1;
  MyWorksheet.TopPaneHeight := 3;
}

  // Write some number cells
  MyWorksheet.WriteNumber(0, 0, 1.0);
  MyWorksheet.WriteUsedFormatting(0, 0, [uffBold, uffNumberFormat]);
  MyWorksheet.WriteNumber(0, 1, 2.0);
  MyWorksheet.WriteNumber(0, 2, 3.0);
  MyWorksheet.WriteNumber(0, 3, 4.0);

  // Write the formula E1 = ABS(A1) as rpn token array
  SetLength(MyRPNFormula, 2);
  MyRPNFormula[0].ElementKind := fekCell;
  MyRPNFormula[0].Col := 0;
  MyRPNFormula[0].Row := 0;
  MyRPNFormula[1].ElementKind := fekFUNC;
  MyRPNFormula[1].FuncName := 'ABS';
  MyWorksheet.WriteRPNFormula(0, 4, MyRPNFormula);

  // Write the formula F1 = ROUND(A1, 0) as rpn token array
  SetLength(MyRPNFormula, 3);
  MyRPNFormula[0].ElementKind := fekCell;
  MyRPNFormula[0].Col := 0;
  MyRPNFormula[0].Row := 0;
  MyRPNFormula[1].ElementKind := fekNum;
  MyRPNFormula[1].DoubleValue := 0.0;
  MyRPNFormula[2].ElementKind := fekFUNC;
  MyRPNFormula[2].FuncName := 'ROUND';
  MyWorksheet.WriteRPNFormula(0, 5, MyRPNFormula);

  // Write a string formula to G1 = "A" & "B"
  MyWorksheet.WriteFormula(0, 6, '="A"&"B"');

  // Write string formula to H1 = sin(A1+B1)
  MyWorksheet.WriteFormula(0, 7, '=SIN(A1+B1)');

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

  r:= 10;
  // Write current date/time and test numbers for various formatting options

  MyWorksheet.WriteUTF8Text(r, 1, 'Formats in gray cells are not supported by BIFF2');

  inc(r, 2);
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
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'dd/mmm''');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, ''mmm/yy''');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'mmm/yy');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfShortTimeAM');
  MyWorksheet.WriteDateTime(r, 1, now, nfShortTimeAM);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfLongTimeAM');
  MyWorksheet.WriteDateTime(r, 1, now, nfLongTimeAM);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, nn:ss');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'nn:ss');
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, nn:ss.z');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'nn:ss.z');
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, mm:ss.zzz');
  MyWorksheet.WriteDateTime(r, 1, now, nfCustom, 'mm:ss.zzz');
  MyWorksheet.WriteFontColor(r, 1, scGray);

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
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 1);
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixed, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixed, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixed, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixed, 3);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixed, 3);
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 0);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 0);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 1 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 1);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 1);
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfFixedTh, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfFixedTh, 3);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number, nfFixedTh, 3);
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 1 dec');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 1);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 1);
  MyWorksheet.WriteFontColor(r, 2, scGray);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 1);
  MyWorksheet.WriteFontColor(r, 3, scGray);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 1);
  MyWorksheet.WriteFontColor(r, 4, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 2);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 2);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 2);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfExp, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfExp, 3);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number, nfExp, 3);
  MyWorksheet.WriteFontColor(r, 2, scGray);
  MyWorksheet.WriteNumber(r, 3, 1.0/number, nfExp, 3);
  MyWorksheet.WriteFontColor(r, 3, scGray);
  MyWorksheet.WriteNumber(r, 4, -1.0/number, nfExp, 3);
  MyWorksheet.WriteFontColor(r, 4, scGray);
  inc(r,2);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrency, 0 decs');
  MyWorksheet.WriteCurrency(r, 1, number, nfCurrency, 0, '$');
  MyWorksheet.WriteCurrency(r, 2, -number, nfCurrency, 0, '$');
  MyWorksheet.WriteCurrency(r, 3, 0.0, nfCurrency, 0, '$');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCurrencyRed, 0 decs');
  MyWorksheet.WriteCurrency(r, 1, number, nfCurrencyRed, 0, 'USD');
  MyWorksheet.WriteCurrency(r, 2, -number, nfCurrencyRed, 0, 'USD');
  MyWorksheet.WriteCurrency(r, 3, 0.0, nfCurrencyRed, 0, 'USD');
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, "$"#,##0_);("$"#,##0)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '"$"#,##0_);("$"#,##0)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '"$"#,##0_);("$"#,##0)');
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, "$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '"$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '"$"#,##0.0_);[Red]("$"#,##0.0)');
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  fmt := '"€"#,##0.0_);[Red]("€"#,##0.0)';
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, '+fmt);
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, UTF8ToAnsi(fmt));
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, UTF8ToAnsi(fmt));
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  fmt := '[Green]"¥"#,##0.0_);[Red]-"¥"#,##0.0';
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, '+fmt);
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, UTF8ToAnsi(fmt));
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, UTF8ToAnsi(fmt));
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfCustom, _("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)');
  MyWorksheet.WriteNumber(r, 1, number);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  MyWorksheet.WriteNumberFormat(r, 1, nfCustom, '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)');
  MyWorksheet.WriteNumber(r, 2, -number);
  MyWorksheet.WriteNumberFormat(r, 2, nfCustom, '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)');
  MyWorksheet.WriteFontColor(r, 2, scGray);
  inc(r, 2);
  number := 1.333333333;
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 0 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 0);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 1 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 1);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 2 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 2);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfPercentage, 3 decs');
  MyWorksheet.WriteNumber(r, 1, number, nfPercentage, 3);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, hh:mm:ss');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval);
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h:m:s');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'H:M:s');
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, hh:mm');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'hh:mm');
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h:m');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'h:m');
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);
  MyWorksheet.WriteUTF8Text(r, 0, 'nfTimeInterval, h');
  MyWorksheet.WriteDateTime(r, 1, number, nfTimeInterval, 'h');
  MyWorksheet.WriteFontColor(r, 1, scGray);
  inc(r);

  // Set width of columns 0 to 3
  MyWorksheet.WriteColWidth(0, 48);    // 48 characters, default is 12 --> 4x default width
  lCol.Width := 24;                    // 24 characters, default is 12 --> 2x default width
  MyWorksheet.WriteColInfo(1, lCol);
  MyWorksheet.WriteColInfo(2, lCol);
  MyWorksheet.WriteColInfo(3, lCol);

  // Set height of rows 5 and 6
  lRow.Height := 4;                    // 4 lines
  MyWorksheet.WriteRowInfo(5, lRow);
  lRow.Height := 2;                    // 2 lines
  MyWorksheet.WriteRowInfo(6, lRow);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test' + STR_EXCEL_EXTENSION, sfExcel2, true);
  MyWorkbook.Free;
end.

