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
  number: Double;
  lCol: TCol;
  lRow: TRow;
begin
  // Open the output file
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  //MyWorksheet.WriteColWidth(0, 5);
  //MyWorksheet.WriteColWidth(1, 30);

  // Write some number cells
  MyWorksheet.WriteNumber(0, 0, 1.0);
  MyWorksheet.WriteUsedFormatting(0, 0, [uffBold, uffNumberFormat]);
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

  // Set width of columns 0 and 1
  MyWorksheet.WriteColWidth(0, 40);
  lCol.Width := 35;
  MyWorksheet.WriteColInfo(1, lCol);

  // Set height of rows 5 and 6
  lRow.Height := 10;
  MyWorksheet.WriteRowInfo(5, lRow);
  lRow.Height := 5;
  MyWorksheet.WriteRowInfo(6, lRow);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test' + STR_EXCEL_EXTENSION, sfExcel2, true);
  MyWorkbook.Free;
end.

