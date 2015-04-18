{
opendocwrite.lpr

Demonstrates how to write an OpenDocument file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program opendocwrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, fpsallformats;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  number1, number2, number3, number4,
  number5, number6, number7, number8: Double;
  dt1, dt2: TDateTime;
  row: Integer = 7;
begin
  MyDir := ExtractFilePath(ParamStr(0));
  number1 := 1.23456789;
  number2 := -number1;
  number3 := 0.123456789;
  number4 := -number3;
  number5 := 10000*number1;
  number6 := -10000*number1;
  number7 := 1/number3;
  number8 := -1/number3;

  dt1 := EncodeDate(2012, 1, 1) + EncodeTime(9, 1, 2, 12);
  dt2 := EncodeDate(2012, 12, 1) + EncodeTime(21, 1, 2, 12);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  // Write some cells
  MyWorksheet.WriteNumber(0, 0, 1.0);       // A1
  MyWorksheet.WriteNumber(0, 1, 2.0);       // B1
  MyWorksheet.WriteNumber(0, 2, 3.0);       // C1
  MyWorksheet.WriteNumber(0, 3, 4.0);       // D1
  MyWorksheet.WriteUTF8Text(4, 2, 'Total:');// C5
  MyWorksheet.WriteNumber(4, 3, 10);        // D5
  MyWorksheet.WriteDateTime(5, 0, now);

  // Add some formatting
  MyWorksheet.WriteFontStyle(0, 0, [fssBold]);
  MyWorksheet.WriteFont(0, 1, 'Times New Roman', 16, [], scRed);

  // Show number formats
  MyWorksheet.WriteUTF8Text(row, 0, 'Number formats:');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfGeneral');
  MyWorksheet.WriteNumber(row, 1, number1, nfGeneral);
  MyWorksheet.WriteNumber(row, 2, number2, nfGeneral);
  MyWorksheet.WriteNumber(row, 3, number3, nfGeneral);
  MyWorksheet.WriteNumber(row, 4, number4, nfGeneral);
  MyWorksheet.WriteNumber(row, 5, number5, nfGeneral);
  MyWorksheet.WriteNumber(row, 6, number6, nfGeneral);
  MyWorksheet.WriteNumber(row, 7, number7, nfGeneral);
  MyWorksheet.WriteNumber(row, 8, number8, nfGeneral);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFixed, 0 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 2, number2, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 3, number3, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 4, number4, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 5, number5, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 6, number6, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 7, number7, nfFixed, 0);
  MyWorksheet.WriteNumber(row, 8, number8, nfFixed, 0);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFixed, 2 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 2, number2, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 3, number3, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 4, number4, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 5, number5, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 6, number6, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 7, number7, nfFixed, 2);
  MyWorksheet.WriteNumber(row, 8, number8, nfFixed, 2);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFixed, 3 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 2, number2, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 3, number3, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 4, number4, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 5, number5, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 6, number6, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 7, number7, nfFixed, 3);
  MyWorksheet.WriteNumber(row, 8, number8, nfFixed, 3);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFixedTh, 0 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 2, number2, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 3, number3, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 4, number4, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 5, number5, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 6, number6, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 7, number7, nfFixedTh, 0);
  MyWorksheet.WriteNumber(row, 8, number8, nfFixedTh, 0);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFixedTh, 2 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 2, number2, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 3, number3, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 4, number4, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 5, number5, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 6, number6, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 7, number7, nfFixedTh, 2);
  MyWorksheet.WriteNumber(row, 8, number8, nfFixedTh, 2);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFixedTh, 3 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 2, number2, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 3, number3, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 4, number4, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 5, number5, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 6, number6, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 7, number7, nfFixedTh, 3);
  MyWorksheet.WriteNumber(row, 8, number8, nfFixedTh, 3);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfPercentage, 0 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 2, number2, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 3, number3, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 4, number4, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 5, number5, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 6, number6, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 7, number7, nfPercentage, 0);
  MyWorksheet.WriteNumber(row, 8, number8, nfPercentage, 0);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfPercentage, 2 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 2, number2, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 3, number3, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 4, number4, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 5, number5, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 6, number6, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 7, number7, nfPercentage, 2);
  MyWorksheet.WriteNumber(row, 8, number8, nfPercentage, 2);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfPercentage, 3 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 2, number2, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 3, number3, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 4, number4, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 5, number5, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 6, number6, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 7, number7, nfPercentage, 3);
  MyWorksheet.WriteNumber(row, 8, number8, nfPercentage, 3);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfExp, 0 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfExp, 0);
  MyWorksheet.WriteNumber(row, 2, number2, nfExp, 0);
  MyWorksheet.WriteNumber(row, 3, number3, nfExp, 0);
  MyWorksheet.WriteNumber(row, 4, number4, nfExp, 0);
  MyWorksheet.WriteNumber(row, 5, number5, nfExp, 0);
  MyWorksheet.WriteNumber(row, 6, number6, nfExp, 0);
  MyWorksheet.WriteNumber(row, 7, number7, nfExp, 0);
  MyWorksheet.WriteNumber(row, 8, number8, nfExp, 0);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfExp, 2 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfExp, 2);
  MyWorksheet.WriteNumber(row, 2, number2, nfExp, 2);
  MyWorksheet.WriteNumber(row, 3, number3, nfExp, 2);
  MyWorksheet.WriteNumber(row, 4, number4, nfExp, 2);
  MyWorksheet.WriteNumber(row, 5, number5, nfExp, 2);
  MyWorksheet.WriteNumber(row, 6, number6, nfExp, 2);
  MyWorksheet.WriteNumber(row, 7, number7, nfExp, 2);
  MyWorksheet.WriteNumber(row, 8, number8, nfExp, 2);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfExp, 3 decimals');
  MyWorksheet.WriteNumber(row, 1, number1, nfExp, 3);
  MyWorksheet.WriteNumber(row, 2, number2, nfExp, 3);
  MyWorksheet.WriteNumber(row, 3, number3, nfExp, 3);
  MyWorksheet.WriteNumber(row, 4, number4, nfExp, 3);
  MyWorksheet.WriteNumber(row, 5, number5, nfExp, 3);
  MyWorksheet.WriteNumber(row, 6, number6, nfExp, 3);
  MyWorksheet.WriteNumber(row, 7, number7, nfExp, 3);
  MyWorksheet.WriteNumber(row, 8, number8, nfExp, 3);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCurrency, 2 decimals');
  MyWorksheet.WriteCurrency(row, 1, number1, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 2, number2, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 3, number3, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 4, number4, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 5, number5, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 6, number6, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 7, number7, nfCurrency, 2, '$');
  MyWorksheet.WriteCurrency(row, 8, number8, nfCurrency, 2, '$');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCurrencyRed, 2 decimals, >0: $ 1000, <0: ($ 1000)');
  MyWorksheet.WriteCurrency(row, 1, number1, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 2, number2, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 3, number3, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 4, number4, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 5, number5, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 6, number6, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 7, number7, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  MyWorksheet.WriteCurrency(row, 8, number8, nfCurrencyRed, 2, '$', pcfCSV, ncfBCSVB);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfFraction, 2 digits');
  MyWorksheet.WriteNumber(row, 1, number1, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 2, number2, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 3, number3, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 4, number4, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 5, number5, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 6, number6, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 7, number7, nfFraction, '# ???/???');
  MyWorksheet.WriteNumber(row, 8, number8, nfFraction, '# ???/???');
  inc(row,2);

  MyWorksheet.WriteUTF8Text(row, 0, 'Some date/time values in various formats:');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfShortDateTime');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfShortDateTime);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfShortDateTime);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfShortDate');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfShortDate);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfShortDate);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfLongDate');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfLongDate);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfLongDate);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfShortTime');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfShortTime);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfShortTime);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfLongTime');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfLongTime);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfLongTime);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfShortTimeAM');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfShortTimeAM);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfShortTimeAM);
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfLongTimeAM');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfLongTimeAM);
  MyWorksheet.WriteDateTime(row, 2, dt2, nfLongTimeAM);
  inc(row,2);

  MyWorksheet.WriteUTF8Text(row, 0, 'Some custom formats');
  inc(row);
  // In order to use a semicolon as a date-time separator it must be escaped either by
  // using the backslash or quotes (because the semicolon is the separator between sections)
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCustom, dddd, dd/mm/yyyy\; hh:nn');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfCustom, 'dddd, dd/mm/yyyy\; hh:nn');
  MyWorksheet.WriteDateTime(row, 2, dt2, nfCustom, 'dddd, dd/mm/yyyy\; hh:nn');
  MyWorksheet.WriteUTF8Text(row, 3, 'The semicolon must be escaped otherwise it would be misunderstood as a section separator.');
  MyWorksheet.WriteUTF8Text(row, 4, 'This format is not displayed correctly by Open/LibreOffice.');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCustom, dddd, dd/mm/yyyy"; "hh:nn');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfCustom, 'dddd, dd/mm/yyyy"; "hh:nn');
  MyWorksheet.WriteDateTime(row, 2, dt2, nfCustom, 'dddd, dd/mm/yyyy"; "hh:nn');
  MyWorksheet.WriteUTF8Text(row, 3, 'The semicolon must be escaped otherwise it would be misunderstood as a section separator.');
  MyWorksheet.WriteUTF8Text(row, 4, 'This format is not displayed correctly by Open/LibreOffice.');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCustom, dd/mmm');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfCustom, 'dd/mmm');
  MyWorksheet.WriteDateTime(row, 2, dt2, nfCustom, 'dd/mmm');
  MyWorksheet.WriteUTF8Text(row, 3, 'The slash is replaced by the date or time separator of the FormatSettings');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCustom, mmm/yy');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfCustom, 'mmm/yy');
  MyWorksheet.WriteDateTime(row, 2, dt2, nfCustom, 'mmm/yy');
  MyWorksheet.WriteUTF8Text(row, 3, 'The slash is replaced by the date or time separator of the FormatSettings');
  inc(row);
  MyWorksheet.WriteUTF8Text(row, 0, 'nfCustom, mmm-yy');
  MyWorksheet.WriteDateTime(row, 1, dt1, nfCustom, 'mmm-yy');
  MyWorksheet.WriteDateTime(row, 2, dt2, nfCustom, 'mmm-yy');
  MyWorksheet.WriteUTF8Text(row, 3, 'The dash is used literally');

  // Creates a new worksheet
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet 2');

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.ods',
    sfOpenDocument);
  MyWorkbook.Free;
end.

