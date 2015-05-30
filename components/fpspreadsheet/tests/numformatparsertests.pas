unit numformatparsertests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, fpsnumformatparser
  {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  TParserTestData = record
    FormatString: String;
    SollFormatString: String;
    SollNumFormat: TsNumberFormat;
    SollSectionCount: Integer;
    SollDecimals: Byte;
    SollFactor: Double;
    SollNumeratorDigits: Integer;
    SollDenominatorDigits: Integer;
    SollCurrencySymbol: String;
    SollSection2Color: TsColor;
  end;

var
  ParserTestData: Array[0..13] of TParserTestData;

procedure InitParserTestData;

type
  TSpreadNumFormatParserTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // Reads numbers values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
  published
    procedure TestNumFormatParser;
  end;


implementation

uses
  TypInfo;

{ The test will use Excel strings and convert them to fpc dialect }
procedure InitParserTestData;
begin
  // Tests with 1 format section only
  with ParserTestData[0] do begin
    FormatString := '0';
    SollFormatString := '0';
    SollNumFormat := nfFixed;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[1] do begin
    FormatString := '0.000';
    SollFormatString := '0.000';
    SollNumFormat := nfFixed;
    SollSectionCount := 1;
    SollDecimals := 3;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[2] do begin
    FormatString := '#,##0.000';
    SollFormatString := '#,##0.000';
    SollNumFormat := nfFixedTh;
    SollSectionCount := 1;
    SollDecimals := 3;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[3] do begin
    FormatString := '0.000%';
    SollFormatString := '0.000%';
    SollNumFormat := nfPercentage;
    SollSectionCount := 1;
    SollDecimals := 3;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[4] do begin
    FormatString := 'hh:mm:ss';
    SollFormatString := 'hh:mm:ss';
    SollNumFormat := nfLongTime;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[5] do begin
    FormatString := 'hh:mm:ss AM/PM';
    SollFormatString := 'hh:mm:ss AM/PM';
    SollNumFormat := nfLongTimeAM;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[6] do begin
    FormatString := '[$-409]hh:mm:ss\ AM/PM;@';
    SollFormatString := 'hh:mm:ss\ AM/PM;@';
    SollNumFormat := nfCustom;
    SollSectionCount := 2;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[7] do begin
    FormatString := '[$-F400]dd.mm.yy\ hh:mm';
    SollFormatString := 'DD.MM.YY\ hh:mm';
    SollNumFormat := nfCustom;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[8] do begin
    FormatString := '[$€] #,##0.00;-[$€] #,##0.00;[$€] 0.00';
    SollFormatString := '[$€] #,##0.00;-[$€] #,##0.00;[$€] 0.00';
    SollNumFormat := nfCurrency;
    SollSectionCount := 3;
    SollDecimals := 2;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '€';
    SollSection2Color := scBlack;
  end;
  with ParserTestData[9] do begin
    FormatString := '[$€] #,##0.00;[red]-[$€] #,##0.00;[$€] 0.00';
    SollFormatString := '[$€] #,##0.00;[red]-[$€] #,##0.00;[$€] 0.00';
    SollNumFormat := nfCurrencyRed;
    SollSectionCount := 3;
    SollDecimals := 2;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '€';
    SollSection2Color := scRed;
  end;
  with ParserTestData[10] do begin
    FormatString := '0.00,,';
    SollFormatString := '0.00,,';
    SollNumFormat := nfCustom;
    SollSectionCount := 1;
    SollDecimals := 2;
    SollFactor := 1e-6;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[11] do begin
    FormatString := '# ??/??';
    SollFormatString := '# ??/??';
    SollNumFormat := nfFraction;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 2;
    SollDenominatorDigits := 2;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[12] do begin
    FormatString := 'General;[Red]-General';
    SollFormatString := 'General;[red]-General';
    SollNumFormat := nfCustom;
    SollSectionCount := 2;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
    SollSection2Color := scRed;
  end;
  with ParserTestData[13] do begin
    FormatString := 'General';
    SollFormatString := 'General';
    SollNumFormat := nfGeneral;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollFactor := 0;
    SollNumeratorDigits := 0;
    SollDenominatorDigits := 0;
    SollCurrencySymbol := '';
  end;

  {
  with ParserTestData[5] do begin
    FormatString := '#,##0.00 "$";-#,##0.00 "$";-';
    SollFormatString := '#,##0.00 "$";-#,##0.00 "$";-';
    SollNumFormat := nfCurrencyDash;
    SollSectionCount := 3;
    SollDecimals := 2;
    SollCurrencySymbol := '$';
  end;                            }

  {
  // This case will report a mismatching FormatString because of the [RED] --> ignore
  with ParserTestData[6] do begin
    FormatString := '#,##0.00 "$";[RED]-#,##0.00 "$";-';
    SollFormatString := '#,##0.00 "$";-#,##0.00 "$";-';
    SollNumFormat := nfCurrencyDashRed;
    SollSectionCount := 3;
    SollDecimals := 2;
    SollCurrencySymbol := '$';
  end;
  }
end;

{ TSpreadNumFormatParserTests }

procedure TSpreadNumFormatParserTests.SetUp;
begin
  inherited SetUp;
  InitParserTestData;
end;

procedure TSpreadNumFormatParserTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadNumFormatParserTests.TestNumFormatParser;
var
  i: Integer;
  parser: TsNumFormatParser;
  MyWorkbook: TsWorkbook;
  actual: String;
begin
  MyWorkbook := TsWorkbook.Create;  // needed to provide the FormatSettings for the parser
  try
    for i:=0 to High(ParserTestData) do begin
      parser := TsNumFormatParser.Create(MyWorkbook, ParserTestData[i].FormatString);
      try
        actual := parser.FormatString;
        CheckEquals(ParserTestData[i].SollFormatString, actual,
          'Test format string ' + ParserTestData[i].SollFormatString + ' construction mismatch');
        CheckEquals(
          GetEnumName(TypeInfo(TsNumberFormat), ord(ParserTestData[i].SollNumFormat)),
          GetEnumName(TypeInfo(TsNumberformat), ord(parser.ParsedSections[0].NumFormat)),
          'Test format (' + ParserTestData[i].FormatString + ') detection mismatch');
        CheckEquals(ParserTestData[i].SollDecimals, parser.ParsedSections[0].Decimals,
          'Test format (' + ParserTestData[i].FormatString + ') decimal detection mismatch');
        CheckEquals(ParserTestData[i].SollCurrencySymbol, parser.ParsedSections[0].CurrencySymbol,
          'Test format (' + ParserTestData[i].FormatString + ') currency symbol detection mismatch');
        CheckEquals(ParserTestData[i].SollSectionCount, parser.ParsedSectionCount,
          'Test format (' + ParserTestData[i].FormatString + ') section count mismatch');
        CheckEquals(ParserTestData[i].SollFactor, parser.ParsedSections[0].Factor,
          'Test format (' + ParserTestData[i].FormatString + ') factor mismatch');
        CheckEquals(ParserTestData[i].SollNumeratorDigits, parser.ParsedSections[0].FracNumerator,
          'Test format (' + ParserTestData[i].FormatString + ') numerator digits mismatch');
        CheckEquals(ParserTestData[i].SollDenominatorDigits, parser.ParsedSections[0].FracDenominator,
          'Test format (' + ParserTestData[i].FormatString + ') denominator digits mismatch');
        if ParserTestData[i].SollSectionCount > 1 then
          CheckEquals(ParserTestData[i].SollSection2Color, parser.ParsedSections[1].Color,
            'Test format (' + ParserTestData[i].FormatString + ') section 2 color mismatch');
      finally
        parser.Free;
      end;
    end;
  finally
    MyWorkbook.Free;
  end;
end;

initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadNumFormatParserTests);
  InitParserTestData; //useful to have norm data if other code want to use this unit
end.

end.

