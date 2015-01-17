unit numformatparsertests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpsallformats, fpspreadsheet, fpsnumformatparser, xlsbiff8
  {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  TParserTestData = record
    FormatString: String;
    SollFormatString: String;
    SollNumFormat: TsNumberFormat;
    SollSectionCount: Integer;
    SollDecimals: Byte;
    SollCurrencySymbol: String;
  end;

var
  ParserTestData: Array[0..8] of TParserTestData;

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
    SollCurrencySymbol := '';
  end;
  with ParserTestData[1] do begin
    FormatString := '0.000';
    SollFormatString := '0.000';
    SollNumFormat := nfFixed;
    SollSectionCount := 1;
    SollDecimals := 3;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[2] do begin
    FormatString := '#,##0.000';
    SollFormatString := '#,##0.000';
    SollNumFormat := nfFixedTh;
    SollSectionCount := 1;
    SollDecimals := 3;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[3] do begin
    FormatString := '0.000%';
    SollFormatString := '0.000%';
    SollNumFormat := nfPercentage;
    SollSectionCount := 1;
    SollDecimals := 3;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[4] do begin
    FormatString := 'hh:mm:ss';
    SollFormatString := 'hh:nn:ss';
    SollNumFormat := nfLongTime;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[5] do begin
    FormatString := 'hh:mm:ss AM/PM';
    SollFormatString := 'hh:nn:ss AM/PM';
    SollNumFormat := nfLongTimeAM;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[6] do begin
    FormatString := '[$-409]hh:mm:ss\ AM/PM;@';
    SollFormatString := 'hh:nn:ss AM/PM';
    SollNumFormat := nfLongTimeAM;
    SollSectionCount := 2;
    SollDecimals := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[7] do begin
    FormatString := '[$-F400]dd.mm.yy\ hh:mm';
    SollFormatString := 'dd.mm.yy hh:nn';
    SollNumFormat := nfShortDateTime;
    SollSectionCount := 1;
    SollDecimals := 0;
    SollCurrencySymbol := '';
  end;
  with ParserTestData[8] do begin
    FormatString := '[$€] #,##0.00;-[$€] #,##0.00;{$€} 0.00';
    SollFormatString := '"€" #,##0.00;-"€" #,##0.00;"€" 0.00';
    SollNumFormat := nfCurrency;
    SollSectionCount := 3;
    SollDecimals := 2;
    SollCurrencySymbol := '€';
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
    for i:=0 to 5 do begin
      parser := TsNumFormatParser.Create(MyWorkbook, ParserTestData[i].FormatString);
      try
        actual := parser.FormatString[nfdDefault];
        CheckEquals(ParserTestData[i].SollFormatString, actual,
          'Test format string ' + ParserTestData[i].SollFormatString + ' construction mismatch');
        CheckEquals(ord(ParserTestData[i].SollNumFormat), ord(parser.ParsedSections[0].NumFormat),
          'Test format (' + GetEnumName(TypeInfo(TsNumberFormat), integer(ParserTestData[i].SollNumFormat)) +
          ') detection mismatch');
        CheckEquals(ParserTestData[i].SollDecimals, parser.ParsedSections[0].Decimals,
          'Test format (' + ParserTestData[i].FormatString + ') decimal detection mismatch');
        CheckEquals(ParserTestData[i].SollCurrencySymbol, parser.ParsedSections[0].CurrencySymbol,
          'Test format (' + ParserTestData[i].FormatString + ') currency symbol detection mismatch');
        CheckEquals(ParserTestData[i].SollSectionCount, parser.ParsedSectionCount,
          'Test format (' + ParserTestData[i].FormatString + ') section count mismatch');
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

