unit formattests;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface
{ Formatted date/time/number tests
This unit tests writing out to and reading back from files.
Tests that verify reading from an Excel/LibreOffice/OpenOffice file are located in other
units (e.g. datetests).
}

uses
  {$IFDEF Unix}
  //required for formatsettings
  clocale,
  {$ENDIF}
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry, testsutility,
  fpstypes, fpsallformats, fpspreadsheet, fpscell, xlsbiff8;

var
  // Norm to test against - list of strings that should occur in spreadsheet
  SollNumberStrings: array[0..6, 0..9] of string;
  SollNumbers: array[0..6] of Double;
  SollNumberFormats: array[0..9] of TsNumberFormat;
  SollNumberDecimals: array[0..9] of word;

  SollDateTimeStrings: array[0..4, 0..8] of string;
  SollDateTimes: array[0..4] of TDateTime;
  SollDateTimeFormats: array[0..8] of TsNumberFormat;
  SollDateTimeFormatStrings: array[0..8] of String;

  SollColWidths: array[0..1] of Single;
  SollRowHeights: Array[0..2] of Single;
  SollBorders: array[0..19] of TsCellBorders;
  SollBorderLineStyles: array[0..6] of TsLineStyle;
  SollBorderColors: array[0..5] of TsColor;

  procedure InitSollFmtData;

type
  { TSpreadWriteReadFormatTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadFormatTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;

    // Test alignments
    procedure TestWriteRead_Alignment(AFormat: TsSpreadsheetFormat);
    // Test background
    procedure TestWriteRead_Background(AFormat: TsSpreadsheetFormat);
    // Test border
    procedure TestWriteRead_Border(AFormat: TsSpreadsheetFormat);
    // Test border styles
    procedure TestWriteRead_BorderStyles(AFormat: TsSpreadsheetFormat);
    // Test column widths
    procedure TestWriteRead_ColWidths(AFormat: TsSpreadsheetFormat);
    // Test row heights
    procedure TestWriteRead_RowHeights(AFormat: TsSpreadsheetFormat);
    // Test text rotation
    procedure TestWriteRead_TextRotation(AFormat:TsSpreadsheetFormat);
    // Test word wrapping
    procedure TestWriteRead_WordWrap(AFormat: TsSpreadsheetFormat);
    // Test number formats
    procedure TestWriteRead_NumberFormats(AFormat: TsSpreadsheetFormat;
      AVariant: Integer = 0);
    // Repeat with date/times
    procedure TestWriteRead_DateTimeFormats(AFormat: TsSpreadsheetFormat);
    // Test merged cells
    procedure TestWriteRead_MergedCells(AFormat: TsSpreadsheetFormat);
    // Many XF records
    procedure TestWriteRead_ManyXF(AFormat: TsSpreadsheetFormat);

  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.

    { BIFF2 Tests }
    procedure TestWriteRead_BIFF2_Alignment;
    procedure TestWriteRead_BIFF2_Border;
    procedure TestWriteRead_BIFF2_ColWidths;
    procedure TestWriteRead_BIFF2_RowHeights;
    procedure TestWriteRead_BIFF2_DateTimeFormats;
    procedure TestWriteRead_BIFF2_MergedCells;
    procedure TestWriteRead_BIFF2_NumberFormats;
    procedure TestWriteRead_BIFF2_ManyXFRecords;
    // These features are not supported by Excel2 --> no test cases required!
    // - Background
    // - BorderStyle
    // - TextRotation
    // - Wordwrap

    { BIFF5 Tests }
    procedure TestWriteRead_BIFF5_Alignment;
    procedure TestWriteRead_BIFF5_Background;
    procedure TestWriteRead_BIFF5_Border;
    procedure TestWriteRead_BIFF5_BorderStyles;
    procedure TestWriteRead_BIFF5_ColWidths;
    procedure TestWriteRead_BIFF5_RowHeights;
    procedure TestWriteRead_BIFF5_DateTimeFormats;
    procedure TestWriteRead_BIFF5_MergedCells;
    procedure TestWriteRead_BIFF5_NumberFormats;
    procedure TestWriteRead_BIFF5_TextRotation;
    procedure TestWriteRead_BIFF5_WordWrap;

    { BIFF8 Tests }
    procedure TestWriteRead_BIFF8_Alignment;
    procedure TestWriteRead_BIFF8_Background;
    procedure TestWriteRead_BIFF8_Border;
    procedure TestWriteRead_BIFF8_BorderStyles;
    procedure TestWriteRead_BIFF8_ColWidths;
    procedure TestWriteRead_BIFF8_RowHeights;
    procedure TestWriteRead_BIFF8_DateTimeFormats;
    procedure TestWriteRead_BIFF8_MergedCells;
    procedure TestWriteRead_BIFF8_NumberFormats;
    procedure TestWriteRead_BIFF8_TextRotation;
    procedure TestWriteRead_BIFF8_WordWrap;

    { ODS Tests }
    procedure TestWriteRead_ODS_Alignment;
    // no background patterns in ods
    procedure TestWriteRead_ODS_Border;
    procedure TestWriteRead_ODS_BorderStyles;
    procedure TestWriteRead_ODS_ColWidths;
    procedure TestWriteRead_ODS_RowHeights;
    procedure TestWriteRead_ODS_DateTimeFormats;
    procedure TestWriteRead_ODS_MergedCells;
    procedure TestWriteRead_ODS_NumberFormats;
    procedure TestWriteRead_ODS_TextRotation;
    procedure TestWriteRead_ODS_WordWrap;

    { OOXML Tests }
    procedure TestWriteRead_OOXML_Alignment;
    procedure TestWriteRead_OOXML_Background;
    procedure TestWriteRead_OOXML_Border;
    procedure TestWriteRead_OOXML_BorderStyles;
    procedure TestWriteRead_OOXML_ColWidths;
    procedure TestWriteRead_OOXML_RowHeights;
    procedure TestWriteRead_OOXML_DateTimeFormats;
    procedure TestWriteRead_OOXML_MergedCells;
    procedure TestWriteRead_OOXML_NumberFormats;
    procedure TestWriteRead_OOXML_TextRotation;
    procedure TestWriteRead_OOXML_WordWrap;

    { CSV Tests }
    procedure TestWriteRead_CSV_DateTimeFormats;
    procedure TestWriteRead_CSV_NumberFormats_0;
    procedure TestWriteRead_CSV_NumberFormats_1;
  end;

implementation

uses
  TypInfo, fpsPatches, fpsutils, fpsnumformat, fpspalette, fpscsv;

const
  FmtNumbersSheet = 'NumbersFormat'; //let's distinguish it from the regular numbers sheet
  FmtDateTimesSheet = 'DateTimesFormat';
  ColWidthSheet = 'ColWidths';
  RowHeightSheet = 'RowHeights';
  BackgroundSheet = 'Background';
  BordersSheet = 'CellBorders';
  AlignmentSheet = 'TextAlignments';
  TextRotationSheet = 'TextRotation';
  WordwrapSheet = 'Wordwrap';

// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollFmtData;
var
  i: Integer;
  fs: TFormatSettings;
  myworkbook: TsWorkbook;
  ch: Char;
begin
  // Set up norm - MUST match spreadsheet cells exactly

  myWorkbook := TsWorkbook.Create;
  try
    // There are some inconsistencies in fpc number-to-string conversions regarding
    // thousand separators and usage of - sign for very small numbers.
    // Therefore, we force the currency format to a given specification and build
    // the expected string accordingly.
    MyWorkbook.FormatSettings.CurrencyString := '€';  // use € for checking UTF8 issues
    // To get matching results also for Excel2 let't use its currency-value sequence.
    MyWorkbook.FormatSettings.Currencyformat := pcfCV;  // €100
    Myworkbook.FormatSettings.NegCurrFormat := ncfBCVB;  // (€100)
    fs := MyWorkbook.FormatSettings;
  finally
    myWorkbook.Free;
  end;

  // Numbers
  SollNumbers[0] := 0.0;
  SollNumbers[1] := 1.0;
  SollNumbers[2] := -1.0;
  SollNumbers[3] :=  1.23456E6;
  SollNumbers[4] := -1.23456E6;
  SollNumbers[5] :=  1.23456E-6;
  SollNumbers[6] := -1.23456E-6;

  SollNumberFormats[0] := nfGeneral;      SollNumberDecimals[0] := 0;
  SollNumberFormats[1] := nfFixed;        SollNumberDecimals[1] := 0;
  SollNumberFormats[2] := nfFixed;        SollNumberDecimals[2] := 2;
  SollNumberFormats[3] := nfFixedTh;      SollNumberDecimals[3] := 0;
  SollNumberFormats[4] := nfFixedTh;      SollNumberDecimals[4] := 2;
  SollNumberFormats[5] := nfExp;          SollNumberDecimals[5] := 2;
  SollNumberFormats[6] := nfPercentage;   SollNumberDecimals[6] := 0;
  SollNumberFormats[7] := nfPercentage;   SollNumberDecimals[7] := 2;
  SollNumberFormats[8] := nfCurrency;     SollNumberDecimals[8] := 0;
  SollNumberFormats[9] := nfCurrency;     SollNumberDecimals[9] := 2;

  SollNumberstrings[0, 0] := CurrToStrF(-1000.1, ffCurrency, 0, fs);

  for i:=Low(SollNumbers) to High(SollNumbers) do
  begin
    SollNumberStrings[i, 0] := FloatToStr(SollNumbers[i], fs);
    SollNumberStrings[i, 1] := FormatFloat('0', SollNumbers[i], fs);
    SollNumberStrings[i, 2] := FormatFloat('0.00', SollNumbers[i], fs);
    SollNumberStrings[i, 3] := FormatFloat('#,##0', SollNumbers[i], fs);
    SollNumberStrings[i, 4] := FormatFloat('#,##0.00', SollNumbers[i], fs);
    SollNumberStrings[i, 5] := FormatFloat('0.00E+00', SollNumbers[i], fs);
    SollNumberStrings[i, 6] := FormatFloat('0', SollNumbers[i]*100, fs) + '%';
    SollNumberStrings[i, 7] := FormatFloat('0.00', SollNumbers[i]*100, fs) + '%';
    {
    SollNumberStrings[i, 8] := FormatCurr('"€"#,##0;("€"#,##0)', SollNumbers[i], fs);
    SollNumberStrings[i, 9] := FormatCurr('"€"#,##0.00;("€"#,##0.00)', SollNumbers[i], fs);
    }
    // Don't use FormatCurr for the next two cases because is reports the sign of
    // very small numbers inconsistenly with the spreadsheet applications.
    SollNumberStrings[i, 8] := FormatFloat('"€"#,##0;("€"#,##0)', SollNumbers[i], fs);
    SollNumberStrings[i, 9] := FormatFloat('"€"#,##0.00;("€"#,##0.00)', SollNumbers[i], fs);
  end;

  // Date/time values
  SollDateTimes[0] := EncodeDate(2012, 1, 12) + EncodeTime(13, 14, 15, 567);
  SolLDateTimes[1] := EncodeDate(2012, 2, 29) + EncodeTime(0, 0, 0, 1);
  SollDateTimes[2] := EncodeDate(2040, 12, 31) + EncodeTime(12, 0, 0, 0);
  SollDateTimes[3] := 1 + EncodeTime(3,45, 0, 0);
  SollDateTimes[4] := EncodeTime(12, 0, 0, 0);

  SollDateTimeFormats[0] := nfShortDateTime;   SollDateTimeFormatStrings[0] := '';
  SollDateTimeFormats[1] := nfShortDate;       SollDateTimeFormatStrings[1] := '';
  SollDateTimeFormats[2] := nfShortTime;       SollDateTimeFormatStrings[2] := '';
  SollDateTimeFormats[3] := nfLongTime;        SollDateTimeFormatStrings[3] := '';
  SollDateTimeFormats[4] := nfShortTimeAM;     SollDateTimeFormatStrings[4] := '';
  SollDateTimeFormats[5] := nfLongTimeAM;      SollDateTimeFormatStrings[5] := '';
  SollDateTimeFormats[6] := nfCustom;          SollDateTimeFormatStrings[6] := 'dd/mmm';
  SolLDateTimeFormats[7] := nfCustom;          SollDateTimeFormatStrings[7] := 'mmm/yy';
  SollDateTimeFormats[8] := nfCustom;          SollDateTimeFormatStrings[8] := 'nn:ss';
//  SollDateTimeFormats[9] := nfTimeInterval;    SollDateTimeFormatStrings[9] := '';

  for i:=Low(SollDateTimes) to High(SollDateTimes) do
  begin
    SollDateTimeStrings[i, 0] := DateToStr(SollDateTimes[i], fs) + ' ' + FormatDateTime('t', SollDateTimes[i], fs);
    SollDateTimeStrings[i, 1] := DateToStr(SollDateTimes[i], fs);
    SollDateTimeStrings[i, 2] := FormatDateTime(fs.ShortTimeFormat, SollDateTimes[i], fs);
    SolLDateTimeStrings[i, 3] := FormatDateTime(fs.LongTimeFormat, SollDateTimes[i], fs);
    SollDateTimeStrings[i, 4] := FormatDateTime(fs.ShortTimeFormat + ' am/pm', SollDateTimes[i], fs);   // dont't use "t" - it does the hours wrong
    SollDateTimeStrings[i, 5] := FormatDateTime(fs.LongTimeFormat + ' am/pm', SollDateTimes[i], fs);
    SollDateTimeStrings[i, 6] := FormatDateTime(SpecialDateTimeFormat('dm', fs, false), SollDateTimes[i], fs);
    SollDateTimeStrings[i, 7] := FormatDateTime(SpecialDateTimeFormat('my', fs, false), SollDateTimes[i], fs);
    SollDateTimeStrings[i, 8] := FormatDateTime(SpecialDateTimeFormat('ms', fs, false), SollDateTimes[i], fs);
//    SollDateTimeStrings[i, 9] := FormatDateTime('[h]:mm:ss', SollDateTimes[i], fs, [fdoInterval]);
  end;

  // Column width
  SollColWidths[0] := 20;  // characters based on width of "0" of default font
  SollColWidths[1] := 40;

  // Row heights
  SollRowHeights[0] := 1;  // Lines of default font
  SollRowHeights[1] := 2;
  SollRowHeights[2] := 4;

  // Cell borders
  SollBorders[0] := [];
  SollBorders[1] := [cbEast];
  SollBorders[2] := [cbSouth];
  SollBorders[3] := [cbWest];
  SollBorders[4] := [cbNorth];
  SollBorders[5] := [cbEast, cbSouth];
  SollBorders[6] := [cbEast, cbWest];
  SollBorders[7] := [cbEast, cbNorth];
  SollBorders[8] := [cbSouth, cbWest];
  SollBorders[9] := [cbSouth, cbNorth];
  SollBorders[10] := [cbWest, cbNorth];
  SollBorders[11] := [cbEast, cbSouth, cbWest];
  SollBorders[12] := [cbEast, cbSouth, cbNorth];
  SollBorders[13] := [cbSouth, cbWest, cbNorth];
  SollBorders[14] := [cbWest, cbNorth, cbEast];
  SollBorders[15] := [cbEast, cbSouth, cbWest, cbNorth];     // BIFF2/5 end here
  SollBorders[16] := [cbDiagUp];
  SollBorders[17] := [cbDiagDown];
  SollBorders[18] := [cbDiagUp, cbDiagDown];
  SollBorders[19] := [cbEast, cbSouth, cbWest, cbNorth, cbDiagUp, cbDiagDown];

  SollBorderLineStyles[0] := lsThin;
  SollBorderLineStyles[1] := lsMedium;
  SollBorderLineStyles[2] := lsThick;
  SollBorderLineStyles[3] := lsDashed;
  SollBorderLineStyles[4] := lsDotted;
  SollBorderLineStyles[5] := lsDouble;
  SollBorderLineStyles[6] := lsHair;

  SollBorderColors[0] := scBlack;
  SollBorderColors[1] := scRed;
  SollBorderColors[2] := scBlue;
  SollBorderColors[3] := scGray;
  SollBorderColors[4] := scSilver;
  SollBorderColors[5] := scMagenta;
end;

{ TSpreadWriteReadFormatTests }

procedure TSpreadWriteReadFormatTests.SetUp;
begin
  inherited SetUp;
  InitSollFmtData; //just for security: make sure the variables are reset to default
end;

procedure TSpreadWriteReadFormatTests.TearDown;
begin
  inherited TearDown;
end;


{ --- Number format tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_NumberFormats(AFormat: TsSpreadsheetFormat;
  AVariant: Integer = 0);
{ AVariant specifies variants for csv:
  0 = decimal and thousand separator as in workbook's FormatSettings,
  1 = intercanged }
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
  ExpectedString: String;
  Row, Col: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.FormatSettings.CurrencyString := '€';  // use € for checking UTF8 issues
    MyWorkbook.FormatSettings.Currencyformat := pcfCV;  // €100
    Myworkbook.FormatSettings.NegCurrFormat := ncfBCVB;  // (€100)
    if (AFormat = sfCSV) then
    begin
      case AVariant of
        0: begin
             CSVParams.FormatSettings.DecimalSeparator := MyWorkbook.FormatSettings.DecimalSeparator;
             CSVParams.FormatSettings.ThousandSeparator := MyWorkbook.FormatSettings.ThousandSeparator;
           end;
        1: begin  // interchanged decimal and thousand separators
             CSVParams.FormatSettings.ThousandSeparator := MyWorkbook.FormatSettings.DecimalSeparator;
             CSVParams.FormatSettings.DecimalSeparator := MyWorkbook.FormatSettings.ThousandSeparator;
           end;
      end;
    end;

    MyWorkSheet:= MyWorkBook.AddWorksheet(FmtNumbersSheet);
    for Row := Low(SollNumbers) to High(SollNumbers) do
      for Col := ord(Low(SollNumberFormats)) to ord(High(SollNumberFormats)) do
      begin
        if IsCurrencyFormat(SollNumberFormats[Col]) then
          MyWorksheet.WriteCurrency(Row, Col, SollNumbers[Row], SollNumberFormats[Col], SollNumberDecimals[Col])
        else
          MyWorksheet.WriteNumber(Row, Col, SollNumbers[Row], SollNumberFormats[Col], SollNumberDecimals[Col]);
        ActualString := MyWorksheet.ReadAsUTF8Text(Row, Col);
        CheckEquals(SollNumberStrings[Row, Col], ActualString,
          'Test unsaved string mismatch, cell ' + CellNotation(MyWorksheet,Row,Col));
      end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.FormatSettings.CurrencyString := '€';  // use € for checking UTF8 issues
    MyWorkbook.FormatSettings.Currencyformat := pcfCV;   // €100
    Myworkbook.FormatSettings.NegCurrFormat := ncfBCVB;  // (€100)
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat in [sfExcel2, sfCSV] then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, FmtNumbersSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for Row := Low(SollNumbers) to High(SollNumbers) do
      for Col := Low(SollNumberFormats) to High(SollNumberFormats) do
      begin
        ActualString := MyWorkSheet.ReadAsUTF8Text(Row,Col);
        ExpectedString := SollNumberStrings[Row, Col];
        if (ExpectedString <> ActualString) then
        begin
          if (AFormat = sfCSV) and (Row=5) and (Col=0) then
            // CSV has an insignificant difference of tiny numbers in
            // general format
            ignore('Ignoring insignificant saved string mismatch, cell ' +
                   CellNotation(MyWorksheet,Row,Col) +
                   ', expected: <' + ExpectedString +
                   '> but was: <' + ActualString + '>')
          else
            CheckEquals(ExpectedString, ActualString,
              'Test saved string mismatch, cell '+CellNotation(MyWorkSheet,Row,Col));
        end;
      end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_NumberFormats;
begin
  TestWriteRead_NumberFormats(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_NumberFormats;
begin
  TestWriteRead_NumberFormats(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_NumberFormats;
begin
  TestWriteRead_NumberFormats(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_NumberFormats;
begin
  TestWriteRead_NumberFormats(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_NumberFormats;
begin
  TestWriteRead_NumberFormats(sfOOXML);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_CSV_NumberFormats_0;
begin
  TestWriteRead_NumberFormats(sfCSV, 0);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_CSV_NumberFormats_1;
begin
  TestWriteRead_NumberFormats(sfCSV, 1);
end;


{ --- Date/time formats --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_DateTimeFormats(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualString: String;
  Row,Col: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet(FmtDateTimesSheet);
    for Row := Low(SollDateTimes) to High(SollDateTimes) do
      for Col := Low(SollDateTimeFormats) to High(SollDateTimeFormats) do
      begin
        if (AFormat = sfExcel2) and (SollDateTimeFormats[Col] in [nfCustom, nfTimeInterval]) then
          Continue;  // The formats nfFmtDateTime and nfTimeInterval are not supported by BIFF2
        if (AFormat = sfCSV) and (SollDateTimeFormats[Col] in [nfCustom, nfTimeInterval]) then
          Continue;  // No chance for csv to detect custom formats without further information                                 MyWorksheet.WriteDateTime(Row, Col, SollDateTimes[Row], SollDateTimeFormats[Col], SollDateTimeFormatStrings[Col]);
        MyWorksheet.WriteDateTime(Row, Col, SollDateTimes[Row], SollDateTimeFormats[Col], SollDateTimeFormatStrings[Col]);
        ActualString := MyWorksheet.ReadAsUTF8Text(Row, Col);
        CheckEquals(
          Lowercase(SollDateTimeStrings[Row, Col]),
          Lowercase(ActualString),
          'Test unsaved string mismatch, cell ' + CellNotation(MyWorksheet,Row,Col)
        );
      end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat in [sfExcel2, sfCSV] then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkbook, FmtDateTimesSheet);
    if MyWorksheet = nil then
      fail('Error in test code. Failed to get named worksheet');
    for Row := Low(SollDateTimes) to High(SollDateTimes) do
      for Col := Low(SollDateTimeFormats) to High(SollDateTimeFormats) do
      begin
        if (AFormat = sfExcel2) and (SollDateTimeFormats[Col] in [nfCustom, nfTimeInterval]) then
          Continue;  // The formats nfFmtDateTime and nfTimeInterval are not supported by BIFF2
        if (AFormat = sfCSV) and (SollDateTimeFormats[Col] in [nfCustom, nfTimeInterval]) then
          Continue;  // No chance for csv to detect custom formats without further information                                 ActualString := MyWorksheet.ReadAsUTF8Text(Row,Col);
        ActualString := MyWorksheet.ReadAsUTF8Text(Row,Col);
        CheckEquals(
          Lowercase(SollDateTimeStrings[Row, Col]),
          Lowercase(ActualString),
          'Test saved string mismatch, cell '+CellNotation(MyWorksheet,Row,Col)
        );
      end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_DateTimeFormats;
begin
  TestWriteRead_DateTimeFormats(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_DateTimeFormats;
begin
  TestWriteRead_DateTimeFormats(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_DateTimeFormats;
begin
  TestWriteRead_DateTimeFormats(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_DateTimeFormats;
begin
  TestWriteRead_DateTimeFormats(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_DateTimeFormats;
begin
  TestWriteRead_DateTimeFormats(sfOOXML);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_CSV_DateTimeFormats;
begin
  TestWriteRead_DateTimeFormats(sfCSV);
end;


{ --- Alignment tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_Alignment(AFormat: TsSpreadsheetFormat);
const
  HORALIGN_TEXT: Array[TsHorAlignment] of String = ('haDefault', 'haLeft', 'haCenter', 'haRight');
  VERTALIGN_TEXT: Array[TsVertAlignment] of String = ('vaDefault', 'vaTop', 'vaCenter', 'vaBottom');
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  horAlign: TsHorAlignment;
  vertAlign: TsVertAlignment;
  row, col: Integer;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values: HorAlignments along rows, VertAlignments along columns
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(AlignmentSheet);

    col := 0;
    for horAlign in TsHorAlignment do
    begin
      row := 0;
      if AFormat = sfExcel2 then
      begin
        // BIFF2 can only do horizontal alignment --> no need for vertical alignment.
        MyWorksheet.WriteUTF8Text(row, col, HORALIGN_TEXT[horAlign]);
        MyWorksheet.WriteHorAlignment(row, col, horAlign);
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        CheckEquals(
          GetEnumName(TypeInfo(TsHorAlignment), Integer(horAlign)),
          GetEnumName(TypeInfo(TsHorAlignment), Integer(MyCell^.HorAlignment)),
          'Test unsaved horizontal alignment, cell ' + CellNotation(MyWorksheet,0,0)
        );
      end
      else
        for vertAlign in TsVertAlignment do
        begin
          MyWorksheet.WriteUTF8Text(row, col, HORALIGN_TEXT[horAlign]+'/'+VERTALIGN_TEXT[vertAlign]);
          MyWorksheet.WriteHorAlignment(row, col, horAlign);
          MyWorksheet.WriteVertAlignment(row, col, vertAlign);
          MyCell := MyWorksheet.FindCell(row, col);
          if MyCell = nil then
            fail('Error in test code. Failed to get cell.');
          CheckEquals(
            GetEnumName(TypeInfo(TsVertAlignment), Integer(vertAlign)),
            GetEnumName(TypeInfo(TsVertAlignment), Integer(MyCell^.VertAlignment)),
            'Test unsaved vertical alignment, cell ' + CellNotation(MyWorksheet,0,0)
          );
          CheckEquals(
            GetEnumName(TypeInfo(TsHorAlignment), Integer(horAlign)),
            GetEnumName(TypeInfo(TsHorAlignment), Integer(MyCell^.HorAlignment)),
            'Test unsaved horizontal alignment, cell ' + CellNotation(MyWorksheet,0,0)
          );
          inc(row);
        end;
      inc(col);
    end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then begin
      MyWorksheet := MyWorkbook.GetFirstWorksheet;
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');
      row := 0;
      for col :=0 to MyWorksheet.GetLastColIndex do
      begin
        MyCell := MyWorksheet.FindCell(row, col);
        if MyCell = nil then
          fail('Error in test code. Failed to get cell.');
        horAlign := TsHorAlignment(col);
        CheckEquals(
          GetEnumName(TypeInfo(TsHorAlignment), Integer(horAlign)),
          GetEnumName(TypeInfo(TsHorAlignment), Integer(MyCell^.HorAlignment)),
          'Test saved horizontal alignment mismatch, cell ' + CellNotation(MyWorksheet,row, col)
        );
      end
    end
    else begin
      MyWorksheet := GetWorksheetByName(MyWorkBook, AlignmentSheet);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet');
      for col :=0 to MyWorksheet.GetLastColIndex do
        for row := 0 to MyWorksheet.GetlastRowIndex do
        begin
          MyCell := MyWorksheet.FindCell(row, col);
          if MyCell = nil then
            fail('Error in test code. Failed to get cell.');
          vertAlign := TsVertAlignment(row);
          if (vertAlign = vaBottom) and (AFormat in [sfExcel5, sfExcel8]) then
            vertAlign := vaDefault;
          CheckEquals(
            GetEnumName(TypeInfo(TsVertAlignment), Integer(vertAlign)),
            GetEnumName(TypeInfo(TsVertAlignment), Integer(MyCell^.VertAlignment)),
            'Test saved vertical alignment mismatch, cell ' + CellNotation(MyWorksheet,row,col)
          );
          horAlign := TsHorAlignment(col);
          CheckEquals(
            GetEnumName(TypeInfo(TsHorAlignment), Integer(horAlign)),
            GetEnumName(TypeInfo(TsHorAlignment), Integer(MyCell^.HorAlignment)),
            'Test saved horizontal alignment mismatch, cell ' + CellNotation(MyWorksheet,row,col)
          );
        end;
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_Alignment;
begin
  TestWriteRead_Alignment(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_Alignment;
begin
  TestWriteRead_Alignment(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_Alignment;
begin
  TestWriteRead_Alignment(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_Alignment;
begin
  TestWriteRead_Alignment(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_Alignment;
begin
  TestWriteRead_Alignment(sfOOXML);
end;


{ This test writes in column A the names of the Background.Styles, in column B
  the background fill with a specific pattern and background color, in column C
  the same, but with transparent background. }
procedure TSpreadWriteReadFormatTests.TestWriteRead_Background(AFormat: TsSpreadsheetFormat);
const
  PATTERN_COLOR = scRed;
  BK_COLOR = scYellow;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  col, row: Integer;
  style: TsFillStyle;
  TempFile: String;
  actualstyle: TsFillStyle;
  actualcolor: TsColor;
  patt: TsFillPattern;
begin
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(BackgroundSheet);
    for style in TsFillStyle do begin
      row := ord(style);
      MyWorksheet.WriteUTF8Text(row, 0, GetEnumName(TypeInfo(TsFillStyle), ord(style)));
      MyWorksheet.WriteBackground(row, 1, style, PATTERN_COLOR, BK_COLOR);
      MyWorksheet.WriteBackground(row, 2, style, PATTERN_COLOR, scTransparent);
    end;
    TempFile:= NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, BackgroundSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
     for style in TsFillStyle do begin
      row := ord(style);

      // Column B has BK_COLOR as backgroundcolor of the patterns
      col := 1;
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell ' + CellNotation(MyWorksheet, row, col));
      patt := MyWorksheet.ReadBackground(MyCell);
      CheckEquals(
        GetEnumName(TypeInfo(TsFillStyle), ord(style)),
        GetEnumName(TypeInfo(TsFillStyle), ord(patt.Style)),
        'Test saved fill style mismatch, cell ' + CellNotation(MyWorksheet, row, col));
      if style <> fsNoFill then
      begin
        if PATTERN_COLOR <> patt.FgColor then
          CheckEquals(
            GetColorName(PATTERN_COLOR),
            GetColorName(patt.FgColor),
            'Test saved fill pattern color mismatch, cell ' + CellNotation(MyWorksheet, row, col));
        if BK_COLOR <> patt.BgColor then
          CheckEquals(
            GetColorName(BK_COLOR),
            GetColorName(patt.BgColor),
            'Test saved fill background color mismatch, cell ' + CellNotation(MyWorksheet, row, col));
      end;

      // Column C has a transparent pattern background.
      col := 2;
      MyCell := Myworksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell ' + CellNotation(MyWorksheet, row, col));
      patt := MyWorksheet.ReadBackground(MyCell);
      CheckEquals(
        GetEnumName(TypeInfo(TsFillStyle), ord(style)),
        GetEnumName(TypeInfo(TsFillStyle), ord(patt.Style)),
        'Test saved fill style mismatch, cell ' + CellNotation(MyWorksheet, row, col));
      if style <> fsNoFill then
      begin
        if PATTERN_COLOR <> patt.FgColor then
          CheckEquals(
            GetColorName(PATTERN_COLOR),
            GetColorName(patt.FgColor),
            'Test saved fill pattern color mismatch, cell ' + CellNotation(MyWorksheet, row, col));
        // SolidFill is a special case: here the background color is always equal
        // to the pattern color - the cell layout does not know this...
        if style = fsSolidFill then
          CheckEquals(
            GetColorName(PATTERN_COLOR),
            GetColorName(patt.BgColor),
            'Test saved fill pattern color mismatch, cell ' + CellNotation(MyWorksheet, row, col))
        else
          CheckEquals(
            GetColorName(scTransparent),
            GetColorName(patt.BgColor),
            'Test saved fill background color mismatch, cell ' + CellNotation(MyWorksheet, row, col));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_Background;
begin
  TestWriteRead_Background(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_Background;
begin
  TestWriteRead_Background(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_Background;
begin
  TestWriteRead_Background(sfOOXML);
end;


{ --- Border on/off tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_Border(AFormat: TsSpreadsheetFormat);
const
  row = 0;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  col, maxCol: Integer;
  TempFile: string; //write xls/xml to this file and read back from it

  function GetBordersAsText(ABorders: TsCellBorders): String;
  var
    cb: TsCellBorder;
  begin
    Result := '';
    for cb in ABorders do
      if Result = '' then
        Result := GetEnumName(TypeInfo(TsCellBorder), ord(cb))
      else
        Result := Result + ', ' + GetEnumName(TypeInfo(TsCellBorder), ord(cb));
    if Result = '' then
      Result := 'no borders';
  end;

begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(BordersSheet);
    if AFormat in [sfExcel2, sfExcel5] then
      maxCol := 15   // no diagonal border support in BIFF2 and BIFF5
    else
      maxCol := High(SollBorders);
    for col := Low(SollBorders) to maxCol do
    begin
      // It is important for the test to write contents to the cell. Without it
      // the first cell (col=0) would not even contain a format and would be
      // dropped by the ods reader resulting in a matching error.
      MyCell := MyWorksheet.WriteUTF8Text(row, col, GetBordersAsText(SollBorders[col]));
      MyWorksheet.WriteBorders(MyCell, SollBorders[col]);
    end;

    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, BordersSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for col := 0 to MyWorksheet.GetLastColIndex do
    begin
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell');
      CheckEquals(
        GetBordersAsText(SollBorders[col]),
        GetBordersAsText(MyCell^.Border),
        'Test saved border mismatch, cell ' + CellNotation(MyWorksheet, row, col)
      );
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_Border;
begin
  TestWriteRead_Border(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_Border;
begin
  TestWriteRead_Border(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_Border;
begin
  TestWriteRead_Border(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_Border;
begin
  TestWriteRead_Border(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_Border;
begin
  TestWriteRead_Border(sfOOXML);
end;


{ --- BorderStyle tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_BorderStyles(AFormat: TsSpreadsheetFormat);
{ This test paints 10x10 cells with all borders, each separated by an empty
  column and an empty row. The border style varies from border to border
  according to the line styles defined in SollBorderStyles. At first, all border
  lines use the first color in SollBorderColors. When all BorderStyles are used
  the next color is taken, etc. }
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  row, col: Integer;
  b: TsCellBorder;
  expected: Integer;
  current: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
  c, ls: Integer;
  borders: TsCellBorders;
  borderstyle: TsCellBorderStyle;
  diagUp_ls: Integer;
  diagUp_clr: integer;
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(BordersSheet);

    borders := [cbNorth, cbSouth, cbEast, cbWest];
    if AFormat in [sfExcel8, sfOpenDocument, sfOOXML] then
      borders := borders + [cbDiagUp, cbDiagDown];

    c := 0;
    ls := 0;
    for row := 1 to 10 do
    begin
      for col := 1 to 10 do
      begin
        MyWorksheet.WriteBorders(row*2-1, col*2-1, borders);
        for b in borders do
        begin
          MyWorksheet.WriteBorderLineStyle(row*2-1, col*2-1, b, SollBorderLineStyles[ls]);
          MyWorksheet.WriteBorderColor(row*2-1, col*2-1, b, SollBorderColors[c]);
          inc(ls);
          if ls > High(SollBorderLineStyles) then
          begin
            ls := 0;
            inc(c);
            if c > High(SollBorderColors) then
              c := 0;
          end;
        end;
      end;
    end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, BordersSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    c := 0;
    ls := 0;
    for row := 1 to 10 do
    begin
      for col := 1 to 10 do
      begin
        MyCell := MyWorksheet.FindCell(row*2-1, col*2-1);
        if myCell = nil then
          fail('Error in test code. Failed to get cell.');
        for b in borders do
        begin
          borderStyle := MyWorksheet.ReadCellBorderStyle(MyCell, b);
          current := ord(borderStyle.LineStyle);
          // In Excel both diagonals have the same line style. The reader picks
          // the line style of the "diagonal-up" border. We use this as expected
          // value in the "diagonal-down" case.
          expected := ord(SollBorderLineStyles[ls]);
          if AFormat in [sfExcel8, sfOOXML] then
            case b of
              cbDiagUp   : diagUp_ls := expected;
              cbDiagDown : expected := diagUp_ls;
            end;
          CheckEquals(expected, current,
            'Test saved border line style mismatch, cell ' + CellNotation(MyWorksheet, row*2, col*2));
          current := borderStyle.Color;
          expected := SollBorderColors[c];
          // In Excel both diagonals have the same line color. The reader picks
          // the color of the "diagonal-up" border. We use this as expected value
          // in the "diagonal-down" case.
          if AFormat in [sfExcel8, sfOOXML] then
            case b of
              cbDiagUp   : diagUp_clr := expected;
              cbDiagDown : expected := diagUp_clr;
            end;
          CheckEquals(expected, current,
            'Test saved border color mismatch, cell ' + CellNotation(MyWorksheet, row*2, col*2));
          inc(ls);
          if ls > High(SollBorderLineStyles) then begin
            ls := 0;
            inc(c);
            if c > High(SollBorderColors) then
              c := 0;
          end;
        end;
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_BorderStyles;
begin
  TestWriteRead_BorderStyles(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_BorderStyles;
begin
  TestWriteRead_BorderStyles(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_BorderStyles;
begin
  TestWriteRead_BorderStyles(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_BorderStyles;
begin
  TestWriteRead_BorderStyles(sfOOXML);
end;


{ --- Column widths tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_ColWidths(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualColWidth: Single;
  Col: Integer;
  lpCol: PCol;
  lCol: TCol;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(ColWidthSheet);
    for Col := Low(SollColWidths) to High(SollColWidths) do
    begin
      lCol.Width := SollColWidths[Col];
      //MyWorksheet.WriteNumber(0, Col, 1);
      MyWorksheet.WriteColInfo(Col, lCol);
    end;
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, ColWidthSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for Col := Low(SollColWidths) to High(SollColWidths) do
    begin
      lpCol := MyWorksheet.GetCol(Col);
      if lpCol = nil then
        fail('Error in test code. Failed to return saved column width');
      ActualColWidth := lpCol^.Width;
      if abs(SollColWidths[Col] - ActualColWidth) > 1E-2 then   // take rounding errors into account
        CheckEquals(SollColWidths[Col], ActualColWidth,
          'Test saved colwidth mismatch, column '+ColNotation(MyWorkSheet,Col));
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_ColWidths;
begin
  TestWriteRead_ColWidths(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_ColWidths;
begin
  TestWriteRead_ColWidths(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_ColWidths;
begin
  TestWriteRead_ColWidths(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_ColWidths;
begin
  TestWriteRead_ColWidths(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_ColWidths;
begin
  TestWriteRead_ColWidths(sfOOXML);
end;


{ --- Row height tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_RowHeights(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualRowHeight: Single;
  Row: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(RowHeightSheet);
    for Row := Low(SollRowHeights) to High(SollRowHeights) do
      MyWorksheet.WriteRowHeight(Row, SollRowHeights[Row]);
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, RowHeightSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for Row := Low(SollRowHeights) to High(SollRowHeights) do
    begin
      ActualRowHeight := MyWorksheet.GetRowHeight(Row);
      // Take care of rounding errors - due to missing details of calculation
      // they can be quite large...
      if abs(ActualRowHeight - SollRowHeights[Row]) > 1e-2 then
        CheckEquals(SollRowHeights[Row], ActualRowHeight,
          'Test saved row height mismatch, row '+RowNotation(MyWorkSheet,Row));
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_RowHeights;
begin
  TestWriteRead_RowHeights(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_RowHeights;
begin
  TestWriteRead_RowHeights(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_RowHeights;
begin
  TestWriteRead_RowHeights(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_RowHeights;
begin
  TestWriteRead_RowHeights(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_RowHeights;
begin
  TestWriteRead_RowHeights(sfOOXML);
end;


{ --- Text rotation tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_TextRotation(AFormat: TsSpreadsheetFormat);
const
  col = 0;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  tr: TsTextRotation;
  row: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(TextRotationSheet);
    for tr := Low(TsTextRotation) to High(TsTextRotation) do
    begin
      row := ord(tr);
      MyWorksheet.WriteTextRotation(row, col, tr);
      MyCell := MyWorksheet.GetCell(row, col);
      CheckEquals(
        GetEnumName(TypeInfo(TsTextRotation), ord(tr)),
        GetEnumName(TypeInfo(TsTextRotation), ord(MyCell^.TextRotation)),
        'Test unsaved textrotation mismatch, cell ' + CellNotation(MyWorksheet, row, col));
    end;
    TempFile:=NewTempFile;
  finally
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, TextRotationSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for row := 0 to MyWorksheet.GetLastRowIndex do
    begin
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell');
      tr := MyCell^.TextRotation;
      CheckEquals(
        GetEnumName(TypeInfo(TsTextRotation), ord(TsTextRotation(row))),
        GetEnumName(TypeInfo(TsTextRotation), ord(MyCell^.TextRotation)),
        'Test saved textrotation mismatch, cell ' + CellNotation(MyWorksheet, row, col));
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_TextRotation;
begin
  TestWriteRead_TextRotation(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_TextRotation;
begin
  TestWriteRead_TextRotation(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_TextRotation;
begin
  TestWriteRead_TextRotation(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_TextRotation;
begin
  TestWriteRead_TextRotation(sfOOXML);
end;


{ --- Wordwrap tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_WordWrap(AFormat: TsSpreadsheetFormat);
const
  LONGTEXT = 'This is a very, very, very, very long text.';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values:
  // Cell A1 is word-wrapped, Cell B1 is NOT word-wrapped
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(WordwrapSheet);
    MyWorksheet.WriteUTF8Text(0, 0, LONGTEXT);
    MyWorksheet.WriteUsedFormatting(0, 0, [uffWordwrap]);
    MyCell := MyWorksheet.FindCell(0, 0);
    if MyCell = nil then
      fail('Error in test code. Failed to get word-wrapped cell.');
    CheckEquals(true, MyWorksheet.ReadWordwrap(MyCell),
      'Test unsaved word wrap mismatch cell ' + CellNotation(MyWorksheet,0,0));
    MyWorksheet.WriteUTF8Text(1, 0, LONGTEXT);
    MyWorksheet.WriteUsedFormatting(1, 0, []);
    MyCell := MyWorksheet.FindCell(1, 0);
    if MyCell = nil then
      fail('Error in test code. Failed to get word-wrapped cell.');
    CheckEquals(false, MyWorksheet.ReadWordwrap(MyCell),
      'Test unsaved non-wrapped cell mismatch, cell ' + CellNotation(MyWorksheet,0,0));
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet, as biff8
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, WordwrapSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    MyCell := MyWorksheet.FindCell(0, 0);
    if MyCell = nil then
      fail('Error in test code. Failed to get word-wrapped cell.');
    CheckEquals(true, MyWorksheet.ReadWordwrap(MyCell),
      'Failed to return correct word-wrap flag, cell ' + CellNotation(MyWorksheet,0,0));
    MyCell := MyWorksheet.FindCell(1, 0);
    if MyCell = nil then
      fail('Error in test code. Failed to get non-wrapped cell.');
    CheckEquals(false, MyWorksheet.ReadWordwrap(MyCell),
      'Failed to return correct word-wrap flag, cell ' + CellNotation(MyWorksheet,0,0));
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_Wordwrap;
begin
  TestWriteRead_Wordwrap(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_Wordwrap;
begin
  TestWriteRead_Wordwrap(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_Wordwrap;
begin
  TestWriteRead_Wordwrap(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_Wordwrap;
begin
  TestWriteRead_Wordwrap(sfOOXML);
end;


{ --- Merged tests --- }

procedure TSpreadWriteReadFormatTests.TestWriteRead_MergedCells(AFormat: TsSpreadsheetFormat);
const
  TEST_RANGES: Array[0..3] of string = ('A1:B1', 'E1:G5', 'H1:H5', 'L2:M4');
  SHEETNAME1 = 'Sheet1';
  SHEETNAME2 = 'Sheet2';
  SHEETNAME3 = 'Sheet3';
  CELL_TEXT = 'Lazarus';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  cell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  r1, c1, r2, c2: Cardinal;
  r, c: Cardinal;
  actual, expected: String;
  i: Integer;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    // 1st sheet: merged ranges with text
    MyWorksheet:= MyWorkBook.AddWorksheet(SHEETNAME1);
    for i:=0 to High(TEST_RANGES) do
    begin
      ParseCellRangeString(TEST_RANGES[i], r1, c1, r2, c2);
      Myworksheet.WriteUTF8Text(r1, c1, CELL_TEXT);
      Myworksheet.MergeCells(r1, c1, r2, c2);
    end;

    // 2nd sheet: merged ranges, empty
    Myworksheet := MyWorkbook.AddWorksheet(SHEETNAME2);
    for i:=0 to High(TEST_RANGES) do
    begin
      ParseCellRangeString(TEST_RANGES[i], r1, c1, r2, c2);
      Myworksheet.MergeCells(r1, c1, r2, c2);
    end;

    // 3rd sheet: merged ranges, with text, then unmerge all
    MyWorksheet:= MyWorkBook.AddWorksheet(SHEETNAME3);
    for i:=0 to High(TEST_RANGES) do
    begin
      ParseCellRangeString(TEST_RANGES[i], r1, c1, r2, c2);
      Myworksheet.WriteUTF8Text(r1, c1, CELL_TEXT);
      Myworksheet.MergeCells(r1, c1, r2, c2);
      Myworksheet.UnmergeCells(r1, c1);
    end;

    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);

    // 1st sheet: merged cells with text
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEETNAME1);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet ' + SHEETNAME1);
    for i:=0 to High(TEST_RANGES) do begin
      ParseCellRangeString(TEST_RANGES[i], r1, c1, r2, c2);
      cell := MyWorksheet.FindCell(r1, c1);
      if MyWorksheet.IsMergeBase(cell) then begin
        MyWorksheet.FindMergedRange(cell, r1, c1, r2, c2);
        actual := GetCellRangeString(r1, c1, r2, c2);
        expected := TEST_RANGES[i];
        if AFormat in [sfExcel2, sfExcel5] then
          CheckNotEquals(expected, actual, 'No merged cells expected, ' + expected)
        else
          CheckEquals(expected, actual, 'Merged cell range mismatch, ' + expected);
      end else
      if not (AFormat in [sfExcel2, sfExcel5]) then
        fail('Unmerged cell found, ' + CellNotation(MyWorksheet, r1, c1));
      CheckEquals(CELL_TEXT, MyWorksheet.ReadAsUTF8Text(cell),
        'Merged cell content mismatch, cell '+ CellNotation(MyWorksheet, r1, c1));
    end;

    if AFormat = sfExcel2 then
      exit;  // only 1 page in Excel2

    // 2nd sheet: merged empty cells
    MyWorksheet := GetWorksheetByName(MyWorkBook, SHEETNAME2);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet' + SHEETNAME2);
    for i:=0 to High(TEST_RANGES) do begin
      ParseCellRangeString(TEST_RANGES[i], r1, c1, r2, c2);
      cell := MyWorksheet.FindCell(r1, c1);
      if MyWorksheet.IsMergeBase(cell) then begin
        MyWorksheet.FindMergedRange(cell, r1, c1, r2, c2);
        actual := GetCellRangeString(r1, c1, r2, c2);
        expected := TEST_RANGES[i];
        if AFormat = sfExcel5 then
          CheckNotEquals(expected, actual, 'Merged cells found in Excel5, ' + expected)
        else
          CheckEquals(expected, actual, 'Merged cell range mismatch, ' + expected);
      end else
      if AFormat <> sfExcel5 then
        fail('Unmerged cell found, ' + CellNotation(MyWorksheet, r1, c1));
      CheckEquals('', MyWorksheet.ReadAsUTF8Text(cell),
        'Merged cell content mismatch, cell '+CellNotation(MyWorksheet, r1, c1));
    end;

    // 3rd sheet: merged & unmerged cells
    MyWorksheet := GetWorksheetByName(MyWorkBook, SHEETNAME3);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet' + SHEETNAME3);
    for i:=0 to High(TEST_RANGES) do begin
      ParseCellRangeString(TEST_RANGES[i], r1, c1, r2, c2);
      cell := MyWorksheet.FindCell(r1, c1);
      if MyWorksheet.IsMergeBase(cell) then
        fail('Unmerged cell expected, cell ' + CellNotation(MyWorksheet, r1, c1));
      CheckEquals(CELL_TEXT, MyWorksheet.ReadAsUTF8Text(cell),
        'Merged/unmerged cell content mismatch, cell '+CellNotation(MyWorksheet, r1, c1));
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_MergedCells;
begin
  TestWriteRead_MergedCells(sfExcel2);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF5_MergedCells;
begin
  TestWriteRead_MergedCells(sfExcel5);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF8_MergedCells;
begin
  TestWriteRead_MergedCells(sfExcel8);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_ODS_MergedCells;
begin
  TestWriteRead_MergedCells(sfOpenDocument);
end;

procedure TSpreadWriteReadFormatTests.TestWriteRead_OOXML_MergedCells;
begin
  TestWriteRead_MergedCells(sfOOXML);
end;

{ If a biff2 file contains more than 62 XF records the XF record index is stored
  in a separats IXFE record. This is tested here. }
procedure TSpreadWriteReadFormatTests.TestWriteRead_ManyXF(AFormat: TsSpreadsheetFormat);
const
  SHEETNAME = 'Too-many-xf-records';
  FontSizes: array[0..7] of Integer = (9, 10, 12, 14, 16, 18, 20, 24);
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  cell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  r1, c1, r2, c2: Cardinal;
  r, c: Cardinal;
  fnt: TsFont;
  actual, expected: String;
  i: Integer;
  palette: TsPalette;
begin
  palette := TsPalette.Create;
  try
    palette.AddBuiltinColors;

    MyWorkbook := TsWorkbook.Create;
    try
      MyWorksheet:= MyWorkBook.AddWorksheet(SHEETNAME);
      for r := 0 to 7 do     // each row has a different font size
        for c := 0 to 7 do   // each column has a different font color
        begin
          MyWorksheet.WriteNumber(r, c, 123);
          MyWorksheet.WriteBackgroundColor(r, c, 0);
          MyWorksheet.WriteFont(r, c, 'Times New Roman', FontSizes[r], [], palette[c]);  // Biff2 has only 8 colors --> re-use the black!
          // --> in total 64 combinations
        end;
      TempFile:=NewTempFile;
      MyWorkBook.WriteToFile(TempFile, AFormat, true);
    finally
      MyWorkbook.Free;
    end;

    // Open the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.ReadFromFile(TempFile, AFormat);

      // 1st sheet: merged cells with text
      if AFormat = sfExcel2 then
        MyWorksheet := MyWorkbook.GetFirstWorksheet
      else
        MyWorksheet := GetWorksheetByName(MyWorkBook, SHEETNAME);
      if MyWorksheet=nil then
        fail('Error in test code. Failed to get named worksheet ' + SHEETNAME);

      for r:=0 to MyWorksheet.GetLastRowIndex do
        for c := 0 to MyWorksheet.GetLastColIndex do
        begin
          cell := MyWorksheet.FindCell(r, c);
          fnt := MyWorksheet.ReadCellFont(cell);
          expected := FloatToStr(FontSizes[r]);
          actual := FloatToStr(fnt.Size);
          CheckEquals(expected, actual,
            'Font size mismatch, cell '+ CellNotation(MyWorksheet, r, c));
          expected := IntToStr(palette[c]);
          actual := IntToStr(fnt.Color);
          CheckEquals(expected, actual,
            'Font color mismatch, cell '+ CellNotation(MyWorksheet, r, c));
        end;

    finally
      MyWorkbook.Free;
      DeleteFile(TempFile);
    end;

  finally
    palette.Free;
  end;
end;


procedure TSpreadWriteReadFormatTests.TestWriteRead_BIFF2_ManyXFRecords;
begin
  TestWriteRead_ManyXF(sfExcel2);
end;

initialization
  RegisterTest(TSpreadWriteReadFormatTests);
  InitSollFmtData;

end.

