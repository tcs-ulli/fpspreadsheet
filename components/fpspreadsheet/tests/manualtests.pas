unit manualtests;

{
 Tests that can be run but need a human to check results.
 Examples are color output, rotation, bold etc
 Of course, you could write Excel macros to do this for you; patches welcome ;)
}

{$mode objfpc}{$H+}

{
Adding tests/test data:
1. Increase Soll* array size
2. Add desired normative value InitNormVariables so you can test against it
3. Add your write test(s) including instructions for the humans check the resulting file
}

interface

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, testutils, testregistry, testdecorator, fpcunit,
  fpsallformats, fpspreadsheet, fpscell,
  xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

{
var
  // Norm to test against - list of dates/times that should occur in spreadsheet
  SollColors: array[0..16] of tsColor; //"Soll" is a German word in Dutch accountancy jargon meaning "normative value to check against". There ;)
  SollColorNames: array[0..16] of string; //matching names for SollColors
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollColors;
 }

type
  { TSpreadManualSetup }
  TSpreadManualSetup= class(TTestSetup)
  protected
    procedure OneTimeSetup; override;
    procedure OneTimeTearDown; override;
  end;

  { TSpreadManualTests }
  // Writes to file and let humans figure out if the correct output was generated
  TSpreadManualTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Current fpspreadsheet does not yet have support for new RPN formulas
    {$DEFINE FPSPREAD_HAS_NEWRPNSUPPORT}
    {$IFDEF FPSPREAD_HAS_NEWRPNSUPPORT}
    // As described in bug 25718: Feature request & patch: Implementation of writing more functions
    // Writes all rpn formulas. Use Excel or Open/LibreOffice to check validity.
    procedure TestRPNFormula;
    // Dto, but writes string formulas.
//    procedure TestStringFormula;
    {$ENDIF}
    // For BIFF8 format, writes all background colors in A1..A16
    procedure TestBiff8CellBackgroundColor;

    procedure TestNumberFormats;
  end;

implementation

uses
  fpstypes, fpsUtils, fpsPalette, rpnFormulaUnit;

const
  COLORSHEETNAME='color_sheet'; //for background color tests
  RPNSHEETNAME='rpn_formula_sheet'; //for rpn formula tests
  FORMULASHEETNAME='formula_sheet';  // for string formula tests
  NUMBERFORMATSHEETNAME='number format sheet';  // for number format tests
  OUTPUT_FORMAT = sfExcel8; //change manually if you want to test different formats. To do: automatically output all formats

var
  Workbook: TsWorkbook = nil;
 (*
// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollColors;
begin
  // Set up norm - MUST match spreadsheet cells exactly
  // Follows fpspreadsheet.TsColor, except custom colors
  SollColors[0]:=scBlack;
  SollColors[1]:=scWhite;
  SollColors[2]:=scRed;
  SollColors[3]:=scGREEN;
  SollColors[4]:=scBLUE;
  SollColors[5]:=scYELLOW;
  SollColors[6]:=scMAGENTA;
  SollColors[7]:=scCYAN;
  SollColors[8]:=scDarkRed;
  SollColors[9]:=scDarkGreen;
  SollColors[10]:=scDarkBlue;
  SollColors[11]:=scOLIVE;
  SollColors[12]:=scPURPLE;
  SollColors[13]:=scTEAL;
  SollColors[14]:=scSilver;
  SollColors[15]:=scGrey;
  SollColors[16]:=scOrange;
  {
  SollColors[16]:=scGrey10pct;
  SollColors[17]:=scGrey20pct;
  SollColors[18]:=scOrange;
  SollColors[19]:=scDarkBrown;
  SollColors[20]:=scBrown;
  SollColors[21]:=scBeige;
  SollColors[22]:=scWheat;
   }
  // Corresponding names for display in cells:
  SollColorNames[0]:='scBlack';
  SollColorNames[1]:='scWhite';
  SollColorNames[2]:='scRed';
  SollColorNames[3]:='scGREEN';
  SollColorNames[4]:='scBLUE';
  SollColorNames[5]:='scYELLOW';
  SollColorNames[6]:='scMAGENTA';
  SollColorNames[7]:='scCYAN';
  SollColorNames[8]:='scDarkRed';
  SollColorNames[9]:='scDarkGreen';
  SollColorNames[10]:='scDarkBlue';
  SollColorNames[11]:='scOLIVE';
  SollColorNames[12]:='scPURPLE';
  SollColorNames[13]:='scTEAL';
  SollColorNames[14]:='scSilver';
  SollColorNames[15]:='scGrey';
  SollColorNames[16]:='scOrange';
  {
  SollColorNames[16]:='scGrey10pct';
  SollColorNames[17]:='scGrey20pct';
  SollColorNames[18]:='scOrange';
  SollColorNames[19]:='scDarkBrown';
  SollColorNames[20]:='scBrown';
  SollColorNames[21]:='scBeige';
  SollColorNames[22]:='scWheat';
  }
end;
   *)

{ TSpreadManualSetup }

procedure TSpreadManualSetup.OneTimeSetup;
begin
  // One time setup for entire suite: nothing needed here yet
end;

procedure TSpreadManualSetup.OneTimeTearDown;
begin
  if Workbook <> nil then begin
    // In Ubuntu explicit deletion of the existing file is needed.
    // Otherwise an error would occur and a defective file would be written }
    if FileExists(TestFileManual) then DeleteFile(TestFileManual);

    Workbook.WriteToFile(TestFileManual, OUTPUT_FORMAT, TRUE);
    Workbook.Free;
    Workbook := nil;
  end;
end;

{ TSpreadManualTests }
procedure TSpreadManualTests.SetUp;
begin
//  InitSollColors;
end;

procedure TSpreadManualTests.TearDown;
begin
  // nothing to do here, yet
end;

procedure TSpreadManualTests.TestBiff8CellBackgroundColor();
// source: forum post
// http://forum.lazarus.freepascal.org/index.php/topic,19887.msg134114.html#msg134114
// possible fix for values there too
var
  Worksheet: TsWorksheet;
  Cell : PCell;
  i: cardinal;
  RowOffset: cardinal;
  palette: TsPalette;
begin
  if OUTPUT_FORMAT <> sfExcel8 then
    Ignore('This test only applies to BIFF8 XLS output format.');

  // No worksheets in BIFF2. Since main interest is here in formulas we just jump
  // off here - need to change this in the future...
  if OUTPUT_FORMAT = sfExcel2 then
    Ignore('BIFF2 does not support worksheets. Ignoring manual tests for now');

  if Workbook = nil then
    Workbook := TsWorkbook.Create;

  palette := TsPalette.Create;
  try
    palette.AddBuiltinColors;
    palette.AddExcelColors;

    Worksheet := Workbook.AddWorksheet(COLORSHEETNAME);
    WorkSheet.WriteUTF8Text(0, 1, 'TSpreadManualTests.TestBiff8CellBackgroundColor');
    RowOffset := 1;
    for i:=0 to palette.Count-1 do begin
      cell := WorkSheet.WriteUTF8Text(i+RowOffset,0,'BACKGROUND COLOR TEST');
      Worksheet.WriteBackgroundColor(Cell, palette[i]);
      Worksheet.WriteFontColor(cell, HighContrastColor(palette[i]));
      WorkSheet.WriteUTF8Text(i+RowOffset,1,'Cell to the left should be '+GetColorName(palette[i])+'. Please check.');
    end;
    Worksheet.WriteColWidth(0, 30);
    Worksheet.WriteColWidth(1, 60);
  finally
    palette.Free;
  end;
end;

procedure TSpreadManualTests.TestNumberFormats();
// source: forum post
// http://forum.lazarus.freepascal.org/index.php/topic,19887.msg134114.html#msg134114
// possible fix for values there too
const
  Values: Array[0..4] of Double = (12000.34, -12000.34, 0.0001234, -0.0001234, 0.0);
  FormatStrings: array[0..24] of String = (
    'General',
    '0',     '0.00',     '0.0000',
    '#,##0', '#,##0.00', '#,##0.0000',
    '0%',    '0.00%',    '0.0000%',
    '0,',    '0.00,',    '0.0000,',
    '0E+00', '0.00E+00', '0.0000E+00',
    '0E-00', '0.00E-00', '0.0000E-00',
    '# ?/?', '# ??/??',  '# ????/????',
    '?/?',   '??/??',    '????/????'
  );
var
  Worksheet: TsWorksheet;
  Cell : PCell;
  i: cardinal;
  r, c: Cardinal;
  palette: TsPalette;
  nfs: String;
begin
  if OUTPUT_FORMAT <> sfExcel8 then
    Ignore('This test only applies to BIFF8 XLS output format.');

  // No worksheets in BIFF2. Since main interest is here in formulas we just jump
  // off here - need to change this in the future...
  if OUTPUT_FORMAT = sfExcel2 then
    Ignore('BIFF2 does not support worksheets. Ignoring manual tests for now');

  if Workbook = nil then
    Workbook := TsWorkbook.Create;

  Worksheet := Workbook.AddWorksheet(NUMBERFORMATSHEETNAME);
  WorkSheet.WriteUTF8Text(0, 1, 'Number format tests');

  for r:=0 to High(FormatStrings) do
  begin
    Worksheet.WriteUTF8Text(r+2, 0, FormatStrings[r]);
    for c:=0 to High(Values) do
      Worksheet.WriteNumber(r+2, c+1, values[c], nfCustom, FormatStrings[r]);
  end;

  Worksheet.WriteColWidth(0, 20);
end;

{$IFDEF FPSPREAD_HAS_NEWRPNSUPPORT}
// As described in bug 25718: Feature request & patch: Implementation of writing more functions
procedure TSpreadManualTests.TestRPNFormula;
var
  Worksheet: TsWorksheet;
begin
  if Workbook = nil then
    Workbook := TsWorkbook.Create;

  Worksheet := Workbook.AddWorksheet(RPNSHEETNAME);
  WriteRPNFormulaSamples(Worksheet, OUTPUT_FORMAT, false);
end;
                                                               (*
procedure TSpreadManualTests.TestStringFormula;
var
  Worksheet: TsWorksheet;
begin
  if Workbook = nil then
    Workbook := TsWorkbook.Create;

  Worksheet := Workbook.AddWorksheet(FORMULASHEETNAME);
  WriteRPNFormulaSamples(Worksheet, OUTPUT_FORMAT, false, false);
end;
*)
{$ENDIF}

initialization
  // Register one time setup/teardown and associated test class to actually run the tests
  RegisterTestDecorator(TSpreadManualSetup,TSpreadManualTests);
  // Initialize the norm variables in case other units want to use it:
//  InitSollColors;

end.


