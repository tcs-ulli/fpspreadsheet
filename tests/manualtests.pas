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
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of dates/times that should occur in spreadsheet
  SollColors: array[0..22] of tsColor; //"Soll" is a German word in Dutch accountancy circles meaning "normative value to check against". There ;)
  SollColorNames: array[0..22] of string; //matching names for SollColors
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollColors;

type
  { TSpreadManualTests }
  // Writes to file and let humans figure out if the correct output was generated
  TSpreadManualTests= class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Writes all background colors in A1..A16
    procedure TestBiff8CellBackgroundColor;
  end;

implementation

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
  SollColors[16]:=scGrey10pct;
  SollColors[17]:=scGrey20pct;
  SollColors[18]:=scOrange;
  SollColors[19]:=scDarkBrown;
  SollColors[20]:=scBrown;
  SollColors[21]:=scBeige;
  SollColors[22]:=scWheat;

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
  SollColorNames[16]:='scGrey10pct';
  SollColorNames[17]:='scGrey20pct';
  SollColorNames[18]:='scOrange';
  SollColorNames[19]:='scDarkBrown';
  SollColorNames[20]:='scBrown';
  SollColorNames[21]:='scBeige';
  SollColorNames[22]:='scWheat';
end;

{ TSpreadManualTests }
procedure TSpreadManualTests.SetUp;
begin
  InitSollColors;
end;

procedure TSpreadManualTests.TearDown;
begin

end;


procedure TSpreadManualTests.TestBiff8CellBackgroundColor();
// source: forum post
// http://forum.lazarus.freepascal.org/index.php/topic,19887.msg134114.html#msg134114
// possible fix for values there too
const
  OUTPUT_FORMAT = sfExcel8;
var
  Workbook: TsWorkbook;
  Worksheet: TsWorksheet;
  Cell : PCell;
  i: cardinal;
  RowOffset: cardinal;
begin
  Workbook := TsWorkbook.Create;
  Worksheet := Workbook.AddWorksheet('colorsheet');
  WorkSheet.WriteUTF8Text(0,1,'TSpreadManualTests.TestBiff8CellBackgroundColor');
  RowOffset:=1;
  for i:=Low(SollColors) to High(SollColors) do
  begin
    WorkSheet.WriteUTF8Text(i+RowOffset,0,'BACKGROUND COLOR TEST');
    Cell := Worksheet.GetCell(i+RowOffset, 0);
    Cell^.BackgroundColor := SollColors[i];
    if not (uffBackgroundColor in Cell^.UsedFormattingFields) then
      include (Cell^.UsedFormattingFields,uffBackgroundColor);
    WorkSheet.WriteUTF8Text(i+RowOffset,1,'Cell to the left should be tsColor value '+SollColorNames[i]+'. Please check.');
  end;
  // todo: move to a shared workbook object, write at tests suite finish
  // http://wiki.lazarus.freepascal.org/fpcunit#Test_decorator:_OneTimeSetup_and_OneTimeTearDown
  Workbook.WriteToFile(TestFileManual, OUTPUT_FORMAT, TRUE);
  Workbook.Free;
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadManualTests);
  // Initialize the norm variables in case other units want to use it:
  InitSollColors;
end.


