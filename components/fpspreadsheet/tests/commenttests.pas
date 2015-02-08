unit commenttests;

{$mode objfpc}{$H+}

interface
{ Color tests
This unit tests writing out to and reading back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadCommemtTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadCommentTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_Comment(AFormat: TsSpreadsheetFormat;
      const ACommentText: String);
  published
    // Writes out comments & reads back.

    { BIFF2 comment tests }
    procedure TestWriteRead_BIFF2_Standard_Comment;
    procedure TestWriteRead_BIFF2_NonAscii_Comment;
    procedure TestWriteRead_BIFF2_NonXMLChar_Comment;
    procedure TestWriteRead_BIFF2_VeryLong_Standard_Comment;
    procedure TestWriteRead_BIFF2_VeryLong_NonAscii_Comment;

    { BIFF5 comment tests }
    procedure TestWriteRead_BIFF5_Standard_Comment;
    procedure TestWriteRead_BIFF5_NonAscii_Comment;
    procedure TestWriteRead_BIFF5_NonXMLChar_Comment;
    procedure TestWriteRead_BIFF5_VeryLong_Standard_Comment;
    procedure TestWriteRead_BIFF5_VeryLong_NonAscii_Comment;

    { BIFF8 comment tests }
    // writing is currently not supported
    //procedure TestWriteRead_BIFF8_Standard_Comment;
    //procedure TestWriteRead_BIFF8_NonAscii_Comment;
    //procedure TestWriteRead_BIFF8_NonXMLChar_Comment;

    { OpenDocument comment tests }
    procedure TestWriteRead_ODS_Standard_Comment;
    procedure TestWriteRead_ODS_NonAscii_Comment;
    procedure TestWriteRead_ODS_NonXMLChar_Comment;
    procedure TestWriteRead_ODS_VeryLong_Comment;

    { OOXML comment tests }
    procedure TestWriteRead_OOXML_Standard_Comment;
    procedure TestWriteRead_OOXML_NonAscii_Comment;
    procedure TestWriteRead_OOXML_NonXMLChar_Comment;
    procedure TestWriteRead_OOXML_VeryLong_Comment;
  end;

implementation

const
  CommentSheet = 'Comments';

  STANDARD_COMMENT = 'This is a comment';
  COMMENT_UTF8 = 'Comment with non-standard characters: ÄÖÜß café au lait'; // водка wódka';
  COMMENT_XML = 'Comment with characters not allowed by XML: <, >';

var
  VERY_LONG_COMMENT: String;
  VERY_LONG_NONASCII_COMMENT: String;

{ TSpreadWriteReadCommentTests }

procedure TSpreadWriteReadCommentTests.SetUp;
var
  i: Integer;
begin
  inherited SetUp;

  // In BIFF2-5, comments longer than 2048 characters are split into several
  // NOTE records.
  VERY_LONG_COMMENT := '';
  repeat
    VERY_LONG_COMMENT := VERY_LONG_COMMENT + '1234567890 ';
  until Length(VERY_LONG_COMMENT) > 3000;

  VERY_LONG_NONASCII_COMMENT := '';
  repeat
    VERY_LONG_NONASCII_COMMENT := VERY_LONG_NONASCII_COMMENT + 'ÄÖÜäöü ';
  until Length(VERY_LONG_NONASCII_COMMENT) > 3000;
end;

procedure TSpreadWriteReadCommentTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_Comment(
  AFormat: TsSpreadsheetFormat; const ACommentText: String);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col, lastCol: Integer;
  expected, actual: String;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(CommentSheet);

    // Comment in empty cell
    row := 0;
    col := 0;
    Myworksheet.WriteComment(row, col, ACommentText);

    // Comment in label cell
    col := 1;
    MyWorksheet.WriteUTF8Text(row, col, 'Cell with comment');
    Myworksheet.WriteComment(row, col, ACommentText);

    // Comment in number cell
    col := 2;
    MyWorksheet.WriteNumber(row, col, 123.456);
    Myworksheet.WriteComment(row, col, ACommentText);

    // Comment in formula cell
    col := 3;
    Myworksheet.WriteFormula(row, col, '1+1');
    Myworksheet.WriteComment(row, col, ACommentText);

    // Comment in boolean cell
    col := 4;
    MyWorksheet.WriteBoolValue(row, col, true);
    Myworksheet.WriteComment(row, col, ACommentText);

    // Comment in error cell
    // Error cell must be the last cell because ODS does not support error cell
    // and the test is to be omitted.
    col := 5;
    Myworksheet.WriteErrorValue(row, col, errWrongType);
    Myworksheet.WriteComment(row, col, ACommentText);

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
      MyWorksheet := GetWorksheetByName(MyWorkBook, CommentSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    row := 0;
    lastCol := MyWorksheet.GetLastColIndex;
    if AFormat = sfOpenDocument then dec(lastCol);  // No error cells supported in ODS --> skip the last test which is for error cells
    for col := 0 to lastCol do
    begin
      MyCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Failure to find cell ' + CellNotation(MyWorksheet, row, col));
      actual := MyWorksheet.ReadComment(MyCell);
      expected := ACommentText;
      CheckEquals(expected, actual,
        'Test saved comment mismatch, cell '+CellNotation(MyWorksheet, row, col));
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ Tests for BIFF2 file format }
procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF2_Standard_Comment;
begin
  TestWriteRead_Comment(sfExcel2, STANDARD_COMMENT);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF2_NonAscii_Comment;
begin
  TestWriteRead_Comment(sfExcel2, COMMENT_UTF8);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF2_NonXMLChar_Comment;
begin
  TestWriteRead_Comment(sfExcel2, COMMENT_XML);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF2_VeryLong_Standard_Comment;
begin
  TestWriteRead_Comment(sfExcel2, VERY_LONG_COMMENT);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF2_VeryLong_NonAscii_Comment;
begin
  TestWriteRead_Comment(sfExcel2, VERY_LONG_NONASCII_COMMENT);
end;

{ Tests for BIFF5 file format }
procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF5_Standard_Comment;
begin
  TestWriteRead_Comment(sfExcel5, STANDARD_COMMENT);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF5_NonAscii_Comment;
begin
  TestWriteRead_Comment(sfExcel5, COMMENT_UTF8);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF5_NonXMLChar_Comment;
begin
  TestWriteRead_Comment(sfExcel5, COMMENT_XML);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF5_VeryLong_Standard_Comment;
begin
  TestWriteRead_Comment(sfExcel5, VERY_LONG_COMMENT);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_BIFF5_VeryLong_NonAscii_Comment;
begin
  TestWriteRead_Comment(sfExcel5, VERY_LONG_NONASCII_COMMENT);
end;


{ Tests for BIFF8 file format }
{  Writing is currently not support --> the test does not make sense! }

{ Tests for Open Document file format }
procedure TSpreadWriteReadCommentTests.TestWriteRead_ODS_Standard_Comment;
begin
  TestWriteRead_Comment(sfOpenDocument, STANDARD_COMMENT);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_ODS_NonAscii_Comment;
begin
  TestWriteRead_Comment(sfOpenDocument, COMMENT_UTF8);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_ODS_NonXMLChar_Comment;
begin
  TestWriteRead_Comment(sfOpenDocument, COMMENT_XML);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_ODS_VeryLong_Comment;
begin
  TestWriteRead_Comment(sfOpenDocument, VERY_LONG_COMMENT);
end;


{ Tests for OOXML file format }
procedure TSpreadWriteReadCommentTests.TestWriteRead_OOXML_Standard_Comment;
begin
  TestWriteRead_Comment(sfOOXML, STANDARD_COMMENT);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_OOXML_NonAscii_Comment;
begin
  TestWriteRead_Comment(sfOOXML, COMMENT_UTF8);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_OOXML_NonXMLChar_Comment;
begin
  TestWriteRead_Comment(sfOOXML, COMMENT_XML);
end;

procedure TSpreadWriteReadCommentTests.TestWriteRead_OOXML_VeryLong_Comment;
begin
  TestWriteRead_Comment(sfOOXML, VERY_LONG_COMMENT);
end;

initialization
  RegisterTest(TSpreadWriteReadCommentTests);

end.

