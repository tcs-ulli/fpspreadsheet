{
xlsbiff5.pas

Writes an Excel 5 file

An Excel worksheet stream consists of a number of subsequent records.
To ensure a properly formed file, the following order must be respected:

1st record:        BOF
2nd to Nth record: Any record
Last record:       EOF

Excel 5 files are OLE compound document files, and must be written using the
fpOLE library.

Records Needed to Make a BIFF5 File Microsoft Excel Can Use:

Required Records:

BOF - Set the 6 byte offset to 0x0005 (workbook globals)
Window1
FONT - At least five of these records must be included
XF - At least 15 Style XF records and 1 Cell XF record must be included
STYLE
BOUNDSHEET - Include one BOUNDSHEET record per worksheet
EOF

BOF - Set the 6 byte offset to 0x0010 (worksheet)
INDEX
DIMENSIONS
WINDOW2
EOF

Excel file format specification obtained from:

http://sc.openoffice.org/excelfileformat.pdf

Records Needed to Make a BIFF5 File Microsoft Excel Can Use obtained from:

http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q147732&ID=KB;EN-US;Q147732&LN=EN-US&rnk=2&SD=msdn&FR=0&qry=BIFF&src=DHCS_MSPSS_msdn_SRCH&SPR=MSALL&

AUTHORS: Felipe Monteiro de Carvalho
}
unit xlsbiff8;

{$ifdef fpc}
{$mode delphi}{$H+}
{$endif}

interface

uses
  Classes, SysUtils,
  fpspreadsheet;

type

  { TsSpreadBIFF5Writer }

  TsSpreadBIFF5Writer = class(TsCustomSpreadWriter)
  public
    { General writing methods }
    procedure WriteToStream(AStream: TStream; AData: TsWorkbook); override;
    { Record writing methods }
    procedure WriteBOF(AStream: TStream);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFontName: Widestring = 'Arial');
    procedure WriteFormat(AStream: TStream; AIndex: Word = 0; AFormatString: Widestring = 'General');
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Word; const AFormula: TRPNFormula); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Word; const AValue: string); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal; const AValue: double); override;
    procedure WriteXF(AStream: TStream);
  end;

implementation

const
  { Excel record IDs }
  INT_EXCEL_ID_BOF        = $0809;
  INT_EXCEL_ID_EOF        = $000A;
  INT_EXCEL_ID_FONT       = $0031;
  INT_EXCEL_ID_FORMAT     = $041E;
  INT_EXCEL_ID_FORMULA    = $0006;
  INT_EXCEL_ID_LABEL      = $0004;
  INT_EXCEL_ID_NUMBER     = $0203;
  INT_EXCEL_ID_XF         = $00E0;

  { Cell Addresses constants }
  MASK_EXCEL_ROW          = $3FFF;
  MASK_EXCEL_RELATIVE_ROW = $4000;
  MASK_EXCEL_RELATIVE_COL = $8000;

  { Unicode string constants }
  INT_EXCEL_UNCOMPRESSED_STRING = $01;
  
  { BOF record constants }
  INT_EXCEL_BIFF8_VER     = $0600;
  INT_EXCEL_WORKBOOK      = $0005;
  INT_EXCEL_SHEET         = $0010;
  INT_EXCEL_CHART         = $0020;
  INT_EXCEL_MACRO_SHEET   = $0040;
  INT_EXCEL_BUILD_ID      = $1FD2;
  INT_EXCEL_BUILD_YEAR    = $07CD;
  INT_EXCEL_FILE_HISTORY  = $0000C0C1;
  INT_EXCEL_LOWEST_VER    = $00000306;

  { FONT record constants}
  INT_EXCEL_FONTWEIGHT_NORMAL = $0190;

  { XF record constants }
  INT_EXCEL_XF_TYPE_PROT_STYLEXF = $FFF4;

{
  Excel files are all written with Little Endian number,
  so it's necessary to swap the numbers to be able to build a
  correct file on big endian systems.
  
  Endianess helper functions
}

function WordToLE(AValue: Word): Word;
begin
  {$IFDEF BIG_ENDIAN}
    Result := ((AValue shl 8) and $FF00) or ((AValue shr 8) and $00FF);
  {$ELSE}
    Result := AValue;
  {$ENDIF}
end;

{
  Exported functions
}

{ TsSpreadBIFF5Writer }

procedure TsSpreadBIFF5Writer.WriteToStream(AStream: TStream; AData: TsWorkbook);
begin

end;

{*******************************************************************
*  TsSpreadBIFF5Writer.WriteBOF ()
*
*  DESCRIPTION:    Writes an Excel 5 BOF record
*
*                  This must be the first record on an Excel 5 stream
*
*******************************************************************}
procedure TsSpreadBIFF5Writer.WriteBOF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_BOF));
  AStream.WriteWord(WordToLE(16));

  { BIFF version }
  AStream.WriteWord(WordToLE(INT_EXCEL_BIFF8_VER));

  { Data type }
  AStream.WriteWord(WordToLE(INT_EXCEL_WORKBOOK));

  { Build identifier, must not be 0 }
  AStream.WriteWord(WordToLE(INT_EXCEL_BUILD_ID));

  { Build year, must not be 0 }
  AStream.WriteWord(WordToLE(INT_EXCEL_BUILD_YEAR));

  { File history flags }
//  AStream.WriteDWord($00000000);
  AStream.WriteWord(WordToLE(INT_EXCEL_FILE_HISTORY));

  { Lowest Excel version that can read all records of this file }
//  AStream.WriteDWord($00000000);
  AStream.WriteWord(WordToLE(INT_EXCEL_LOWEST_VER));
end;

{*******************************************************************
*  TsSpreadBIFF5Writer.WriteEOF ()
*
*  DESCRIPTION:    Writes an Excel 5 EOF record
*
*                  This must be the last record on an Excel 5 stream
*
*******************************************************************}
procedure TsSpreadBIFF5Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_EOF));
  AStream.WriteWord($0000);
end;

procedure TsSpreadBIFF5Writer.WriteFont(AStream: TStream;
  AFontName: Widestring);
var
  Len: Byte;
begin
  Len := Length(AFontName);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FONT));
  AStream.WriteWord(WordToLE(14 + 2 + Len*2));

  { Height of the font in twips = 1/20 of a point }
  AStream.WriteWord(WordToLE(200));

  { Option flags }
  AStream.WriteWord(0);

  { Colour index }
  AStream.WriteWord(0);

  { Font weight }
  AStream.WriteWord(WordToLE(INT_EXCEL_FONTWEIGHT_NORMAL));

  { Underline type }
  AStream.WriteByte(0);

  { Font family }
  AStream.WriteByte(0);

  { Character set }
  AStream.WriteByte(0);

  { Not used }
  AStream.WriteByte(0);

  { Font name: Unicode string, 8-bit length }
  AStream.WriteByte(Len);
  AStream.WriteByte(INT_EXCEL_UNCOMPRESSED_STRING);
  AStream.WriteBuffer(AFontName[1], Len*2);
end;

procedure TsSpreadBIFF5Writer.WriteFormat(AStream: TStream; AIndex: Word;
  AFormatString: Widestring);
var
  Len: Integer;
begin
  Len := Length(AFormatString);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMAT));
  AStream.WriteWord(WordToLE(2 + 3 + Len*2));

  { Format index used by other records }
  AStream.WriteWord(WordToLE(AIndex));

  { Unicode string, 16-bit length }
  AStream.WriteWord(WordToLE(Len));
  AStream.WriteByte(INT_EXCEL_UNCOMPRESSED_STRING);
  AStream.WriteBuffer(AFormatString[1], Len*2);
end;

{*******************************************************************
*  TsSpreadBIFF5Writer.WriteFormula ()
*
*  DESCRIPTION:    Writes an Excel 5 FORMULA record
*
*                  To input a formula to this method, first convert it
*                  to RPN, and then list all it's members in the
*                  AFormula array
*
*******************************************************************}
procedure TsSpreadBIFF5Writer.WriteFormula(AStream: TStream; const ARow,
  ACol: Word; const AFormula: TRPNFormula);
var
  FormulaResult: double;
  i: Integer;
  RPNLength: Word;
  TokenArraySizePos, RecordSizePos, FinalPos: Cardinal;
begin
  RPNLength := 0;
  FormulaResult := 0.0;

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(WordToLE(17 + RPNLength));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  AStream.WriteByte($0);
  AStream.WriteByte($0);
  AStream.WriteByte($0);

  { Result of the formula in IEE 754 floating-point value }
  AStream.WriteBuffer(FormulaResult, 8);

  { 0 = Do not recalculate
    1 = Always recalculate }
  AStream.WriteByte($1);

  { Formula }

  { The size of the token array is written later,
    because it's necessary to calculate if first,
    and this is done at the same time it is written }
  TokenArraySizePos := AStream.Position;
  AStream.WriteByte(RPNLength);

  { Formula data (RPN token array) }
  for i := 0 to Length(AFormula) - 1 do
  begin
    { Token identifier }
    AStream.WriteByte(AFormula[i].TokenID);
    Inc(RPNLength);

    { Additional data }
    case AFormula[i].TokenID of

    { binary operation tokens }

    INT_EXCEL_TOKEN_TADD, INT_EXCEL_TOKEN_TSUB, INT_EXCEL_TOKEN_TMUL,
     INT_EXCEL_TOKEN_TDIV, INT_EXCEL_TOKEN_TPOWER: begin end;

    INT_EXCEL_TOKEN_TNUM:
    begin
      AStream.WriteBuffer(AFormula[i].DoubleValue, 8);
      Inc(RPNLength, 8);
    end;

    INT_EXCEL_TOKEN_TREFR, INT_EXCEL_TOKEN_TREFV, INT_EXCEL_TOKEN_TREFA:
    begin
      AStream.WriteWord(AFormula[i].Row and MASK_EXCEL_ROW);
      AStream.WriteByte(AFormula[i].Col);
      Inc(RPNLength, 3);
    end;

    end;
  end;

  { Write sizes in the end, after we known them }
  FinalPos := AStream.Position;
  AStream.position := TokenArraySizePos;
  AStream.WriteByte(RPNLength);
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(17 + RPNLength));
  AStream.position := FinalPos;
end;

{*******************************************************************
*  TsSpreadBIFF5Writer.WriteLabel ()
*
*  DESCRIPTION:    Writes an Excel 8 LABEL record
*
*                  Writes a string to the sheet
*
*******************************************************************}
procedure TsSpreadBIFF5Writer.WriteLabel(AStream: TStream; const ARow,
  ACol: Word; const AValue: string);
var
  L: Byte;
begin
  L := Length(AValue);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_LABEL));
  AStream.WriteWord(WordToLE(8 + L));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  AStream.WriteByte($0);
  AStream.WriteByte($0);
  AStream.WriteByte($0);

  { String with 8-bit size }
  AStream.WriteByte(L);
  AStream.WriteBuffer(Pointer(AValue)^, L);
end;

{*******************************************************************
*  TsSpreadBIFF5Writer.WriteNumber ()
*
*  DESCRIPTION:    Writes an Excel 5 NUMBER record
*
*                  Writes a number (64-bit floating point) to the sheet
*
*******************************************************************}
procedure TsSpreadBIFF5Writer.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_NUMBER));
  AStream.WriteWord(WordToLE(14));

  { BIFF Record data }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { Index to XF record }
  AStream.WriteWord($0);

  { IEE 754 floating-point value }
  AStream.WriteBuffer(AValue, 8);
end;

procedure TsSpreadBIFF5Writer.WriteXF(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_XF));
  AStream.WriteWord(WordToLE(12));

  { Index to FONT record }
  AStream.WriteByte($00);

  { Index to FORMAT record }
  AStream.WriteByte($00);

  { XF type, cell protection and parent style XF }
  AStream.WriteWord(WordToLE(INT_EXCEL_XF_TYPE_PROT_STYLEXF));

  { Alignment, text break and text orientation }
  AStream.WriteByte($00);

  { Flags for used attribute groups }
  AStream.WriteByte($00);

  { XF_AREA_34 - Cell background area }
  AStream.WriteWord($0000);

  { XF_BORDER_34 - Cell border lines }
  AStream.WriteDWord($00000000);
end;

end.

