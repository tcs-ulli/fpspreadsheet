unit xlscommon;

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  fpspreadsheet,
  fpsutils;

const
  { Formula constants TokenID values }

  { Binary Operator Tokens }
  INT_EXCEL_TOKEN_TADD    = $03;
  INT_EXCEL_TOKEN_TSUB    = $04;
  INT_EXCEL_TOKEN_TMUL    = $05;
  INT_EXCEL_TOKEN_TDIV    = $06;
  INT_EXCEL_TOKEN_TPOWER  = $07; // Power Exponentiation
  INT_EXCEL_TOKEN_TCONCAT = $08;
  INT_EXCEL_TOKEN_TLT     = $09; // Less than
  INT_EXCEL_TOKEN_TLE     = $0A; // Less than or equal
  INT_EXCEL_TOKEN_TEQ     = $0B; // Equal
  INT_EXCEL_TOKEN_TGE     = $0C; // Greater than or equal
  INT_EXCEL_TOKEN_TGT     = $0D; // Greater than
  INT_EXCEL_TOKEN_TNE     = $0E; // Not equal
  INT_EXCEL_TOKEN_TISECT  = $0F; // Cell range intersection
  INT_EXCEL_TOKEN_TLIST   = $10; // Cell range list
  INT_EXCEL_TOKEN_TRANGE  = $11; // Cell range

  { Constant Operand Tokens }
  INT_EXCEL_TOKEN_TNUM    = $1F;

  { Operand Tokens }
  INT_EXCEL_TOKEN_TREFR   = $24;
  INT_EXCEL_TOKEN_TREFV   = $44;
  INT_EXCEL_TOKEN_TREFA   = $64;

  { Function Tokens }
  INT_EXCEL_TOKEN_FUNCVAR_R = $22;
  INT_EXCEL_TOKEN_FUNCVAR_V = $42;
  INT_EXCEL_TOKEN_FUNCVAR_A = $62;

  { Built-in functions }
  INT_EXCEL_SHEET_FUNC_ABS = 24;
  INT_EXCEL_SHEET_FUNC_ROUND = 27;

  { Built In Color Pallete Indexes }
  BUILT_IN_COLOR_PALLETE_BLACK     = $08; // 000000H
  BUILT_IN_COLOR_PALLETE_WHITE     = $09; // FFFFFFH
  BUILT_IN_COLOR_PALLETE_RED       = $0A; // FF0000H
  BUILT_IN_COLOR_PALLETE_GREEN     = $0B; // 00FF00H
  BUILT_IN_COLOR_PALLETE_BLUE      = $0C; // 0000FFH
  BUILT_IN_COLOR_PALLETE_YELLOW    = $0D; // FFFF00H
  BUILT_IN_COLOR_PALLETE_MAGENTA   = $0E; // FF00FFH
  BUILT_IN_COLOR_PALLETE_CYAN      = $0F; // 00FFFFH
  BUILT_IN_COLOR_PALLETE_DARK_RED  = $10; // 800000H
  BUILT_IN_COLOR_PALLETE_DARK_GREEN= $11; // 008000H
  BUILT_IN_COLOR_PALLETE_DARK_BLUE = $12; // 000080H
  BUILT_IN_COLOR_PALLETE_OLIVE     = $13; // 808000H
  BUILT_IN_COLOR_PALLETE_PURPLE    = $14; // 800080H
  BUILT_IN_COLOR_PALLETE_TEAL      = $15; // 008080H
  BUILT_IN_COLOR_PALLETE_SILVER    = $16; // C0C0C0H
  BUILT_IN_COLOR_PALLETE_GREY      = $17; // 808080H

  EXTRA_COLOR_PALETTE_GREY10PCT    = $18; // E6E6E6H
  EXTRA_COLOR_PALETTE_GREY20PCT    = $19; // E6E6E6H

type

  { TsSpreadBIFFReader }

  TsSpreadBIFFReader = class(TsCustomSpreadReader)
  protected
  end;

  { TsSpreadBIFFWriter }

  TsSpreadBIFFWriter = class(TsCustomSpreadWriter)
  protected
    FLastRow: Integer;
    FLastCol: Word;
    function FPSColorToEXCELPallete(AColor: TsColor): Word;
    procedure GetLastRowCallback(ACell: PCell; AStream: TStream);
    function GetLastRowIndex(AWorksheet: TsWorksheet): Integer;
    procedure GetLastColCallback(ACell: PCell; AStream: TStream);
    function GetLastColIndex(AWorksheet: TsWorksheet): Word;
    function FormulaElementKindToExcelTokenID(AElementKind: TFEKind): Byte;
  end;

implementation

function TsSpreadBIFFWriter.FPSColorToEXCELPallete(AColor: TsColor): Word;
begin
  case AColor of
    scBlack: Result := BUILT_IN_COLOR_PALLETE_BLACK;
    scWhite: Result := BUILT_IN_COLOR_PALLETE_WHITE;
    scRed: Result := BUILT_IN_COLOR_PALLETE_RED;
    scGREEN: Result := BUILT_IN_COLOR_PALLETE_GREEN;
    scBLUE: Result := BUILT_IN_COLOR_PALLETE_BLUE;
    scYELLOW: Result := BUILT_IN_COLOR_PALLETE_YELLOW;
    scMAGENTA: Result := BUILT_IN_COLOR_PALLETE_MAGENTA;
    scCYAN: Result := BUILT_IN_COLOR_PALLETE_CYAN;
    scDarkRed: Result := BUILT_IN_COLOR_PALLETE_DARK_RED;
    scDarkGreen: Result := BUILT_IN_COLOR_PALLETE_DARK_GREEN;
    scDarkBlue: Result := BUILT_IN_COLOR_PALLETE_DARK_BLUE;
    scOLIVE: Result := BUILT_IN_COLOR_PALLETE_OLIVE;
    scPURPLE: Result := BUILT_IN_COLOR_PALLETE_PURPLE;
    scTEAL: Result := BUILT_IN_COLOR_PALLETE_TEAL;
    scSilver: Result := BUILT_IN_COLOR_PALLETE_SILVER;
    scGrey: Result := BUILT_IN_COLOR_PALLETE_GREY;
    //
    scGrey10pct: Result := EXTRA_COLOR_PALETTE_GREY10PCT;
    scGrey20pct: Result := EXTRA_COLOR_PALETTE_GREY20PCT;
  end;
end;

procedure TsSpreadBIFFWriter.GetLastRowCallback(ACell: PCell; AStream: TStream);
begin
  if ACell^.Row > FLastRow then FLastRow := ACell^.Row;
end;

function TsSpreadBIFFWriter.GetLastRowIndex(AWorksheet: TsWorksheet): Integer;
begin
  FLastRow := 0;
  IterateThroughCells(nil, AWorksheet.Cells, GetLastRowCallback);
  Result := FLastRow;
end;

procedure TsSpreadBIFFWriter.GetLastColCallback(ACell: PCell; AStream: TStream);
begin
  if ACell^.Col > FLastCol then FLastCol := ACell^.Col;
end;

function TsSpreadBIFFWriter.GetLastColIndex(AWorksheet: TsWorksheet): Word;
begin
  FLastCol := 0;
  IterateThroughCells(nil, AWorksheet.Cells, GetLastColCallback);
  Result := FLastCol;
end;

function TsSpreadBIFFWriter.FormulaElementKindToExcelTokenID(
  AElementKind: TFEKind): Byte;
begin
  case AElementKind of
    { Operand Tokens }
    fekCell:  Result := INT_EXCEL_TOKEN_TREFR;
    fekCellRange: Result := INT_EXCEL_TOKEN_TRANGE;
    fekNum:   Result := INT_EXCEL_TOKEN_TNUM;
    { Basic operations }
    fekAdd:   Result := INT_EXCEL_TOKEN_TADD;
    fekSub:   Result := INT_EXCEL_TOKEN_TSUB;
    fekDiv:   Result := INT_EXCEL_TOKEN_TDIV;
    fekMul:   Result := INT_EXCEL_TOKEN_TMUL;
    { Build-in Functions}
    fekABS:   Result := INT_EXCEL_SHEET_FUNC_ABS;
    fekROUND: Result := INT_EXCEL_SHEET_FUNC_ROUND;
    { Other operations }
    fekOpSUM: Result := 0;
  else
    Result := 0;
  end;
end;

end.

