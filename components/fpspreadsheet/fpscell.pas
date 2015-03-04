unit fpsCell;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, fpstypes, fpspreadsheet;

type
  TCellHelper = record helper for TCell
  private
    function GetBackgroundColor: TsColor;
    function GetBorder: TsCellBorders;
    function GetBorderStyle(const ABorder: TsCellBorder): TsCellBorderStyle;
    function GetBorderStyles: TsCellBorderStyles;
    function GetCellFormat: TsCellFormat;
    function GetComment: String;
    function GetFont: TsFont;
    function GetFontIndex: integer;
    function GetHorAlignment: TsHorAlignment;
    function GetHyperlink: TsHyperlink;
    function GetNumberFormat: TsNumberFormat;
    function GetNumberFormatStr: String;
    function GetTextRotation: TsTextRotation;
    function GetUsedFormattingFields: TsUsedFormattingFields;
    function GetVertAlignment: TsVertAlignment;
    function GetWordwrap: Boolean;
    procedure SetBackgroundColor(const AValue: TsColor);
    procedure SetBorder(const AValue: TsCellBorders);
    procedure SetBorderStyle(const ABorder: TsCellBorder; const AValue: TsCellBorderStyle);
    procedure SetBorderStyles(const AValue: TsCellBorderStyles);
    procedure SetCellFormat(const AValue: TsCellFormat);
    procedure SetComment(const AValue: String);
    procedure SetFontIndex(const AValue: Integer);
    procedure SetHorAlignment(const AValue: TsHorAlignment);
    procedure SetHyperlink(const AValue: TsHyperlink);
    procedure SetNumberFormat(const AValue: TsNumberFormat);
    procedure SetNumberFormatStr(const AValue: String);
    procedure SetTextRotation(const AValue: TsTextRotation);
    procedure SetUsedFormattingFields(const AValue: TsUsedFormattingFields);
    procedure SetVertAlignment(const AValue: TsVertAlignment);
    procedure SetWordwrap(const AValue: Boolean);

  protected
    function GetWorkbook: TsWorkbook; inline;
    function GetWorksheet: TsWorksheet; inline;

  public
    property BackgroundColor: TsColor
      read GetBackgroundColor write SetBackgroundColor;
    property Border: TsCellBorders
      read GetBorder write SetBorder;
    property BorderStyle[ABorder: TsCellBorder]: TsCellBorderStyle
      read GetBorderStyle write SetBorderStyle;
    property BorderStyles: TsCellBorderStyles
      read GetBorderStyles write SetBorderStyles;
    property CellFormat: TsCellFormat
      read GetCellFormat write SetCellFormat;
    property Comment: String
      read GetComment write SetComment;
    property Font: TsFont read GetFont;
    property FontIndex: Integer
      read GetFontIndex write SetFontIndex;
    property HorAlignment: TsHorAlignment
      read GetHorAlignment write SetHorAlignment;
    property Hyperlink: TsHyperlink
      read GetHyperlink write SetHyperlink;
    property NumberFormat: TsNumberFormat
      read GetNumberFormat write SetNumberFormat;
    property NumberFormatStr: String
      read GetNumberFormatStr write SetNumberFormatStr;
    property TextRotation: TsTextRotation
      read GetTextRotation write SetTextRotation;
    property UsedFormattingFields: TsUsedFormattingFields
      read GetUsedFormattingFields write SetUsedFormattingFields;
    property VertAlignment: TsVertAlignment
      read GetVertAlignment write SetVertAlignment;
    property Wordwrap: Boolean
      read GetWordwrap write SetWordwrap;
    property Workbook: TsWorkbook read GetWorkbook;
  end;

implementation

function TCellHelper.GetBackgroundColor: TsColor;
begin
  Result := GetWorksheet.ReadBackgroundColor(@self);
end;

function TCellHelper.GetBorder: TsCellBorders;
begin
  Result := GetWorksheet.ReadCellBorders(@self);
end;

function TCellHelper.GetBorderStyle(const ABorder: TsCellBorder): TsCellBorderStyle;
begin
  Result := GetWorksheet.ReadCellBorderStyle(@self, ABorder);
end;

function TCellHelper.GetBorderStyles: TsCellBorderStyles;
begin
  Result := GetWorksheet.ReadCellBorderStyles(@self);
end;

function TCellHelper.GetCellFormat: TsCellFormat;
begin
  Result := GetWorkbook.GetCellFormat(FormatIndex);
end;

function TCellHelper.GetComment: String;
begin
  Result := GetWorksheet.ReadComment(@self);
end;

function TCellHelper.GetFont: TsFont;
begin
  Result := GetWorksheet.ReadCellFont(@self);
end;

function TCellHelper.GetFontIndex: Integer;
var
  fmt: PsCellFormat;
begin
  fmt := Workbook.GetPointerToCellFormat(FormatIndex);
  Result := fmt^.FontIndex;
end;

function TCellHelper.GetHorAlignment: TsHorAlignment;
begin
  Result := GetWorksheet.ReadHorAlignment(@Self);
end;

function TCellHelper.GetHyperlink: TsHyperlink;
begin
  Result := GetWorksheet.ReadHyperlink(@self);
end;

function TCellHelper.GetNumberFormat: TsNumberFormat;
var
  fmt: PsCellFormat;
begin
  fmt := Workbook.GetPointerToCellFormat(FormatIndex);
  Result := fmt^.NumberFormat;
end;

function TCellHelper.GetNumberFormatStr: String;
var
  fmt: PsCellFormat;
begin
  fmt := Workbook.GetPointerToCellFormat(FormatIndex);
  Result := fmt^.NumberFormatStr;
end;

function TCellHelper.GetTextRotation: TsTextRotation;
begin
  Result := GetWorksheet.ReadTextRotation(@Self);
end;

function TCellHelper.GetUsedFormattingFields: TsUsedFormattingFields;
begin
  Result := GetWorksheet.ReadUsedFormatting(@Self);
end;

function TCellHelper.GetVertAlignment: TsVertAlignment;
begin
  Result := GetWorksheet.ReadVertAlignment(@self);
end;

function TCellHelper.GetWordwrap: Boolean;
begin
  Result := GetWorksheet.ReadWordwrap(@self);
end;

function TCellHelper.GetWorkbook: TsWorkbook;
begin
  Result := GetWorksheet.Workbook;
end;

function TCellHelper.GetWorksheet: TsWorksheet;
begin
  Result := TsWorksheet(Worksheet);
end;

procedure TCellHelper.SetBackgroundColor(const AValue: TsColor);
begin
  GetWorksheet.WriteBackgroundColor(@self, AValue);
end;

procedure TCellHelper.SetBorder(const AValue: TsCellBorders);
begin
  GetWorksheet.WriteBorders(@self, AValue);
end;

procedure TCellHelper.SetBorderStyle(const ABorder: TsCellBorder;
  const AValue: TsCellBorderStyle);
begin
  GetWorksheet.WriteBorderStyle(@self, ABorder, AValue);
end;

procedure TCellHelper.SetBorderStyles(const AValue: TsCellBorderStyles);
begin
  GetWorksheet.WriteBorderStyles(@self, AValue);
end;

procedure TCellHelper.SetCellFormat(const AValue: TsCellFormat);
begin
  GetWorksheet.WriteCellFormat(@self, AValue);
end;

procedure TCellHelper.SetComment(const AValue: String);
begin
  GetWorksheet.WriteComment(@self, AValue);
end;

procedure TCellHelper.SetFontIndex(const AValue: Integer);
begin
  GetWorksheet.WriteFont(@self, AValue);
end;

procedure TCellHelper.SetHorAlignment(const AValue: TsHorAlignment);
begin
  GetWorksheet.WriteHorAlignment(@self, AValue);
end;

procedure TCellHelper.SetHyperlink(const AValue: TsHyperlink);
begin
  GetWorksheet.WriteHyperlink(@self, AValue.Target, AValue.Tooltip);
end;

procedure TCellHelper.SetNumberFormat(const AValue: TsNumberFormat);
var
  fmt: TsCellFormat;
begin
  fmt := Workbook.GetCellFormat(FormatIndex);
  fmt.NumberFormat := AValue;
  GetWorksheet.WriteCellFormat(@self, fmt);
end;

procedure TCellHelper.SetNumberFormatStr(const AValue: String);
var
  fmt: TsCellFormat;
begin
  fmt := Workbook.GetCellFormat(FormatIndex);
  fmt.NumberFormatStr := AValue;
  GetWorksheet.WriteCellFormat(@self, fmt);
end;

procedure TCellHelper.SetTextRotation(const AValue: TsTextRotation);
begin
  GetWorksheet.WriteTextRotation(@self, AValue);
end;

procedure TCellHelper.SetUsedFormattingFields(const AValue: TsUsedFormattingFields);
begin
  GetWorksheet.WriteUsedFormatting(@self, AValue);
end;

procedure TCellHelper.SetVertAlignment(const AValue: TsVertAlignment);
begin
  GetWorksheet.WriteVertAlignment(@self, AValue);
end;

procedure TCellHelper.SetWordwrap(const AValue: Boolean);
begin
  GetWorksheet.WriteWordwrap(@self, AValue);
end;


end.

