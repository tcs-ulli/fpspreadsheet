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
    function GetBorderStyle(ABorder: TsCellBorder): TsCellBorderStyle;
    function GetCellFormat: TsCellFormat;
    function GetFont: TsFont;
    function GetFontIndex: integer;
    function GetHorAlignment: TsHorAlignment;
    function GetNumberFormat: TsNumberFormat;
    function GetNumberFormatStr: String;
    function GetTextRotation: TsTextRotation;
    function GetVertAlignment: TsVertAlignment;
    function GetWordwrap: Boolean;
    procedure SetBackgroundColor(AValue: TsColor);
    procedure SetBorder(AValue: TsCellBorders);
    procedure SetBorderStyle(ABorder: TsCellBorder; AValue: TsCellBorderStyle);
    procedure SetFontIndex(AValue: Integer);

  protected
    function GetWorkbook: TsWorkbook;
  public
    property BackgroundColor: TsColor read GetBackgroundColor write SetBackgroundColor;
    property Border: TsCellBorders read GetBorder write SetBorder;
    property CellFormat: TsCellFormat read GetCellFormat;
    property Font: TsFont read GetFont;
    property FontIndex: Integer read GetFontIndex write SetFontIndex;
    property HorAlignment: TsHorAlignment read GetHorAlignment;
    property NumberFormat: TsNumberFormat read GetNumberFormat;
    property NumberFormatStr: String read GetNumberFormatStr;
    property TextRotation: TsTextRotation read GetTextRotation;
    property VertAlignment: TsVertAlignment read GetVertAlignment;
    property Wordwrap: Boolean read GetWordwrap;
    property Workbook: TsWorkbook read GetWorkbook;
  end;

implementation

function TCellHelper.GetBackgroundColor: TsColor;
begin
  Result := Worksheet.ReadBackgroundColor(@self);
end;

function TCellHelper.GetBorder: TsCellBorders;
begin
  Result := Worksheet.ReadCellBorders(@self);
end;

function TCellHelper.GetBorderStyle(ABorder: TsCellBorder): TsCellBorderStyle;
begin
  Result := Worksheet.ReadCellBorderStyle(@self, ABorder);
end;

function TCellHelper.GetCellFormat: TsCellFormat;
begin
  Result := Workbook.GetCellFormat(FormatIndex);
end;

function TCellHelper.GetFont: TsFont;
begin
  Result := Worksheet.ReadCellFont(@self);
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
  Result := Worksheet.ReadHorAlignment(@Self);
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
  Result := Worksheet.ReadTextRotation(@Self);
end;

function TCellHelper.GetVertAlignment: TsVertAlignment;
begin
  Result := Worksheet.ReadVertAlignment(@self);
end;

function TCellHelper.GetWordwrap: Boolean;
begin
  Result := Worksheet.ReadWordwrap(@self);
end;

function TCellHelper.GetWorkbook: TsWorkbook;
begin
  Result := Worksheet.Workbook;
end;

procedure TCellHelper.SetBackgroundColor(AValue: TsColor);
begin
  Worksheet.WriteBackgroundColor(@self, AValue);
end;

procedure TCellHelper.SetBorder(AValue: TsCellBorders);
begin
  Worksheet.WriteBorders(@self, AValue);
end;

procedure TCellHelper.SetBorderStyle(ABorder: TsCellBorder;
  AValue: TsCellBorderStyle);
begin
  Worksheet.WriteBorderStyle(@self, ABorder, AValue);
end;

procedure TCellHelper.SetFontIndex(AValue: Integer);
begin
  Worksheet.WriteFont(@self, AValue);
end;

end.

