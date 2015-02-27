unit beBIFFGrid;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Controls, Grids, fpstypes, fpspreadsheet, beTypes;

type
  TBIFFBuffer = array of byte;

  TBIFFDetailsEvent = procedure(Sender: TObject; ADetails: TStrings) of object;

  TBIFFGrid = class(TStringGrid)
  private
    FRecType: Word;
    FBuffer: TBIFFBuffer;
    FBufferIndex: LongWord;
    FFormat: TsSpreadsheetFormat;
    FInfo: Integer;
    FCurrRow: Integer;
    FDetails: TStrings;
    FOnDetails: TBIFFDetailsEvent;
    function  GetStringType: String;

    procedure ShowBackup;
    procedure ShowBlankCell;
    procedure ShowBOF;
    procedure ShowBookBool;
    procedure ShowBoolCell;
    procedure ShowBottomMargin;
    procedure ShowCalcCount;
    procedure ShowCalcMode;
    procedure ShowCellAddress;
    procedure ShowCellAddressRange(AFormat: TsSpreadsheetFormat);
    procedure ShowClrtClient;
    procedure ShowCodePage;
    procedure ShowColInfo;
    procedure ShowColWidth;
    procedure ShowContinue;
    procedure ShowCountry;
    procedure ShowDateMode;
    procedure ShowDBCell;
    procedure ShowDefColWidth;
    procedure ShowDefinedName;
    procedure ShowDefRowHeight;
    procedure ShowDelta;
    procedure ShowDimensions;
    procedure ShowDSF;
    procedure ShowEOF;
    procedure ShowExcel9File;
    procedure ShowExternalBook;
    procedure ShowExternCount;
    procedure ShowExternSheet;
    procedure ShowFileSharing;
    procedure ShowFnGroupCount;
    procedure ShowFont;
    procedure ShowFontColor;
    procedure ShowFooter;
    procedure ShowFormat;
    procedure ShowFormatCount;
    procedure ShowFormula;
    procedure ShowFormulaTokens(ATokenBytes: Integer);
    procedure ShowGCW;
    procedure ShowHeader;
    procedure ShowHideObj;
    procedure ShowHyperLink;
    procedure ShowHyperLinkTooltip;
    procedure ShowInteger;
    procedure ShowInterfaceEnd;
    procedure ShowInterfaceHdr;
    procedure ShowIteration;
    procedure ShowIXFE;
    procedure ShowLabelCell;
    procedure ShowLabelSSTCell;
    procedure ShowLeftMargin;
    procedure ShowMergedCells;
    procedure ShowMMS;
    procedure ShowMSODrawing;
    procedure ShowMulBlank;
    procedure ShowMulRK;
    procedure ShowNote;
    procedure ShowNumberCell;
    procedure ShowObj;
    procedure ShowPageSetup;
    procedure ShowPalette;
    procedure ShowPane;
    procedure ShowPassword;
    procedure ShowPrecision;
    procedure ShowPrintGridLines;
    procedure ShowPrintHeaders;
    procedure ShowProt4Rev;
    procedure ShowProt4RevPass;
    procedure ShowProtect;
    procedure ShowRecalc;
    procedure ShowRefMode;
    procedure ShowRefreshAll;
    procedure ShowRightMargin;
    procedure ShowRK;
    procedure ShowRow;
    procedure ShowSelection;
    procedure ShowSharedFormula;
    procedure ShowSheet;
    procedure ShowSheetPR;
    procedure ShowSST;
    procedure ShowStandardWidth;
    procedure ShowString;
    procedure ShowStyle;
    procedure ShowStyleExt;
    procedure ShowTabID;
    procedure ShowTopMargin;
    procedure ShowTXO;
    procedure ShowWindow1;
    procedure ShowWindow2;
    procedure ShowWindowProtect;
    procedure ShowWriteAccess;
    procedure ShowWriteProt;
    procedure ShowXF;
    procedure ShowXFCRC;
    procedure ShowXFEXT;

  protected
    procedure Click; override;
    procedure DoExtractDetails;
    function DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint): Boolean; override;
    function DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint): Boolean; override;
    procedure ExtractString(ABufIndex: Integer; ALenBytes: Byte; AUnicode: Boolean;
      out AString: String; out ANumBytes: Integer; IgnoreCompressedFlag: Boolean = false);
    procedure PopulateGrid;
    procedure ShowInRow(var ARow: Integer; var AOffs: LongWord; ASize: Word;
      AValue,ADescr: String; ADescrOnly: Boolean = false);
    procedure ShowRowColData(var ABufIndex: LongWord);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure SetBIFFNodeData(AData: TBIFFNodeData; ABuffer: TBIFFBuffer; AFormat: TsSpreadsheetFormat);
//    procedure SetRecordType(ARecType: Word; ABuffer: TBIFFBuffer; AFormat: TsSpreadsheetFormat);

  published
    property OnDetails: TBIFFDetailsEvent read FOnDetails write FOnDetails;
  end;

implementation

uses
  StrUtils, Math, lazutf8,
  fpsutils,
  beBIFFUtils;

const
  ABS_REL: array[boolean] of string = ('abs', 'rel');

constructor TBIFFGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  ColCount := 4;
  FixedCols := 0;
  RowCount := 2;
  Cells[0, 0] := 'Offset';
  Cells[1, 0] := 'Size';
  Cells[2, 0] := 'Value';
  Cells[3, 0] := 'Description';
  ColWidths[0] := 60;
  ColWidths[1] := 60;
  ColWidths[2] := 120;
  ColWidths[3] := 350;
  Options := Options
    + [goThumbTracking, goColSizing, goTruncCellHints, goCellHints]
    - [goVertLine, goSmoothScroll];
  MouseWheelOption := mwGrid;
  FDetails := TStringList.Create;
end;


destructor TBIFFGrid.Destroy;
begin
  FDetails.Free;
  inherited;
end;


procedure TBIFFGrid.Click;
begin
  inherited;
  if (FBuffer <> nil) then
    DoExtractDetails;
end;


procedure TBIFFGrid.DoExtractDetails;
begin
  if Assigned(FOnDetails) then begin
    PopulateGrid;
    FOnDetails(self, FDetails);
  end;
end;

function TBIFFGrid.DoMouseWheelDown(Shift: TShiftState; MousePos: TPoint
  ): Boolean;
begin
  Result := inherited;
  Click;
end;

function TBIFFGrid.DoMouseWheelUp(Shift: TShiftState; MousePos: TPoint
  ): Boolean;
begin
  Result := inherited;
  Click;
end;

procedure TBIFFGrid.ExtractString(ABufIndex: Integer; ALenBytes: Byte; AUnicode: Boolean;
  out AString: String; out ANumBytes: Integer; IgnoreCompressedFlag: Boolean = false);
var
  ls: Integer;
  sa: ansiString;
  sw: WideString;
  w: Word;
  optn: Byte;
begin
  if Length(FBuffer) = 0 then begin
    AString := '';
    ANumBytes := 0;
    exit;
  end;
  if ALenBytes = 1 then
    ls := FBuffer[ABufIndex]
  else begin
    Move(FBuffer[ABufIndex], w, 2);
    ls := WordLEToN(w);
  end;
  if AUnicode then begin
    optn := FBuffer[ABufIndex + ALenBytes];
    if (optn  and $01 = 0) and (not IgnoreCompressedFlag)
    then begin   // compressed --> 1 byte per character
      SetLength(sa, ls);
      ANumbytes := ls*SizeOf(AnsiChar) + ALenBytes + 1;
      Move(FBuffer[ABufIndex + ALenBytes + 1], sa[1], ls*SizeOf(AnsiChar));
      AString := AnsiToUTF8(sa);
    end else begin
      SetLength(sw, ls);
      ANumBytes := ls*SizeOf(WideChar) + ALenBytes + 1;
      Move(FBuffer[ABufIndex + ALenBytes + 1], sw[1], ls*SizeOf(WideChar));
      AString := UTF8Encode(WideStringLEToN(sw));
    end;
  end else begin
    SetLength(sa, ls);
    ANumBytes := ls*SizeOf(AnsiChar) + ALenBytes;
    Move(FBuffer[ABufIndex + ALenBytes], sa[1], ls*SizeOf(AnsiChar));
    AString := AnsiToUTF8(sa);
  end;
end;


function TBIFFGrid.GetStringType: String;
begin
  case FFormat of
    sfExcel2: Result := 'Byte';
    sfExcel5: Result := 'Byte';
    sfExcel8: Result := 'Unicode';
  end;
end;


procedure TBIFFGrid.PopulateGrid;
begin
  FBufferIndex := 0;
  FCurrRow := FixedRows;
  FDetails.Clear;
  case FRecType of
    $0000, $0200:
      ShowDimensions;
    $0001, $0201:
      ShowBlankCell;
    $0002:
      ShowInteger;
    $0003, $0203:
      ShowNumberCell;
    $0004, $0204:
      ShowLabelCell;
    $0005, $0205:
      ShowBoolCell;
    $0006:
      ShowFormula;
    $0007, $0207:
      ShowString;
    $0008, $0208:
      ShowRow;
    $0009, $0209, $0409, $0809:
      ShowBOF;
    $000A:
      ShowEOF;
    $000C:
      ShowCalcCount;
    $000D:
      ShowCalcMode;
    $000E:
      ShowPrecision;
    $000F:
      ShowRefMode;
    $0010:
      ShowDelta;
    $0011:
      ShowIteration;
    $0012:
      ShowProtect;
    $0013:
      ShowPassword;
    $0014:
      ShowHeader;
    $0015:
      ShowFooter;
    $0016:
      ShowExternCount;
    $0017:
      ShowExternSheet;
    $0018, $0218:
      ShowDefinedName;
    $0019:
      ShowWindowProtect;
    $001C:
      ShowNote;
    $001D:
      ShowSelection;
    $001E, $041E:
      ShowFormat;
    $001F:
      ShowFormatCount;
    $0022:
      ShowDateMode;
    $0024:
      ShowColWidth;
    $0025, $0225:
      ShowDefRowHeight;
    $0026:
      ShowLeftMargin;
    $0027:
      ShowRightMargin;
    $0028:
      ShowTopMargin;
    $0029:
      ShowBottomMargin;
    $002A:
      ShowPrintHeaders;
    $002B:
      ShowPrintGridLines;
    $0031:
      ShowFont;
    $003C:
      ShowContinue;
    $003D:
      ShowWindow1;
    $003E, $023E:
      ShowWindow2;
    $0040:
      ShowBackup;
    $0041:
      ShowPane;
    $0042:
      ShowCodePage;
    $0043:
      ShowXF;
    $0044:
      ShowIXFE;
    $0045:
      ShowFontColor;
    $0055:
      ShowDefColWidth;
    $005B:
      ShowFileSharing;
    $005C:
      ShowWriteAccess;
    $005D:
      ShowObj;
    $005F:
      ShowRecalc;
    $007D:
      ShowColInfo;
    $0081:
      ShowSheetPR;
    $0085:
      ShowSheet;
    $0086:
      ShowWriteProt;
    $008C:
      ShowCountry;
    $008D:
      ShowHideObj;
    $0092:
      ShowPalette;
    $099:
      ShowStandardWidth;
    $00A1:
      ShowPageSetup;
    $00AB:
      ShowGCW;
    $00C1:
      ShowMMS;
    $009C:
      ShowFnGroupCount;
    $00BE:
      ShowMulBlank;
    $00BD:
      ShowMulRK;
    $00D7:
      ShowDBCell;
    $00DA:
      ShowBookBool;
    $00E0:
      ShowXF;
    $00E1:
      ShowInterfaceHdr;
    $00E2:
      ShowInterfaceEnd;
    $00E5:
      ShowMergedCells;
    $00EC:
      ShowMSODrawing;
    $00FC:
      ShowSST;
    $00FD:
      ShowLabelSSTCell;
    $013D:
      ShowTabID;
    $0161:
      ShowDSF;
    $01AE:
      ShowExternalBook;
    $01AF:
      ShowProt4Rev;
    $01B6:
      ShowTXO;
    $01B7:
      ShowRefreshAll;
    $01B8:
      ShowHyperlink;
    $01BC:
      ShowProt4RevPass;
    $01C0:
      ShowExcel9File;
    $027E:
      ShowRK;
    $0293:
      ShowStyle;
    $04BC:
      ShowSharedFormula;
    $0800:
      ShowHyperlinkTooltip;
    $087C:
      ShowXFCRC;
    $087D:
      ShowXFEXT;
    $0892:
      ShowStyleExt;
    $105C:
      ShowClrtClient;
    else
      RowCount := 2;
      Rows[1].Clear;
  end;
end;

 {
procedure TBIFFGrid.SetRecordType(ARecType: Word; ABuffer: TBIFFBuffer;
  AFormat: TsSpreadsheetFormat);
begin
  FFormat := AFormat;
  FRecType := ARecType;
  SetLength(FBuffer, Length(ABuffer));
  if Length(FBuffer) > 0 then
    Move(ABuffer[0], FBuffer[0], Length(FBuffer));
  PopulateGrid;
  if Assigned(FOnDetails) then FOnDetails(self, FDetails);
end;
}

procedure TBIFFGrid.SetBIFFNodeData(AData: TBIFFNodeData; ABuffer: TBIFFBuffer;
  AFormat: TsSpreadsheetFormat);
begin
  if AData = nil then
    exit;
  FFormat := AFormat;
  FRecType := AData.RecordID;
  FInfo := AData.Tag;
  SetLength(FBuffer, Length(ABuffer));
  if Length(FBuffer) > 0 then
    Move(ABuffer[0], FBuffer[0], Length(FBuffer));
  PopulateGrid;
  if Assigned(FOnDetails) then FOnDetails(self, FDetails);
end;

procedure TBIFFGrid.ShowBackup;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Save backup copy of workbook:'#13);
    if w = 0
      then FDetails.Add('0 = no backup')
      else FDetails.Add('1 = backup copy is saved when workbook is saved');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    'Save backup copy of workbook');
end;


procedure TBIFFGrid.ShowBlankCell;
var
  numBytes: Integer;
  b: Byte = 0;
  w: Word = 0;
  dbl: Double;
begin
  RowCount := IfThen(FFormat = sfExcel2, FixedRows + 5, FixedRows + 3);
  // Offset 0: Row & Offset 2: Column
  ShowRowColData(FBufferIndex);

  // Offset 4: Cell attributes (BIFF2) or XF record index (> BIFF2)
  if FFormat = sfExcel2 then begin
    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell protection and XF index:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: XF Index', [b and $3F]));
      case b and $40 of
        0: FDetails.Add('Bit 6 = 0: Cell is NOT locked.');
        1: FDetails.Add('Bit 6 = 1: Cell is locked.');
      end;
      case b and $80 of
        0: FDetails.Add('Bit 7 = 0: Formula is NOT hidden.');
        1: FDetails.Add('Bit 7 = 1: Formula is hidden.');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b, b]),
      'Cell protection and XF index');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Indexes to format and font records:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: Index to FORMAT record', [b and $3f]));
      FDetails.Add(Format('Bits 7-6 = %d: Index to FONT record', [(b and $C0) shr 6]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b, b]),
      'Indexes of format and font records');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell style:'#13);
      case b and $07 of
        0: FDetails.Add('Bits 2-0 = 0: Horizontal alignment is GENERAL');
        1: FDetails.Add('Bits 2-0 = 1: Horizontal alignment is LEFT');
        2: FDetails.Add('Bits 2-0 = 2: Horizontal alignment is CENTERED');
        3: FDetails.Add('Bits 2-0 = 3: Horizontal alignment is RIGHT');
        4: FDetails.Add('Bits 2-0 = 4: Horizontal alignment is FILLED');
      end;
      if b and $08 = 0
        then FDetails.Add('Bit 3 = 0: Cell has NO left border')
        else FDetails.Add('Bit 3 = 1: Cell has left black border');
      if b and $10 = 0
        then FDetails.Add('Bit 4 = 0: Cell has NO right border')
        else FDetails.Add('Bit 4 = 1: Cell has right black border');
      if b and $20 = 0
        then FDetails.Add('Bit 5 = 0: Cell has NO top border')
        else FDetails.Add('Bit 5 = 1: Cell has top black border');
      if b and $40 = 0
        then FDetails.Add('Bit 6 = 0: Cell has NO bottom border')
        else FDetails.Add('Bit 6 = 1: Cell has bottom black border');
      if b and $80 = 0
        then FDetails.Add('Bit 7 = 0: Cell has NO shaded background')
        else FDetails.Add('Bit 7 = 1: Cell has shaded background');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.2x)', [b,b]),
      'Cell style');
  end else
  begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrROw, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Index of XF record');
  end;
end;


procedure TBIFFGrid.ShowBOF;
var
  numBytes: Integer;
  w: Word;
  s: String;
begin
  case FFormat of
    sfExcel2: RowCount := FixedRows + 2;
    { //Excel3 & 4 not supported by fpspreadsheet
    sfExcel3, sfExcel4: RowCount := FixedRows + 3;
    }
    sfExcel5: RowCount := FixedRows + 4;
    sfExcel8: RowCount := FixedRows + 6;
  end;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('BIFF version:'#13);
    case FRecType of
      $0009,
      $0209,
      $0409: FDetails.Add('not used');
      $0809: case FFormat of
               sfExcel5: FDetails.Add('$0500 = BIFF5');
               sfExcel8: FDetails.Add('$0600 = BIFF8');
             end;
      else   case w of
               $0000: FDetails.Add('$0000 = BIFF5');
               $0200: FDetails.Add('$0200 = BIFF2');
               $0300: FDetails.Add('$0300 = BIFF3');
               $0400: FDetails.Add('$0400 = BIFF4');
               $0500: FDetails.Add('$0500 = BIFF5');
               $0600: FDetails.Add('$0600 = BIFF8');
             end;
    end;
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    'BIFF version');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  s := '$0010=Sheet, $0020=Chart, $0040=Macro sheet';
  if FFormat > sfExcel2 then
    s := '$0005=WB globals, $0006=VB module, ' + s + ', $0100=Workspace';
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Type of data:'#13);
    FDetails.Add(Format('$%.4x = %s', [w, BofName(w)]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    Format('Type of data (%s)', [s]));

  if FFormat > sfExcel2 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    { Excel3/4 not supported in fpSpreadsheet
    if FFormat in [sfExcel3, sfExcel4] then
      ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(WordLEToN(w)),
        'not used')
    else}
    begin
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
        'Build identifier (must not be zero)');

      numBytes := 2;
      Move(FBuffer[FBufferIndex], w, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
        'Build year (must not be zero)');
    end;
  end;

  if FFormat = sfExcel8 then begin
    numBytes := 4;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'File history flags');

    numBytes :=4;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'Lowest Excel version that can read all records of this file');
  end;
end;


procedure TBIFFGrid.ShowBookBool;
var
  numbytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Some properties assosciated with notebook:'#13);
    if w and $0001 = 0
      then FDetails.Add('Bit 0 = 0: External link values are saved.')
      else FDetails.Add('Bit 0 = 1: External link values are NOT saved.');
    FDetails.Add('Bit 1: to be ignored');
    if w and $0004 = 0
      then FDetails.Add('Bit 2 = 0: Workbook does not have a mail envelope')
      else FDetails.Add('Bit 2 = 1: Workbook has a mail envelope');
    if w and $0008 = 0
      then FDetails.Add('Bit 3 = 0: Mail envelope is NOT visible.')
      else FDetails.Add('Bit 3 = 1: Mail envelope is visible.');
    if w and $0010 = 0
      then FDetails.Add('Bit 4 = 0: Mail envelope has NOT been initialized.')
      else FDetails.Add('Bit 4 = 1: Mail envelope has been initialized.');
    case (w and $0060) shr 5 of
      0: FDetails.Add('Bits 5-6 (Update external links) = 0: Prompt user to update');
      1: FDetails.Add('Bits 5-6 (Update external linls) = 1: Do not update, and do not prompt user.');
      2: FDetails.Add('Bits 5-6 (Update external links) = 2: Silently update external links.');
    end;
    FDetails.Add('Bit 7: undefined, must be ignored');
    if w and $0100 = 0
      then FDetails.Add('Bit 8 = 0: Do not hide borders of tables that do not contain the active cell')
      else FDetails.Add('Bit 8 = 1: Hide borders of tables that do not contain the active cell');
    FDetails.Add('Bits 9-15: MUST BE zero, MUST be ignored');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
    'Specifies some properties assosciated with a workbook');
end;

procedure TBIFFGrid.ShowBoolCell;
var
  numBytes: Integer;
  w: Word;
  b: Byte;
begin
  if FFormat = sfExcel2 then
    RowCount := FixedRows + 7
  else
    RowCount := FixedRows + 5;

  ShowRowColData(FBufferIndex);

  if FFormat = sfExcel2 then begin
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Cell protection and XF index:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: XF Index', [b and $3F]));
      case b and $40 of
        0: FDetails.Add('Bit 6 = 0: Cell is NOT locked.');
        1: FDetails.Add('Bit 6 = 1: Cell is locked.');
      end;
      case b and $80 of
        0: FDetails.Add('Bit 7 = 0: Formula is NOT hidden.');
        1: FDetails.Add('Bit 7 = 1: Formula is hidden.');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Cell protection and XF index');

    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Indexes to format and font records:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: Index to FORMAT record', [b and $3f]));
      FDetails.Add(Format('Bits 7-6 = %d: Index to FONT record', [(b and $C0) shr 6]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Indexes of format and font records');

    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Cell style:'#13);
      case b and $07 of
        0: FDetails.Add('Bits 2-0 = 0: Horizontal alignment is GENERAL');
        1: FDetails.Add('Bits 2-0 = 1: Horizontal alignment is LEFT');
        2: FDetails.Add('Bits 2-0 = 2: Horizontal alignment is CENTERED');
        3: FDetails.Add('Bits 2-0 = 3: Horizontal alignment is RIGHT');
        4: FDetails.Add('Bits 2-0 = 4: Horizontal alignment is FILLED');
      end;
      if b and $08 = 0
        then FDetails.Add('Bit 3 = 0: Cell has NO left border')
        else FDetails.Add('Bit 3 = 1: Cell has left black border');
      if b and $10 = 0
        then FDetails.Add('Bit 4 = 0: Cell has NO right border')
        else FDetails.Add('Bit 4 = 1: Cell has right black border');
      if b and $20 = 0
        then FDetails.Add('Bit 5 = 0: Cell has NO top border')
        else FDetails.Add('Bit 5 = 1: Cell has top black border');
      if b and $40 = 0
        then FDetails.Add('Bit 6 = 0: Cell has NO bottom border')
        else FDetails.Add('Bit 6 = 1: Cell has bottom black border');
      if b and $80 = 0
        then FDetails.Add('Bit 7 = 0: Cell has NO shaded background')
        else FDetails.Add('Bit 7 = 1: Cell has shaded background');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.2x)', [b,b]),
      'Cell style');
  end else
  begin  // BIFF3 - BIFF 8
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrROw, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Index of XF record');
  end;

  // boolean value
  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBufferIndex, numbytes,
    Format('%d (%s)', [b, Uppercase(BoolToStr(Boolean(b), true))]),
    'Boolean value (0=FALSE, 1=TRUE)'
  );

  // bool/error flag
  numBytes := 1;
  b := FBuffer[FBufferIndex];
  if b = 0 then
    ShowInRow(FCurrRow, FBufferIndex, numbytes, '0 (boolean value)',
      'Boolean/Error value flag (0=boolean, 1=error value)')
  else
    ShowInRow(FCurrRow, FBufferIndex, numbytes, '1 (error value)',
      'Boolean/Error value flag (0=boolean, 1=error value)');
end;

procedure TBIFFGrid.ShowBottomMargin;
var
  numBytes: Integer;
  dbl: Double;
begin
  RowCount := FixedRows + 1;
  numBytes := 8;
  Move(FBuffer[FBufferIndex], dbl, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl),
    'Bottom page margin in inches (IEEE 754 floating-point value, 64-bit double precision)');
end;


procedure TBIFFGrid.ShowCalcCount;
var
  numBytes: Word;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Maximum number of iterations allowed in circular references');
end;


procedure TBIFFGrid.ShowCalcMode;
var
  numBytes: Word;
  w: word;
  s: String;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if w = $FFFF then
    s := '–1 = automatically except for multiple table operations'
  else if w = 0 then
    s := '0 = manually'
  else if w = 1 then
    s := '1 = automatically (default)';
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]), s);
end;


procedure TBIFFGrid.ShowCellAddress;
{ Note: The bitmask assignment to relative column/row is reversed in relation
  to OpenOffice documentation in order to match with Excel files. }
var
  numBytes: Word;
  b: Byte;
  w: Word;
  r,c: Integer;
  s: String;
begin
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);   // row --> w
  r := WordLEToN(w);
  if FFormat = sfExcel8 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex+2], w, numBytes);  // column --w1
    c := WordLEToN(w);
    if FCurrRow = Row then begin
      FDetails.Add('RowIndex information:'#13);
      FDetails.Add(Format('RowIndex = %d (%s)', [r, ABS_REL[c and $4000 <> 0]]));
    end;
    s := Format('%d ($%.4x)', [r, r]);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'Row index');
    if FCurrRow = Row then begin
      FDetails.Add('ColIndex information:'#13);
      FDetails.Add(Format('Bits 0-13: ColIndex = %d (%s)', [c and $3FFF, ABS_REL[c and $8000 <> 0]]));
      if c and $4000 = 0
        then FDetails.Add('Bit 14=0: absolute column index')
        else FDetails.Add('Bit 14=1: relative column index');
      if c and $8000 = 0
        then FDetails.Add('Bit 15=0: absolute row index')
        else FDetails.Add('Bit 15=1: relative row index');
    end;
    s := Format('%d ($%.4x)', [c, c]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Column index');
  end else
  begin
    numbytes := 1;
    Move(FBuffer[FBufferIndex+2], b, numBytes);
    c := b;
    if FCurrRow = Row then begin
      FDetails.Add('RowIndex information:'#13);
      FDetails.Add(Format('Bits 0-13: RowIndex = %d (%s)', [r and $3FFF, ABS_REL[r and $4000 <> 0]]));
      if r and $4000 = 0
        then FDetails.Add('Bit 14=0: absolute column index')
        else FDetails.Add('Bit 14=1: relative column index');
      if r and $8000 = 0
        then FDetails.Add('Bit 15=0: absolute row index')
        else FDetails.Add('Bit 15=1: relative row index');
    end;
    //s := Format('$%.4x (%d, %s)', [r, r and $3FFF, ABS_REL[r and $4000 <> 0]]);
    s := Format('%d ($%.4x)', [r, r]);
    ShowInRow(FCurrRow, FBufferIndex, 2, s, 'Row index');
    if FCurrRow = Row then begin
      FDetails.Add('ColIndex information:'#13);
      FDetails.Add(Format('ColIndex = %d (%s)', [c, ABS_REL[r and $8000 <> 0]]));
    end;
    //s := Format('$%.2x (%d, %s)', [c, c, ABS_REL[r and $8000 <> 0]]);
    s := Format('%d ($%.4x)', [c, c]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Column index');
  end;
end;


procedure TBIFFGrid.ShowCellAddressRange(AFormat: TsSpreadsheetFormat);
{ Note: The bitmask assignment to relative column/row is reversed in relation
  to OpenOffice documentation in order to match with Excel files.

  The spreadsheet format is passed as a parameter because some BIFF8 records
  used these fields in BIFF5 format. }
var
  numbytes: Word;
  b: Byte;
  w: Word;
  r, c, r2, c2: Integer;
  s: String;
begin
  if AFormat = sfExcel8 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    r := WordLEToN(w);
    if FCurrRow = Row then begin
      FDetails.Add('RowIndex information:'#13);
      FDetails.Add(Format('RowIndex = %d (%s)', [r, ABS_REL[c and $4000 <> 0]]));
    end;
    s := Format('%d ($%.4x)', [r, r]);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'First row index');

    Move(FBuffer[FBufferIndex], w, numBytes);
    r2 := WordLEToN(w);
    if FCurrRow = Row then begin
      FDetails.Add('RowIndex information:'#13);
      FDetails.Add(Format('RowIndex = %d (%s)', [r2, ABS_REL[c and $4000 <> 0]]));
    end;
    s := Format('%d ($%.4x)', [r2, r2]);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'Last row index');

    Move(FBuffer[FBufferIndex], w, numBytes);  // column
    c := WordLEToN(w);
    Move(FBuffer[FBufferIndex+2], w, numBytes);
    c2 := WordLEToN(w);

    if FCurrRow = Row then begin
      FDetails.Add('ColIndex information:'#13);
      FDetails.Add(Format('Bits 0-13: ColIndex = %d (%s)', [c and $3FFF, ABS_REL[c and $8000 <> 0]]));
      if c and $4000 = 0
        then FDetails.Add('Bit 14=0: absolute column index')
        else FDetails.Add('Bit 14=1: relative column index');
      if c and $8000 = 0
        then FDetails.Add('Bit 15=0: absolute row index')
        else FDetails.Add('Bit 15=1: relative row index');
    end;
    s := Format('%d ($%.4x)', [c, c]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'First column index');

    if FCurrRow = Row then
    begin
      FDetails.Add('ColIndex information:'#13);
      FDetails.Add(Format('Bits 0-13: ColIndex = %d (%s)', [c2 and $3FFF, ABS_REL[c2 and $8000 <> 0]]));
      if c2 and $4000 = 0
        then FDetails.Add('Bit 14=0: absolute column index')
        else FDetails.Add('Bit 14=1: relative column index');
      if c2 and $8000 = 0
        then FDetails.Add('Bit 15=0: absolute row index')
        else FDetails.Add('Bit 15=1: relative row index');
    end;
    s := Format('%d ($%.4x)', [c2, c2]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Last column index');
  end
  else
  begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    r := WordLEToN(w);
    Move(FBuffer[FBufferIndex+2], w, numBytes);
    r2 := WordLEToN(w);

    numbytes := 1;
    c := FBuffer[FBufferIndex+4];
    c2 := FBuffer[FBufferIndex+5];

    if FCurrRow = Row then
    begin
      FDetails.Add('RowIndex information:'#13);
      FDetails.Add(Format('Bits 0-13: RowIndex = %d (%s)', [r and $3FFF, ABS_REL[r and $4000 <> 0]]));
      if r and $4000 = 0
        then FDetails.Add('Bit 14=0: absolute column index')
        else FDetails.Add('Bit 14=1: relative column index');
      if r and $8000 = 0
        then FDetails.Add('Bit 15=0: absolute row index')
        else FDetails.Add('Bit 15=1: relative row index');
    end;
    s := Format('%d ($%.4x)', [r, r]);
    ShowInRow(FCurrRow, FBufferIndex, 2, s, 'First row index');

    if FCurrRow = Row then
    begin
      FDetails.Add('RowIndex information:'#13);
      FDetails.Add(Format('Bits 0-13: RowIndex = %d (%s)', [r2 and $3FFF, ABS_REL[r2 and $4000 <> 0]]));
      if r2 and $4000 = 0
        then FDetails.Add('Bit 14=0: absolute column index')
        else FDetails.Add('Bit 14=1: relative column index');
      if r2 and $8000 = 0
        then FDetails.Add('Bit 15=0: absolute row index')
        else FDetails.Add('Bit 15=1: relative row index');
    end;
    s := Format('%d ($%.4x)', [r2, r2]);
    ShowInRow(FCurrRow, FBufferIndex, 2, s, 'Last row index');

    if FCurrRow = Row then
    begin
      FDetails.Add('ColIndex information:'#13);
      FDetails.Add(Format('ColIndex = %d (%s)', [c, ABS_REL[r and $8000 <> 0]]));
    end;
    s := Format('%d ($%.4x)', [c, c]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'First column index');

    if FCurrRow = Row then
    begin
      FDetails.Add('ColIndex information:'#13);
      FDetails.Add(Format('ColIndex = %d (%s)', [c2, ABS_REL[r2 and $8000 <> 0]]));
    end;
    s := Format('%d ($%.4x)', [c2, c2]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Last column index');
  end;
end;


procedure TBIFFGrid.ShowClrtClient;
var
  w: Word;
  dw: DWord;
  numbytes: Word;
begin
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);

  RowCount := FixedRows + w + 1;

  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Number of colors (must be 3)');

  numBytes := 4;
  Move(FBuffer[FBufferIndex], dw, numbytes);
  dw := DWordLEToN(dw);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [dw]),
    'Foreground color (system window text color)');
  Move(FBuffer[FBufferIndex], dw, numbytes);
  dw := DWordLEToN(dw);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [dw]),
    'Background color (system window color)');
  Move(FBuffer[FBufferIndex], dw, numbytes);
  dw := DWordLEToN(dw);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [dw]), '???');
end;


procedure TBIFFGrid.ShowCodePage;
var
  numBytes: Word;
  w: Word;
  s: String;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  s := CodePageName(w);
  if Row = FCurrRow then begin
    FDetails.Add('Code page:'#13);
    FDetails.Add(Format('$%.04x = %s', [w, s]));
  end;
  if s <> '' then s := 'Code page identifier (' + s + ')' else s := 'Code page identifier';
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]), s);
end;


procedure TBIFFGrid.ShowColInfo;
var
  numBytes: Integer;
  w: Word;
begin
  if FFormat = sfExcel2 then
    exit;

  RowCount := FixedRows + 5;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index of first column in range');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index of last column in range');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d (%f characters)', [w, w/256]),
    'Width of the columns in 1/256 of the width of the zero character, using default font (first FONT record in the file)');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to XF record for default column formatting');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Column options:'#13);
    if w and $0001 = 0
      then FDetails.Add('Bit  $0001 = 0: Columns are NOT hidden')
      else FDetails.Add('Bit  $0001 = 1: Columns are hidden');
    FDetails.Add(Format('Bits $0700 = %d: Outline level of the columns (0 = no outline)', [(w and $0700) shr 8]));
    if w and $1000 = 0
      then FDetails.Add('Bit  $1000 = 0: Columns are NOT collapsed')
      else FDetails.Add('Bit  $1000 = 1: Columns are collapsed');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w), 'Option flags');
end;


procedure TBIFFGrid.ShowColWidth;
var
  numBytes: Integer;
  w: Word;
  b: Byte;
begin
  if FFormat <> sfExcel2 then
    exit;

  RowCount := FixedRows + 3;

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
    'Index of first column');

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
    'Index of last column');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Width of the columns in 1/256 of the width of the zero character, using default font (first FONT record in the file)');
end;


procedure TBIFFGrid.ShowContinue;
var
  numbytes: Integer;
  s: String;
  sa: ansistring;
  sw: widestring;
  ls: Integer;
  i: Integer;
  w: Word;
  n: Integer;
  run: Integer;
begin
  case FInfo of
    BIFFNODE_TXO_CONTINUE1:
      begin
        RowCount := FixedRows + 1;
        numbytes := Length(FBuffer);
        if FBuffer[FBufferIndex] = 0 then begin
          ls := Length(FBuffer)-1;
          SetLength(sa, ls);
          Move(FBuffer[FBufferIndex+1], sa[1], ls);
          s := AnsiToUTF8(sa);
        end else
        if FBuffer[FBufferIndex] = 1 then begin
          ls := (Length(FBuffer) - 1) div SizeOf(WideChar);
          SetLength(sw, ls);
          Move(FBuffer[FBufferIndex+1], sw[1], ls*SizeOf(WideChar));
          s := UTF8Encode(sw);
        end else
          s := 'ERROR!!!';
        s := UTF8StringReplace(s, #$0A, '[/n]', [rfReplaceAll]);
        ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'Comment text');
      end;

    BIFFNODE_TXO_CONTINUE2:
      begin
        RowCount := FixedRows + 1000;
        n := 0;
        numBytes := 2;
        run := 1;
        while FBufferIndex < Length(FBuffer) - 4*SizeOf(Word) do begin
          Move(FBuffer[FBufferIndex], w, numBytes);
          ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)), Format(
            'Run %d: Index of first character using this font (0-based)', [run]));
          inc(n);

          Move(FBuffer[FBufferIndex], w, numbytes);
          ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToSTr(WordLEToN(w)),
            Format('Run %d: Index to FONT record', [run]));
          inc(n);

          Move(FBuffer[FBufferIndex], w, numbytes);
          ShowInrow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
            'Not used');
          inc(n);

          Move(FBuffer[FBufferIndex], w, numbytes);
          ShowInrow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
            'Not used');
          inc(n);

          inc(run);
        end;

        // lastRun
        Move(FBuffer[FBufferIndex], w, numbytes);
        ShowInrow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
          'Number of characters');
        inc(n);

        Move(FBuffer[FBufferIndex], w, numbytes);
        ShowInrow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
          'Not used');
        inc(n);

        Move(FBuffer[FBufferIndex], w, numbytes);
        ShowInrow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
          'Not used');
        inc(n);

        Move(FBuffer[FBufferIndex], w, numbytes);
        ShowInrow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
          'Not used');
        inc(n);

        RowCount := FixedRows + n;
      end;
  end;
end;


procedure TBIFFGrid.ShowCountry;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 2;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Windows country identifier for UI language of Excel');

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Windows country identifier of system regional settings');
end;


procedure TBIFFGrid.ShowDateMode;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    '0 = Base date is 1899-Dec-31, 1 = Base date is 1904-Jan-01');
end;

procedure TBIFFGrid.ShowDBCell;
var
  i, n: Integer;
  dw: DWord;
  w: Word;
  numBytes: Integer;
begin
  if FFormat < sfExcel5 then exit;

  n := (Length(FBuffer) - 4) div 2;
  RowCount := FixedRows + 1 + n;

  numBytes := 4;
  Move(FBuffer[FBufferIndex], dw, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
    'Relative offset to first ROW record in the Row Block');

  numBytes := 2;
  for i:=1 to n do begin
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Relative offsets to calculate stream position of the first cell record in row');
  end;
end;

procedure TBIFFGrid.ShowDefColWidth;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Column width in characters, using the width of the zero character from default '+
    'font (first FONT record in the file) + some extra space.');
end;


procedure TBIFFGrid.ShowDefinedName;
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  isFuncMacro: Boolean;
  lenName: Word;
  ansistr: AnsiString;
  widestr: WideString;
  s: String;
  macro: Boolean;
  formulaSize: Word;
  firstTokenBufIdx: Integer;
  token: Byte;
  r,c, r2,c2: Integer;
begin
  BeginUpdate;
  RowCount := FixedRows + 1000;
  // Brute force simplification because of unknown row count at this point
  // Will be reduced at the end.

  if FFormat = sfExcel2 then
  begin
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    isFuncMacro := b and $02 <> 0;
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:'#13);
      if b and $02 = 0 then
        FDetails.Add('Bit $02 = 0: NO function macro or command macro')
      else
        FDetails.Add('Bit $02 = 1: Function macro or command macro');
      if b and $04 = 0 then
        FDetails.Add('Bit $04 = 0: NO Complex function (array formula or user defined)')
      else
        FDetails.Add('Bit $04 = 1: Complex function (array formula or user defined)');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
      'Option flags');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    if isFuncMacro then
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
        '$01 = Function macro, $02 = Command macro')
    else
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
        'unknown');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Keyboard shortcut (only for command macro names)');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
      'Length of the name (character count)');
    lenName := b;

    numbytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(b),
      'Size of the formula data');
    formulaSize := b;
  end
  else
  begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      macro := (w and $0008 <> 0);
      FDetails.Add('Option flags:'#13);
      if w and $0001 = 0
        then FDetails.Add('Bit $0001 (flag name "hidden") = 0: Visible')
        else FDetails.Add('Bit $0001 (flag name "hidden")= 1: Hidden');
      if w and $0002 = 0
        then FDetails.Add('Bit §0002 (flag name "func") = 0: Command macro')
        else FDetails.Add('Bit $0002 (flag name "func") = 1: Function macro');
      if w and $0004 = 0
        then FDetails.Add('Bit $0004 (flag name "vbasic") = 0: Sheet macro')
        else FDetails.Add('Bit $0004 (flag name "vbasic") = 1: Visual basic macro');
      if w and $0008 = 0
        then FDetails.Add('Bit $0008 (flag name "macro") = 0: Standard name')
        else FDetails.Add('Bit $0008 (flag name "macro") = 1: Macro name');
      if w and $0010 = 0
        then FDetails.Add('Bit $0010 (flag name "complex") = 0: Simple formula')
        else FDetails.Add('Bit $0010 (flag name "complex") = 1: Complex formula (array formula or user defined)');
      if w and $0020 = 0
        then FDetails.Add('Bit $0020 (flag name "builtin") = 0: User-defined name')
        else FDetails.Add('Bit $0020 (flag name "builtin") = 1: Built-in name');
      case (w and $0FC0) shr 6 of
        0: if macro then
             FDetails.Add('Bit $0FC0 = 0: --- forbidden value, must be > 0 ---')
           else
             FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 0: not used (requires "macro" = 1)');
        1: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 1: financial');
        2: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 2: date & time');
        3: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 3: math & trig');
        4: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 4: statistical');
        5: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 5: lookup & reference');
        6: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 6: database');
        7: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 7: text');
        8: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 8: logical');
        9: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 9: information');
       10: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 10: commands');
       11: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 11: customizing');
       12: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 12: macro control');
       13: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 13: dde/external');
       14: FDetails.Add('Bit $0FC0 (flag name "funcgroup") = 14: user defined');
      end;
      if w and $1000 = 0
        then FDetails.Add('Bit $1000 (flag name "binary") = 0: formula definition')
        else FDetails.add('Bit $1000 (flag name "binary") = 1: binary data (BIFF5-BIFF8)');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Option flags');

    numbytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInrow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b, b]),
      'Keyboard shortcurt (only for command macro names)');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      'Length of the name (character count)');
    lenName := b;

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
      'Size of the formula data');
    formulaSize := w;

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if FFormat = sfExcel5 then
      ShowInRow(FCurrRow, FBufferIndex, NumBytes, IntToStr(w),
        '0 = Global name, otherwise index to EXTERNSHEET record (one-based)')
    else
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
        'not used');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, NumBytes, IntToStr(w),
      '0 = Global name, otherwise index to sheet (one-based)');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, nuMbytes, IntToStr(b),
      'Length of the menu text (character count)');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, nuMbytes, IntToStr(b),
      'Length of the description text (character count)');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, nuMbytes, IntToStr(b),
      'Length of the help topic text (character count)');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, nuMbytes, IntToStr(b),
      'Length of the status bar text (character count)');

    if FFormat = sfExcel5 then begin
      numBytes := lenName * sizeOf(ansiChar);
      SetLength(ansiStr, lenName);
      Move(FBuffer[FBufferIndex], ansiStr[1], numbytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, ansiStr,
        'Character array of the name');
    end else begin

      if (FBuffer[FBufferIndex] and $01 = 0) //and (not IgnoreCompressedFlag)
      then begin   // compressed --> 1 byte per character
        SetLength(ansiStr, lenName);
        numbytes := lenName*SizeOf(ansiChar) + 1;
        Move(FBuffer[FBufferIndex + 1], ansiStr[1], lenName*SizeOf(AnsiChar));
        s := AnsiToUTF8(ansiStr);
      end else begin
        SetLength(wideStr, lenName);
        numBytes := lenName*SizeOf(WideChar) + 1;
        Move(FBuffer[FBufferIndex + 1], wideStr[1], lenName*SizeOf(WideChar));
        s := UTF8Encode(WideStringLEToN(wideStr));
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
        'Name (Unicode string without length field)');
    end;
  end;

  firstTokenBufIdx := FBufferIndex;
  while FBufferIndex < firstTokenBufIdx + formulaSize do begin
    token := FBuffer[FBufferIndex];
    numBytes := 1;
    case token of
      $3A, $3B, $5A, $5B, $7A, $7B:
        begin
          case token of
            $3A: s := 'Token tRef3dR for "3D or external reference to a cell" (R = Reference)';
            $5A: s := 'Token tRef3dV for "3D or external reference to a cell" (V = Value)';
            $7A: s := 'Token tRef3dA for "3D or external reference to a cell" (A = Area)';
            $3B: s := 'Token tArea3dR for "3D or external reference to a cell range" (R = Reference)';
            $5B: s := 'Token tArea3dV for "3D or external reference to a cell range" (V = Value)';
            $7B: s := 'Token tArea3dA for "3D or external reference to a cell range" (A = Area)';
          end;
          ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [token]), s);

          numbytes := 2;
          Move(FBuffer[FBufferIndex], w, numBytes);
          w := WordLEToN(w);
          if FFormat = sfExcel5 then begin
            if w and $8000 <> 0 then begin  // negative value --> 3D reference
              ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(SmallInt(w)),
                '3D reference, 1-based index to EXTERNSHEET record');
              numBytes := 8;
              ShowInRow(FCurrRow, FBufferIndex, numBytes, '', 'Not used');
              numBytes := 2;
              Move(FBuffer[FBufferIndex], w, numBytes);
              ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
                'Index to first referenced sheet ($FFFF = deleted sheet)');
              Move(FBuffer[FBufferIndex], w, numBytes);
              ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
                'Index to last referenced sheet ($FFFF = deleted sheet)');
            end else
            begin
              ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
                'External reference, 1-based index to EXTERNSHEET record');
              numBytes := 12;
              ShowInRow(FCurrRow, FBufferIndex, numBytes, '', 'Not used');
            end;
          end else
            ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
              'Index to REF entry in EXTERNSHEET record');

          if token in [$3A, $5A, $7A] then
            ShowCellAddress        // Cell address
          else
            ShowCellAddressRange(FFormat);  // Cell range
        end;

      else
        numBytes := 1;
        ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
          '(unknown token)');
    end;  // case
  end;  // while

  RowCount := FCurrRow;
  EndUpdate(true);
end;

procedure TBIFFGrid.ShowDefRowHeight;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + IfThen(FFormat = sfExcel2, 1, 2);

  if FFormat = sfExcel2 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Default height for unused rows:'#13);
      FDetails.Add(Format(
        'Bits $7FFF = %d: Default height for unused rows, in twips = 1/20 of a point',
        [w and $7FFF]));
      if w and $8000 = 0 then
        FDetails.Add('Bit $8000 = 0: Row height changed manually')
      else
        FDetails.Add('Bit $8000 = 1: Row height not changed manually');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Default height for unused rows, in twips = 1/20 of a point');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:'#13);
      if w and $0001 = 0
        then FDetails.Add('Bit $0001 = 0: Row height and default font height do match')
        else FDetails.Add('Bit $0001 = 1: Row height and default font height do not match');
      if w and $0002 = 0
        then FDetails.Add('Bit $0002 = 0: Row is visible')
        else FDetails.Add('Bit $0002 = 1: Row is hidden');
      if w and $0004 = 0
        then FDetails.Add('Bit $0004 = 0: No additional space above the row')
        else FDetails.Add('Bit $0004 = 1: Additional space above the row');
      if w and $0008 = 0
        then FDetails.Add('Bit $0008 = 0: No additional space below the row')
        else FDetails.Add('Bit $0008 = 1: Additional space below the row');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [WordLEToN(w)]),
      'Option flags');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'Default height for unused rows, in twips = 1/20 of a point');
  end;
end;


procedure TBIFFGrid.ShowDelta;
var
  numBytes: Integer;
  dbl: Double;
begin
  RowCount := FixedRows + 1;
  numBytes := 8;
  Move(FBuffer[FBufferIndex], dbl, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl),
    'Maximum change in iteration (IEEE 754 floating-point value, 64-bit double precision)');
end;


procedure TBIFFGrid.ShowDimensions;
var
  numBytes: Integer;
  dw: DWord;
  w: Word;
begin
  RowCount := FixedRows + IfThen(FFormat = sfExcel2, 4, 5);

  if FFormat = sfExcel8 then begin
    numBytes := 4;
    Move(FBuffer[FBufferIndex], dw, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(DWordLEToN(dw)),
      'Index to first used row');

    Move(FBuffer[FBufferIndex], dw, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(DWordLEToN(dw)),
      'Index to last used row, increased by 1');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'Index to first used row');

    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'Index to last used row, increased by 1');
  end;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to first used column');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to last used column, increased by 1');

  if FFormat <> sfExcel2 then begin
    numBytes := 2;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, '',
      '(not used)');
  end;
end;


procedure TBIFFGrid.ShowDSF;
var
  w: Word;
  numbytes: Integer;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w), 'Reserved, MUST be ignored');
end;


procedure TBIFFGrid.ShowEOF;
begin
  RowCount := FixedRows + 1;
  ShowInRow(FCurrRow, FBufferIndex, 0, '', '(no content)');
end;


procedure TBIFFGrid.ShowExcel9File;
begin
  RowCount := FixedRows + 1;
  ShowInRow(FCurrRow, FBufferIndex, 0, '', 'Optional and unused');
end;


procedure TBIFFGrid.ShowExternalBook;
var
  numBytes: Integer;
  w: Word;
  wideStr: WideString;
  ansiStr: AnsiString;
  s: String;
  i, n: Integer;
begin
  BeginUpdate;
  RowCount := FixedRows + 1000;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  n := WordLEToN(w);

  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(n),
    'Number of sheet names / number of sheets');

  if Length(FBuffer) - FBufferIndex = 2 then begin
    SetLength(ansiStr, 1);
    Move(FBuffer[FBufferIndex], ansiStr[1], 1);

    ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
      '(relict of BIFF5)');
  end else begin
    ExtractString(FBufferIndex, 2, true, s, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Encoded URL without sheet name:'#13);
      case s[1] of
        #0: FDetails.Add('First character = #00: Reference relative to current sheet');
        #1: FDetails.Add('First character = #01: Encoded URL follows');
        #2: if FFormat = sfExcel8 then
              FDetails.Add('First character = #02: Reference to a sheet in own document; sheet name follows')
            else
              FDetails.Add('First character = #02: Reference to the corrent sheet (nothing will follow)');
        #3: if FFormat = sfExcel5 then
              FDetails.Add('First character = #03: Reference to a sheet in own document; sheet name follows')
            else
              FDetails.Add('First character = #03: not used');
        #4: if FFormat = sfExcel5 then
              FDetails.ADd('First character = #03: Reference to the own workbook, sheet is unspecified (nothing will follow)')
            else
              FDetails.Add('First character = #03: not used');
      end;
    end;
    if s[1] in [#0, #1, #2, #3, #4] then Delete(s, 1, 1);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
      'Encoded URL without sheet name (Unicode string, 16-bit string length)');

    for i:=0 to n-1 do begin
      ExtractString(FBufferIndex, 2, true, s, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
        'Sheet name (Unicode string with 16-bit string length)');
    end;
  end;

  RowCount := FCurrRow;
  EndUpdate(true);
end;


procedure TBIFFGrid.ShowExternCount;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Number of following EXTERNSHEET records');
end;


procedure TBIFFGrid.ShowExternSheet;
var
  numBytes: Integer;
  w: Word;
  s: String;
  nREF: Integer;
  i: Integer;
  len: Byte;
  ansiStr: AnsiString;
begin
  if FFormat <= sfExcel5 then begin
    RowCount := FixedRows + 1;
    len := FBuffer[0];
    if FBuffer[1] = $03 then inc(len);
    numBytes := len*SizeOf(AnsiChar) + 1;
    SetLength(ansiStr, len);
    Move(FBuffer[1], ansiStr[1], len*SizeOf(AnsiChar));
    s := AnsiToUTF8(ansiStr);
    if FCurrRow = Row then begin
      FDetails.Add('Encoded document and sheet name:'#13);
      if s[1] = #03 then begin
        FDetails.Add('First character = $03: EXTERNSHEET stores a reference to one of the own sheets');
        FDetails.Add('Document name: ' + Copy(s, 2, Length(s)));
      end else
      if (s[1] = ':') and (Length(s) = 1) then begin
        FDetails.Add('Special EXTERNSHEET record for an add-in function. EXTERNName record with the name of the function follows.');
      end else
        FDetails.Add('Document name: ' + s);
    end;
    if s[1] = #03 then
      Delete(s, 1, 1);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
      'Encoded document and sheet name (Byte string, 8-bit string length)');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    nREF := WordLEToN(w);

    RowCount := FixedRows + 1 + nREF*3;

    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(nREF),
      'Number of following REF structures');

    for i:=1 to nREF do begin
      Move(FBuffer[FBufferIndex], w, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
        Format('REF #%d: Index to EXTERNALBOOK record', [i]));

      Move(FBuffer[FBufferIndex], w, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
        Format('REF #%d: Index to first sheet in EXTERNALBOOK sheet list', [i]));

      Move(FBuffer[FBufferIndex], w, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
        Format('REF #%d: Index to last sheet in EXTERNALBOOK sheet list', [i]));
    end;
  end;
end;


procedure TBIFFGrid.ShowFileSharing;
var
  numbytes: Integer;
  w: Word;
  s: String;
begin
  RowCount := FixedRows + 3;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Recommend read-only state when loading the file:'#13);
    if w = 0 then FDetails.Add('0 = no') else FDetails.Add('1 = yes');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Recommend read-only state when loading the file');

  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    'Hash value calculated from the read-only password');

  ExtractString(FBufferIndex, IfThen(FFormat=sfExcel8, 2, 1), FFormat=sfExcel8,
    s, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
    'User name of the file creator' + IfThen(FFormat = sfExcel8,
    ' (Unicode string, 16-bit string length)',
    ' (byte string, 8-bit string length)'
  ));
end;


procedure TBIFFGrid.ShowFnGroupCount;
var
  numbytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Number of built-in function categories:'#13);
    case w of
      $000E:
        FDetails.Add(
          'There are 14 built-in function categories in the workbook.'#13+
          'This implies that the file was last saved by a specific version of the application.'#13+
          'The following 9 built-in function categories are visible to the end-user:'#13+
          '  Financial'#13+
          '  Date & Time'#13+
          '  Math & Trig'#13+
          '  Statistical'#13+
          '  Lookup & Reference'#13+
          '  Database'#13+
          '  Text'#13+
          '  Logical'#13+
          '  Information'#13+
          'The following 5 built-in function categories are not visible to the end-user:'#13+
          '  UserDefined'#13+
          '  Commands'#13+
          '  Customize'#13+
          '  MacroControl'#13+
          '  DDEExternal'
        );
      $0010:
        FDetails.Add(
          'There are 16 built-in function categories in the workbook.'#13+
          'This implies that the file was last saved by a specific version of the application'#13+
          'The following 11 built-in function categories are visible to the end-user:'#13+
          '  Financial'#13+
          '  Date & time'#13+
          '  Math & Trig'#13+
          '  Statistical'#13+
          '  Lookup & Reference'+
          '  Database'#13+
          '  Text'#13+
          '  Logical'#13+
          '  Information'#13+
          '  Engineering'#13+
          '  Cube'#13+
          'The following 5 built-in function categories are not visible to the end-user:'#13+
          '  UserDefined'#13+
          '  Commands'#13+
          '  Customize'#13+
          '  MacroControl'#13+
          '  DDEExternal'
        );
    end;
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Number of built-in function categories');
end;


procedure TBIFFGrid.ShowFont;
var
  numbytes: Integer;
  w: Word;
  b: Byte;
  s: String;
begin
  RowCount := IfThen(FFormat = sfExcel2, 3, 10) + FixedRows;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d (= %.1gpt)', [w, w/20]),
    'Font height in twips (=1/20 point)');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Option flags:'#13);
    if w and $0001 = 0
      then FDetails.Add('  Bit $0001 = 0: not bold')
      else FDetails.Add('x Bit $0001 = 1: bold (redundant in BIFF5-BIFF8)');
    if w and $0002 = 0
      then FDetails.Add('  Bit $0002 = 0: not italic')
      else FDetails.Add('x Bit $0002 = 1: italic');
    if w and $0004 = 0
      then FDetails.Add('  Bit $0004 = 0: not underlined')
      else FDetails.Add('x Bit $0004 = 1: underlined (redundant in BIFF5-BIFF8)');
    if w and $0008 = 0
      then FDetails.Add('  Bit $0008 = 0: not struck out')
      else FDetails.Add('x Bit $0008 = 1: struck out');
    if w and $0010 = 0
      then FDetails.Add('  Bit $0010 = 0: not outlined')
      else FDetails.Add('x Bit $0010 = 1: outlined');
    if w and $0020 = 0
      then FDetails.Add('  Bit $0020 = 0: not shadowed')
      else FDetails.Add('x Bit $0020 = 1: shadowed');
    if w and $0040 = 0
      then FDetails.Add('  Bit $0040 = 0: not condensed')
      else FDetails.Add('x Bit $0040 = 1: condensed');
    if w and $0080 = 0
      then FDetails.Add('  Bit $0080 = 0: not extended')
      else FDetails.Add('x Bit $0080 = 1: extended');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
    'Option flags');

  if FFormat <> sfExcel2 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(2), 'Color index');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w,w]),
      'Font weight (400=normal, 700=bold)');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Escapement:'#13);
      case w of
        0: FDetails.Add('0 = none');
        1: FDetails.Add('1 = superscript');
        2: FDetails.Add('2 = subscript');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
      'Escapement ($00=none, $01=superscript, $02=subscript)');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Underline type:'#13);
      case b of
        $00: FDetails.Add('$00 = no underline');
        $01: FDetails.Add('$01 = single underline');
        $02: FDetails.Add('$02 = double underline');
        $21: FDetails.Add('$21 = single accounting');
        $22: FDetails.Add('$22 = double accounting');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
      'Underline type ($00=none, $01=single, $02=double, ...)');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Font family:'#13);
      case b of
        $00: FDetails.Add('$00 = None (unknown or don''t care)');
        $01: FDetails.Add('$01 = Roman (variable width, serifed)');
        $02: FDetails.Add('$02 = Swiss (variable width, sans-serifed)');
        $03: FDetails.Add('$03 = Modern (fixed width, serifed or sans-serifed)');
        $04: FDetails.Add('$04 = Script (cursive)');
        $05: FDetails.Add('$05 = Decorative (specialised, for example Old English, Fraktur)');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%.2x', [b]),
      'Font family');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    case b of
      $00: s := 'ANSI Latin';
      $01: s := 'System default';
      $02: s := 'Symbol';
      $4D: s := 'Apple Roman';
      $80: s := 'ANSI Japanese Shift-JIS';
      $81: s := 'ANSI Korean (Hangul)';
      $82: s := 'ANSI Korean (Johab)';
      $86: s := 'ANSI Chinese Simplified GBK';
      $88: s := 'ANSI Chinese Traditional BIG5';
      $A1: s := 'ANSI Greek';
      $A2: s := 'ANSI Turkish';
      $A3: s := 'ANSI Vietnamese';
      $B1: s := 'ANSI Hebrew';
      $B2: s := 'ANSI Arabic';
      $BA: s := 'ANSI Baltic';
      $CC: s := 'ANSI Cyrillic';
      $DE: s := 'ANSI Thai';
      $EE: s := 'ANSI Latin II (Central European)';      // East Europe in MS docs!
      $FF: s := 'OEM Latin I';
      else s := '';
    end;
    if s <> '' then s := Format('$%.2x: %s', [b, s]) else s := Format('$%.2x', [b]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
      'Character set');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]), 'Not used');
  end;

  ExtractString(FBufferIndex, 1, FFormat=sfExcel8, s, numbytes);
  if FFormat = sfExcel8 then
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'Font name (unicode string, 8-bit string length)')
  else
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'Font name (byte string, 8-bit string length)');
end;


procedure TBIFFGrid.ShowFontColor;
var
  numBytes: Integer;
  w: Word;
  s: String;
begin
  RowCount := FixedRows + 1;
  NumBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  case w of
    $0000: s := 'EGA Black (rgb = $000000)';
    $0001: s := 'EGA White (rgb = $FFFFFF)';
    $0002: s := 'EGA Red (rgb = $0000FF)';
    $0003: s := 'EGA Green (rgb = $00FF00)';
    $0004: s := 'EGA Blue (rgb = $FF0000)';
    $0005: s := 'EGA Yellow (rgb = $00FFFF)';
    $0006: s := 'EGA Magenta (rgb = $FF00FF)';
    $0007: s := 'EGA Cyan (rgb = $FFFF00)';
    $7FFF: s := 'Automatic (system window text colour)';
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.04x)', [w, w]),
    Format('Font color index into preceding FONT record (%s)', [s]));
end;

procedure TBIFFGrid.ShowFooter;
var
  numbytes: Integer;
  s: String;
begin
  RowCount := FixedRows + 1;
  ExtractString(FBufferIndex, IfThen(FFormat=sfExcel8, 2, 1), FFormat=sfExcel8,
    s, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
    'Page footer string' + IfThen(FFormat = sfExcel8,
    ' (Unicode string, 16-bit string length)',
    ' (byte string, 8-bit string length)'
  ));
end;


procedure TBIFFGrid.ShowFormat;
var
  numBytes: Integer;
  w: word;
  b: Byte;
  s: String;
begin
  RowCount := IfThen(FFormat = sfExcel2, FixedRows + 1, FixedRows + 2);
  if FFormat <> sfExcel2 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
      'FormatIndex used in other records');
  end;
  b := IfThen(FFormat=sfExcel8, 2, 1);
  ExtractString(FBufferIndex, b, (FFormat=sfExcel8), s, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
    Format('Number format string (%s string, %d-bit string length)', [GetStringType, b*8]));
end;


procedure TBIFFGrid.ShowFormatCount;
var
  numBytes: Integer;
  w: Word;
begin
  if FFormat = sfExcel2 then begin
    RowCount := 1 + FixedRows;
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Number of FORMAT records');
  end;
end;


procedure TBIFFGrid.ShowFormula;
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  q: QWord;
  dbl: double absolute q;
  bytearr: array[0..7] of byte absolute q;
  wordarr: array[0..3] of word absolute q;
//  s: String;
  tokenBytes: Integer;
//  firstTokenBufIdx: Integer;
//  token: Byte;
  r,c, r2,c2: Integer;
begin
  BeginUpdate;
  RowCount := FixedRows + 1000;
  // Brute force simplification because of unknown row count at this point
  // Will be reduced at the end.

  // Offset 0 = Row, Offset 2 = Column
  ShowRowColData(FBufferIndex);
  // Offset 4 = Cell attributes (BIFF2) or XF ecord index (> BIFF2)
  if FFormat = sfExcel2 then begin
    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell protection and XF index:'#13);
      FDetails.Add(Format('x Bits 5-0 = %d: XF Index', [b and $3F]));
      case b and $40 of
        0: FDetails.Add('  Bit 6 = 0: Cell is NOT locked.');
        1: FDetails.Add('x Bit 6 = 1: Cell is locked.');
      end;
      case b and $80 of
        0: FDetails.Add('  Bit 7 = 0: Formula is NOT hidden.');
        1: FDetails.Add('x Bit 7 = 1: Formula is hidden.');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Cell protection and XF index');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Indexes to format and font records:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: Index to FORMAT record', [b and $3f]));
      FDetails.Add(Format('Bits 7-6 = %d: Index to FONT record', [(b and $C0) shr 6]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Indexes of format and font records');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell style:'#13);
      case b and $07 of
        0: FDetails.Add('Bits 2-0 = 0: Horizontal alignment is GENERAL');
        1: FDetails.Add('Bits 2-0 = 1: Horizontal alignment is LEFT');
        2: FDetails.Add('Bits 2-0 = 2: Horizontal alignment is CENTERED');
        3: FDetails.Add('Bits 2-0 = 3: Horizontal alignment is RIGHT');
        4: FDetails.Add('Bits 2-0 = 4: Horizontal alignment is FILLED');
      end;
      if b and $08 = 0
        then FDetails.Add('Bit 3 = 0: Cell has NO left border')
        else FDetails.Add('Bit 3 = 1: Cell has left black border');
      if b and $10 = 0
        then FDetails.Add('Bit 4 = 0: Cell has NO right border')
        else FDetails.Add('Bit 4 = 1: Cell has right black border');
      if b and $20 = 0
        then FDetails.Add('Bit 5 = 0: Cell has NO top border')
        else FDetails.Add('Bit 5 = 1: Cell has top black border');
      if b and $40 = 0
        then FDetails.Add('Bit 6 = 0: Cell has NO bottom border')
        else FDetails.Add('Bit 6 = 1: Cell has bottom black border');
      if b and $80 = 0
        then FDetails.Add('Bit 7 = 0: Cell has NO shaded background')
        else FDetails.Add('Bit 7 = 1: Cell has shaded background');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.2x)', [b,b]),
      'Cell style');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrROw, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Index of XF record');
  end;

  // Offset 6: Result of formula
  numBytes := 8;
  Move(FBuffer[FBufferIndex], q, numBytes);
  if wordarr[3] <> $FFFF then begin
    if FCurrRow = Row then begin
      FDetails.Add('Formula result:'#13);
      FDetails.Add(Format('Bytes 0-7: $%.15x --> IEEE 764 floating-point value, 64-bit double precision'#13+
                          '           = %g', [q, dbl]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl),
      'Result of formula (IEEE 764 floating-point value, 64-bit double precision)');
  end else begin
    case bytearr[0] of
      0: begin     // String result
           if FCurrRow = Row then begin
             FDetails.Add('Formula result:'#13);
             FDetails.Add('Byte 0 = 0  --> Result is string, follows in STRING record');
             FDetails.Add('Byte 1-5: Not used');
             FDetails.Add('Byte 6&7: $FFFF --> no floating point number');
           end;
           ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.16x', [q]),
             'Result is a string, follows in STRING record');
         end;
      1: begin     // BOOL result
           if FCurrRow = Row then begin
             FDetails.Add('Formula result:'#13);
             FDetails.Add('Byte 0 = 1 --> Result is BOOLEAN');
             FDetails.Add('Byte 1: Not used');
             if bytearr[2] = 0
               then FDetails.Add('Byte 2 = 0 --> FALSE')
               else FDetails.Add('Byte 2 = 1 --> TRUE');
             FDetails.Add('Bytes 3-5: Not used');
             FDetails.Add('Bytes 6&7: $FFFF --> no floating point number');
           end;
           ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.16x', [q]),
             'Result is BOOLEAN');
         end;
      2: begin        // ERROR result
           if FCurrRow = Row then begin
             FDetails.Add('Formula result:'#13);
             FDetails.Add('Byte 0 = 2 --> Result is an ERROR value');
             FDetails.Add('Byte 1: Not used');
             case bytearr[2] of
               $00: FDetails.Add('Byte 2 = $00 --> #NULL! Intersection of two cell ranges is empty');
               $07: FDetails.Add('Byte 2 = $07 --> #DIV/0! Division by zero');
               $0F: FDetails.Add('Byte 2 = $0F --> #VALUE! Wrong type of operand');
               $17: FDetails.Add('Byte 2 = $17 --> #REF! Illegal or deleted cell reference');
               $1D: FDetails.Add('Byte 2 = $1D --> #NAME? Wrong function or range name');
               $24: FDetails.Add('Byte 2 = $24 --> #NUM! Value range overflow');
               $2A: FDetails.Add('Byte 2 = $2A --> #N/A Argument or function not available');
             end;
             FDetails.Add('Bytes 6&7: $FFFF --> no floating point number');
           end;
           ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.16x', [q]),
             'Result is an ERROR value');
         end;
      3: begin       // EMPTY cell
           if FCurrRow = Row then begin
             FDetails.Add('Formula result:'#13);
             FDetails.Add('Byte 0 = 3 --> Result is an empty cell, for example an empty string');
             FDetails.Add('Byte 1-5: Not used');
             FDetails.Add('Bytes 6&7: $FFFF --> no floating point number');
           end;
           ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.16x', [q]),
             'Result is an EMPTY cell (empty string)');
         end;
    end;
  end;

  // Option flags
  if FFormat = sfExcel2 then begin
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:'#13);
      case b of
        0: FDetails.Add('0 = Do not recalculate');
        1: FDetails.Add('1 = Recalculate always');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b), 'Option flags');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:'#13);
      if w and $0001 = 0
        then FDetails.Add('Bit $0001 = 0: Do not recalculate')
        else FDetails.Add('Bit $0001 = 1: Recalculate always');
      FDetails.Add('Bit $0002: Reserved - MUST be zero, MUST be ignored');
      if w and $0004 = 0
        then FDetails.Add('Bit $0004 = 0: Cell does NOT have a fill alignment or a center-across-selection alignment.')
        else FDetails.Add('Bit $0004 = 1: Cell has either a fill alignment or a center-across-selection alignment.');
      if w and $0008 = 0
        then FDetails.Add('Bit $0008 = 0: Formula is NOT part of a shared formula')
        else FDetails.Add('Bit $0008 = 1: Formula is part of a shared formula');
      FDetails.Add('Bit $0010: Reserved - MUST be zero, MUST be ignored');
      if w and $0020 = 0
        then FDetails.Add('Bit $0020 = 0: Formula is NOT excluded from formula error checking')
        else FDetails.Add('Bit $0020 = 1: Formula is excluded from formula error checking');
      FDetails.Add('Bits $FC00: Reserved - MUST be zero, MUST be ignored');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
      'Option flags');
  end;

  // Not used
  if (FFormat >= sfExcel5) then begin
    numBytes := 4;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, '', '(not used');
  end;

  // Size of Token array (in Bytes)
  if FFormat = sfExcel2 then begin
    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    tokenBytes := b;
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    tokenBytes := WordLEToN(w);
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(tokenBytes),
    'Size of formula data (in Bytes)');

  ShowFormulaTokens(tokenBytes);

  RowCount := FCurrRow;
  EndUpdate(true);
end;


procedure TBIFFGrid.ShowFormulaTokens(ATokenBytes: Integer);
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  s: String;
  dbl: Double;
  firstTokenBufIndex: Integer;
  token: Byte;
  r, c: Word;
begin
  // Tokens and parameters
  firstTokenBufIndex := FBufferIndex;
  while FBufferIndex < firstTokenBufIndex + ATokenBytes do begin
    token := FBuffer[FBufferIndex];
    numBytes := 1;
    case token of
      $01: begin
             ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [token]),
               'Token for "Cell is part of shared formula"');
             numbytes := 2;
             Move(FBuffer[FBufferIndex], w, numBytes);
             ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
               'Index to row of first FORMULA record in the formula range');
             if FFormat = sfExcel2 then begin
               numbytes := 1;
               b := FBuffer[FBufferIndex];
               ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
                 'Index to column of first FORMULA record in the formula range');
             end else begin
               numbytes := 2;
               Move(FBuffer[FBufferIndex], w, numbytes);
               ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
                 'Index to column of first FORMULA record in the formula range');
             end;
           end;
      $02: begin
             ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [token]),
               'Token for "Cell is part of a multiple operations table"');
             numbytes := 2;
             Move(FBuffer[FBufferIndex], w, numBytes);
             ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
               'Index to first row of the table range');
             if FFormat = sfExcel2 then begin
               numbytes := 1;
               b := FBuffer[FBufferIndex];
               ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
                 'Index to first column of the table range');
             end else begin
               numbytes := 2;
               Move(FBuffer[FBufferIndex], w, numbytes);
               ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
                 'Index to first column of the table range');
             end;
           end;
      $03: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "+" (add)');
      $04: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "-" (subtract)');
      $05: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "*" (multiply)');
      $06: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "/" (divide)');
      $07: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "^" (power)');
      $08: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "&" (concat)');
      $09: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "<" (less than)');
      $0A: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "<=" (less equal)');
      $0B: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "=" (equal)');
      $0C: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token ">=" (greater equal)');
      $0D: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token ">" (greater than)');
      $0E: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "<>" (not equal)');
      $0F: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token " " (intersect)');
      $10: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "list character"');
      $11: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token ":" (range)');
      $12: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "+" (unary plus)');
      $13: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "-" (unary minus)');
      $14: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "%" (percent)');
      $15: ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [token]),
             'Token "()" (operator in parenthesis)');
      $16: ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
             'Token "missing argument"');
      $17: begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tSTR (Label)');
             ExtractString(FBufferIndex, 1, (FFormat = sfExcel8), s, numBytes);
             ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'String value');
           end;
      $1C: begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tERR (Error)');
             numBytes := 1;
             b := FBuffer[FBufferIndex];
             if FCurrRow = Row then begin
               FDetails.Add('Error code:'#13);
               FDetails.Add(Format('Code $%.2x --> "%s"', [b, GetErrorValueStr(TsErrorValue(b))]));
             end;
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]), 'Error code');
           end;
      $1D: begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tBOOL');
             numBytes := 1;
             b := FBuffer[FBufferIndex];
             ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
               '0=FALSE, 1=TRUE');
           end;
      $1E: begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tINT (Integer)');
             numBytes := 2;
             Move(FBuffer[FBufferIndex], w, numBytes);
             ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
               'Integer value');
           end;
      $1F: begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tNUM (Number)');
             numBytes := 8;
             Move(FBuffer[FBufferIndex], dbl, numBytes);
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%g', [dbl]), //FloatToStr(dbl),
               'IEEE 754 floating-point value');
           end;
      $20, $40, $60:
           begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tARRAY');
             if FFormat = sfExcel2 then numBytes := 6 else numBytes := 7;
             ShowInRow(FCurrRow, FBufferIndex, numbytes, '', '(not used)');
           end;
      $21, $41, $61:
           begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tFUNC (Function with fixed argument count)');
             if FFormat = sfExcel2 then begin
               numBytes := 1;
               b := FBuffer[FBufferIndex];
               s := Format('Index of function (%s)', [SheetFuncName(b)]);
               ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), s);
             end else begin
               numBytes := 2;
               Move(FBuffer[FBufferIndex], w, numBytes);
               w := WordLEToN(w);
               s := Format('Index of function (%s)', [SheetFuncName(w)]);
               ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), s);
             end;
           end;
      $22, $42, $62:
           begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tFUNCVAR (Function with variable argument count)');
             numBytes := 1;
             b := FBuffer[FBufferIndex];
             ShowInRow(FCurrRow, FBufferIndex, numBytes,  IntToStr(b),
               'Number of arguments');
             if FFormat = sfExcel2 then begin
               numBytes := 1;
               s := Format('Index of built-in function (%s)', [SheetFuncName(b)]);
               ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), s);
             end else begin
               numBytes := 2;
               Move(FBuffer[FBufferIndex], w, numbytes);
               w := WordLEToN(w);
               s := Format('Index of built-in function (%s)', [SheetFuncName(w)]);
               ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), s);
             end;
           end;
      $23, $43, $63:
           begin
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               'Token tNAME');
             numBytes := 2;
             Move(FBuffer[FBufferIndex], w, numBytes);
             case FFormat of
               sfExcel2: s := 'DEFINEDNAME or EXTERNALNAME record';
               sfExcel5: s := 'DEFINEDNAME record in Global Link Table';
               sfExcel8: s := 'DEFINEDNAME record in Link Table';
             end;
             ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
               '1-based index to '+s);
             case FFormat of
               sfExcel2: numBytes := 5;
               sfExcel5: numBytes := 12;
               sfExcel8: numBytes := 2;
             end;
             ShowInRow(FCurrRow, FBufferIndex, numBytes, '', '(not used)');
           end;
      $24, $44, $64:
           begin
             case token of
               $24: s := 'reference';
               $44: s := 'value';
               $64: s := 'array';
             end;
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               Format('Token tREF (Cell %s)', [s]));
             ShowCellAddress;
           end;
      $25, $45, $65:
           begin
             case token of
               $25: s := 'reference';
               $45: s := 'value';
               $65: s := 'array';
             end;
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
               Format('Token tAREA (Cell range %s)', [s]));
             ShowCellAddressRange(FFormat);
           end;
      $2C, $4C, $6C:
          begin
            case token of
              $2C: s := 'reference';
              $4C: s := 'value';
              $6C: s := 'array';
            end;
            ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
              Format('Token tREFN (Relative reference to cell %s in same sheet)', [s]));

            // Encoded relative cell address
            if FFormat = sfExcel8 then
            begin
              numBytes := 2;
              Move(FBuffer[FBufferIndex], w, numBytes);    // row
              r := WordLEToN(w);
              Move(FBuffer[FBufferIndex+2], w, numBytes);  // column with flags
              c := WordLEToN(w);

              { Note: The bitmask assignment to relative column/row is reversed in relation
                to OpenOffice documentation in order to match with Excel files. }
              if Row = FCurrRow then begin
                FDetails.Add('Encoded cell address (row):'#13);
                if c and $8000 = 0 then
                begin
                  FDetails.Add('Row index is ABSOLUTE (see Encoded column index)');
                  FDetails.Add('Absolute row index: ' + IntToStr(r));
                end else
                begin
                  FDetails.Add('Row index is RELATIVE (see Encoded column index)');
                  FDetails.Add('Relative row index: ' + IntToStr(Smallint(r)));
                end;
              end;
              ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [r, r]),
                'Encoded row index');

              // Bit mask $4000 --> column
              // Bit mask $8000 --> row
              if Row = FCurrRow then
              begin
                FDetails.Add('Encoded cell address (column):'#13);
                if c and $4000 = 0
                  then FDetails.Add('Bit 14=0: Column index is ABSOLUTE')
                  else FDetails.Add('Bit 14=1: Column index is RELATIVE');
                if c and $8000 = 0
                  then FDetails.Add('Bit 15=0: Row index is ABSOLUTE')
                  else FDetails.Add('Bit 15=1: Row index is RELATIVE');
                FDetails.Add('');
                if c and $4000 = 0
                  then FDetails.Add('Absolute column index: ' + IntToStr(Lo(c)))
                  else FDetails.Add('Relative column index: ' + IntToStr(ShortInt(Lo(c))));
              end;

              ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [c, c]),
                'Encoded column index');
            end
            else
            // Excel5 (Excel2 does not support shared formulas)
            begin
              numBytes := 2;
              Move(FBuffer[FBufferIndex], w, numBytes);
              w := WordLEToN(w);
              r := w and $3FFF;
              b := FBuffer[FBufferIndex+2];
              c := b;

              // Bit mask $4000 --> column
              // Bit mask $8000 --> row
              if Row = FCurrRow then begin
                FDetails.Add('Encoded cell address (row):'#13);
                if w and $4000 = 0
                  then FDetails.Add('Bit 14=0: Column index is ABSOLUTE')
                  else FDetails.Add('Bit 14=0: Column index is RELATIVE');
                if w and $8000 = 0
                  then FDetails.Add('Bit 15=0: Row index is ABSOLUTE')
                  else FDetails.Add('Bit 15=1: Row index is RELATIVE');
                FDetails.Add('');
                if w and $8000 = 0
                  then FDetails.Add('Absolute row index: ' + IntToStr(r))
                  else FDetails.Add('Relative row index: ' + IntToStr(Smallint(r)));
              end;
              ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
                'Encoded row index');

              if Row = FCurrRow then
              begin
                FDetails.Add('Encoded cell address (column):'#13);
                if w and $4000 = 0 then begin
                  FDetails.Add('Column index is ABSOLUTE (see Encoded row index)');
                  FDetails.Add('Absolute column index: ' + IntToStr(c));
                end else begin
                  FDetails.Add('Column index is RELATIVE (see Encoded row index)');
                  FDetails.Add('Relative column index: ' + IntToStr(ShortInt(c)));
                end;
              end;
              numBytes := 1;
              ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b, b]),
                'Encoded column index');
            end;
          end;

    else
          ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [token]),
            '(unknown token)');

    end;  // case
  end;  // while

end;

procedure TBIFFGrid.ShowGCW;
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  i,j: Integer;
  bit: Byte;
begin
  RowCount := FixedRows + 33;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Size of Global Column Width bit field (1 bit per column), must be 32');

  numBytes := 1;
  for i:= 0 to w-1 do begin
    b := FBuffer[FBufferIndex];
    if FCurrRow = Row then begin
      FDetails.Add(Format('GCW (Global column width) record, byte #%d:'#13, [i]));
      bit := 1;
      for j:=0 to 7 do begin
        if b and bit = 0
          then FDetails.Add(Format('Bit $%.2x=0: Column %d uses width of COLWIDTH record.', [bit, j+i*8]))
          else FDetails.Add(Format('Bit $%.2x=1: Column %d uses width of STANDARDWIDTH record '+
                                   '(or, if not available, DEFCOLWIDTH record)', [bit, j+i*8]));
        bit := bit * 2;
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
      Format('Widths of columns %d-%d', [i*8, i*8+7]));
  end;
end;


procedure TBIFFGrid.ShowHeader;
var
  numbytes: Integer;
  s: String;
begin
  RowCount := FixedRows + 1;
  ExtractString(FBufferIndex, IfThen(FFormat=sfExcel8, 2, 1), FFormat=sfExcel8,
    s, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
    'Page header string' + IfThen(FFormat = sfExcel8,
    ' (Unicode string, 16-bit string length)',
    ' (byte string, 8-bit string length)'
  ));
end;


procedure TBIFFGrid.ShowHideObj;
var
  numBytes: word;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Viewing mode for objects:'#13);
    case w of
      0: FDetails.Add('0 = Show all objects');
      1: FDetails.Add('1 = Show placeholders');
      2: FDetails.Add('2 = Do not show objects');
    end;
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Viewing mode for objects');
end;


procedure TBIFFGrid.ShowHyperlink;
var
  numbytes: Word;
  w: Word;
  dw: DWord;
  n: Integer;
  guid: TGUID;
  flags: DWord;
  nchar, size: DWord;
  widestr: WideString;
  ansistr: ansistring;
  s: String;
begin
  n := 0;
  RowCount := FixedRows + 1000;

  ShowCellAddressRange(FFormat);
  inc(n, 4);

  numbytes := 16;
  Move(FBuffer[FBufferIndex], guid, numbytes);
  s := GuidToString(guid);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, GuidToString(guid), 'GUID of standard link');
  inc(n);

  numbytes := 4;
  Move(FBuffer[FBufferIndex], dw, numbytes);
  dw := DWordToLE(dw);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x (%d)', [dw, dw]), 'Unknown');
  inc(n);

  numbytes := 4;
  Move(FBuffer[FBufferIndex], dw, numbytes);
  flags := DWordToLE(dw);
  if FCurrRow = row then begin
    FDetails.Add('Option flags:'#13);
    if flags and $0001 = 0
      then FDetails.Add('  Bit  $0001=0: No link')
      else FDetails.ADd('* Bit  $0001=1: File link or URL');
    if flags and $0002 = 0
      then FDetails.Add('  Bit  $0002=0: Relative path')
      else FDetails.Add('* Bit  $0002=1: Absolute path');
    if flags and $0014 = 0
      then FDetails.Add('  Bits $0014=0: No desriptions')
      else FDetails.Add('* Bits $0014=1: Description (both bits)');
    if flags and $0008 = 0
      then FDetails.Add('  Bit  $0008=0: No text mark')
      else FDetails.Add('* Bit  $0008=1: Text mark');
    if flags and $0080 = 0
      then FDetails.Add('  Bit  $0080=0: No target frame')
      else FDetails.Add('* Bit  $0080=1: Target frame');
    if flags and $0100 = 0
      then FDetails.Add('  Bit  $0100=0: File link or URL')
      else FDetails.Add('* Bit  $0100=1: UNC path (incl server name)');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x (%d)', [flags, flags]),
    'Option flags');
  inc(n);

  if flags and $0014 = $0014 then  // hyperlink has description
  begin
    numBytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    nchar := DWordToLE(dw);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(nchar),
      'Character count of description text, incl trailing zero word');
    inc(n);

    numbytes := 2*nchar;
    SetLength(widestr, nchar);
    Move(FBuffer[FBufferIndex], widestr[1], 2*nchar);
    s := UTF16ToUTF8(widestr);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
      'Character array of description text (no unicode string header, always 16-bit characters, zero-terminated)');
    inc(n);
  end;

  if flags and $0080 <> 0 then  // Hyperlink has target frame
  begin
    numBytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    nchar := DWordToLE(dw);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(nchar),
      'Character count of target frame, incl trailing zero word');
    inc(n);

    numbytes := 2*nchar;
    SetLength(widestr, nchar);
    Move(FBuffer[FBufferIndex], widestr[1], 2*nchar);
    s := UTF16ToUTF8(widestr);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
      'Character array of target frame (no unicode string header, always 16-bit characters, zero-terminated)');
    inc(n);
  end;

  if flags and $0011 <> 0 then // hyperlink contains URL ***OR*** a local file
  begin
    numbytes := 16;
    Move(FBuffer[FBufferIndex], guid, numbytes);
    s := GuidToString(guid);
    if s = '{79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}' then    // case: URL
    begin
      ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
        'GUID of URL Moniker');
      inc(n);

      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numbytes);
      size := DWordToLE(dw);
      nchar := size div 2 - 1;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(size),
        'Size of URL character array, incl. trailing zero word');
      inc(n);

      numbytes := 2*nchar;
      SetLength(widestr, nchar);
      Move(FBuffer[FBufferIndex], widestr[1], 2*nchar);
      s := UTF16ToUTF8(widestr);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
        'Character array of the URL (no unicode string header, always 16-bit characters, zero-terminated');
      inc(n);
    end else
    if s = '{00000303-0000-0000-C000-000000000046}' then   // case: local file
    begin
      ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
        'GUID of File Moniker');
      inc(n);

      numbytes := 2;
      Move(FBuffer[FBufferIndex], w, numbytes);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
        'Directory up-level count. Each leading "..\" in the file link is deleted and increases this counter.');
      inc(n);

      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numbytes);
      nchar := DWordToLE(dw);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(nchar),
        'Character count of the shortened file path and name, incl trailing zero byte.');
      inc(n);

      numbytes := nchar;
      SetLength(ansistr, nchar);
      Move(FBuffer[FBufferIndex], ansistr[1], nchar);
      s := AnsiToUTF8(ansistr);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
        'Character array of the shortened file path and name (no unicode string header, always 8-bit characters, zero-terminated)');
      inc(n);

      for w := 1 to 6 do begin
        numbytes := 4;
        Move(FBuffer[FBufferIndex], dw, numbytes);
        ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [DWordToLE(dw)]),
          'Unknown');
        inc(n);
      end;

      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numbytes);
      size := DWordToLE(dw);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(size),
        'Size of following file link field incl string length field and additional data field');
      inc(n);

      if size > 0 then begin
        numBytes := 4;
        Move(FBuffer[FBufferIndex], dw, numbytes);
        size := DWordToLE(dw);
        nchar := size div 2;
        ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(size),
          'Size of extended file path and name character array');
        inc(n);

        numBytes := 2;
        Move(FBuffer[FBufferIndex], w, numbytes);
        ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordToLE(w)),
          'Unknown');
        inc(n);

        numBytes := size;
        SetLength(widestr, nchar);
        Move(FBuffer[FBufferIndex], widestr[1], numbytes);
        s := UTF16ToUTF8(widestr);
        ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
          'Character array of extended file path and array (No unicode string header, always 16-bit characters, NOT zero-terminated)');
        inc(n);
      end;
    end;
  end  // if flags and $0011 <> 0
  else begin
    // case: Hyperlink to current workbook
  end;

  if flags and $0008 <> 0 then   // hyperlink contains text mark field
  begin
    numBytes := 4;

    Move(FBuffer[FBufferIndex], dw, numbytes);
    nchar := DWordToLE(dw);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(nchar),
      'Character count of the text mark, incl trailing zero word');
    inc(n);

    numbytes := 2*nchar;
    SetLength(widestr, nchar);
    Move(FBuffer[FBufferIndex], widestr[1], 2*nchar);
    s := UTF16ToUTF8(widestr);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
      'Character array of the text mark, without "#" sign, no unicode string header, always 16-bit characters, zero-terminated)');
    inc(n);
  end;

  RowCount := FixedRows + n;
end;


procedure TBIFFGrid.ShowHyperlinkTooltip;
var
  numbytes: Word;
  w: Word;
  widestr: widestring;
  s: string;
begin
  RowCount := FixedRows + 6;

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordToLE(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x (%d)', [w, w]),
    'Repeated record ID');

  ShowCellAddressRange(sfExcel8);

  numbytes := Length(FBuffer) - FBufferIndex;
  SetLength(widestr, numbytes div 2);
  Move(FBuffer[FBufferIndex], widestr[1], numbytes);
  s := UTF16toUTF8(widestr);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, s,
    'Character array of the tool tip, no Unicode string header, always 16-bit characters, zero-terminated');
end;


procedure TBIFFGrid.ShowInRow(var ARow: Integer; var AOffs: LongWord;
  ASize: Word; AValue,ADescr: String; ADescrOnly: Boolean = false);
begin
  if ADescrOnly then
  begin
    Cells[0, ARow] := '';
    Cells[1, ARow] := '';
    Cells[2, ARow] := '';
  end else
  begin
    Cells[0, ARow] := IntToStr(AOffs);
    Cells[1, ARow] := IntToStr(ASize);
    Cells[2, ARow] := AValue;
  end;
  Cells[3, ARow] := ADescr;
  inc(ARow);
  inc(AOffs, ASize);
end;


procedure TBIFFGrid.ShowInteger;
var
  numBytes: Integer;
  w: Word;
  b: Byte;
begin
  // BIFF2 only
  if (FFormat <> sfExcel2) then
    exit;

  RowCount := FixedRows + 5;
  ShowRowColData(FBufferIndex);

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  if Row = FCurrRow then begin
    FDetails.Add('Cell protection and XF index:'#13);
    FDetails.Add(Format('x Bits 5-0 = %d: XF Index', [b and $3F]));
    case b and $40 of
      0: FDetails.Add('  Bit 6 = 0: Cell is NOT locked.');
      1: FDetails.Add('x Bit 6 = 1: Cell is locked.');
    end;
    case b and $80 of
      0: FDetails.Add('  Bit 7 = 0: Formula is NOT hidden.');
      1: FDetails.Add('x Bit 7 = 1: Formula is hidden.');
    end;
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
    'Cell protection and XF index');

  b := FBuffer[FBufferIndex];
  if Row = FCurrRow then begin
    FDetails.Add('Indexes to format and font records:'#13);
    FDetails.Add(Format('Bits 5-0 = %d: Index to FORMAT record', [b and $3f]));
    FDetails.Add(Format('Bits 7-6 = %d: Index to FONT record', [(b and $C0) shr 6]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
    'Indexes of format and font records');

  b := FBuffer[FBufferIndex];
  if Row = FCurrRow then begin
    FDetails.Add('Cell style:'#13);
    case b and $07 of
      0: FDetails.Add('x Bits 2-0 = 0: Horizontal alignment is GENERAL');
      1: FDetails.Add('x Bits 2-0 = 1: Horizontal alignment is LEFT');
      2: FDetails.Add('x Bits 2-0 = 2: Horizontal alignment is CENTERED');
      3: FDetails.Add('x Bits 2-0 = 3: Horizontal alignment is RIGHT');
      4: FDetails.Add('x Bits 2-0 = 4: Horizontal alignment is FILLED');
    end;
    if b and $08 = 0
      then FDetails.Add('  Bit 3 = 0: Cell has NO left border')
      else FDetails.Add('x Bit 3 = 1: Cell has left black border');
    if b and $10 = 0
      then FDetails.Add('  Bit 4 = 0: Cell has NO right border')
      else FDetails.Add('x Bit 4 = 1: Cell has right black border');
    if b and $20 = 0
      then FDetails.Add('  Bit 5 = 0: Cell has NO top border')
      else FDetails.Add('x Bit 5 = 1: Cell has top black border');
    if b and $40 = 0
      then FDetails.Add('  Bit 6 = 0: Cell has NO bottom border')
      else FDetails.Add('x Bit 6 = 1: Cell has bottom black border');
    if b and $80 = 0
      then FDetails.Add('  Bit 7 = 0: Cell has NO shaded background')
      else FDetails.Add('x Bit 7 = 1: Cell has shaded background');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.2x)', [b,b]),
    'Cell style');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Unsigned 16-bit integer cell value');
end;


procedure TBIFFGrid.ShowInterfaceEnd;
begin
  RowCount := FixedRows + 1;
  ShowInRow(FCurrRow, FBufferIndex, 0, '', 'End of Globals Substream');
end;


procedure TBIFFGrid.ShowInterfaceHdr;
var
  numbytes: Integer;
  w: Word;
begin
  if FFormat < sfExcel8 then begin
    RowCount := FixedRows;
    exit;
  end;

  RowCount := FixedRows + 1;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Code page of user interface:'#13);
    FDetails.Add(Format('$%.4x = %s', [w, CodePageName(w)]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w, w]),
    'Begin of Globals Substream, code page of user interface');
end;


procedure TBIFFGrid.ShowIteration;
var
  numBytes: Integer;
  w: word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Iterations:'#13);
    case w of
      0: FDetails.Add('0 = Iterations off');
      1: FDetails.Add('1 = Iterations on');
    end;
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Iterations on/off');
end;


procedure TBIFFGrid.ShowIXFE;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to XF record');
end;

procedure TBIFFGrid.ShowLabelCell;
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  s: String;
begin
  RowCount := IfThen(FFormat = sfExcel2, FixedRows + 6, FixedRows + 4);
  ShowRowColData(FBufferIndex);
  if (FFormat = sfExcel2) then begin
    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell protection and XF index:'#13);
      FDetails.Add(Format('x Bits 5-0 = %d: XF Index', [b and $3F]));
      case b and $40 of
        0: FDetails.Add('  Bit 6 = 0: Cell is NOT locked.');
        1: FDetails.Add('x Bit 6 = 1: Cell is locked.');
      end;
      case b and $80 of
        0: FDetails.Add('  Bit 7 = 0: Formula is NOT hidden.');
        1: FDetails.Add('x Bit 7 = 1: Formula is hidden.');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Cell protection and XF index');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Indexes to format and font records:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: Index to FORMAT record', [b and $3f]));
      FDetails.Add(Format('Bits 7-6 = %d: Index to FONT record', [(b and $C0) shr 6]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Indexes of format and font records');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell style:'#13);
      case b and $07 of
        0: FDetails.Add('Bits 2-0 = 0: Horizontal alignment is GENERAL');
        1: FDetails.Add('Bits 2-0 = 1: Horizontal alignment is LEFT');
        2: FDetails.Add('Bits 2-0 = 2: Horizontal alignment is CENTERED');
        3: FDetails.Add('Bits 2-0 = 3: Horizontal alignment is RIGHT');
        4: FDetails.Add('Bits 2-0 = 4: Horizontal alignment is FILLED');
      end;
      if b and $08 = 0
        then FDetails.Add('Bit 3 = 0: Cell has NO left border')
        else FDetails.Add('Bit 3 = 1: Cell has left black border');
      if b and $10 = 0
        then FDetails.Add('Bit 4 = 0: Cell has NO right border')
        else FDetails.Add('Bit 4 = 1: Cell has right black border');
      if b and $20 = 0
        then FDetails.Add('Bit 5 = 0: Cell has NO top border')
        else FDetails.Add('Bit 5 = 1: Cell has top black border');
      if b and $40 = 0
        then FDetails.Add('Bit 6 = 0: Cell has NO bottom border')
        else FDetails.Add('Bit 6 = 1: Cell has bottom black border');
      if b and $80 = 0
        then FDetails.Add('Bit 7 = 0: Cell has NO shaded background')
        else FDetails.Add('Bit 7 = 1: Cell has shaded background');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.2x)', [b,b]),
      'Cell style');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Index of XF record');
  end;

  b := IfThen(FFormat=sfExcel2, 1, 2);
  ExtractString(FBufferIndex, b, (FFormat = sfExcel8), s, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, s,
    Format('%s string, %d-bit string length', [GetStringType, b*8]));
end;


procedure TBIFFGrid.ShowLabelSSTCell;
var
  numBytes: Integer;
  w: Word;
  dw: DWord;
begin
  RowCount := FixedRows + 4;
  ShowRowColData(FBufferIndex);

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
    'Index of XF record');

  numBytes := 4;
  Move(FBuffer[FBufferIndex], dw, numBytes);
  dw := DWordLEToN(dw);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(dw),
    'Index into SST record (shared string table)');
end;


procedure TBIFFGrid.ShowLeftMargin;
var
  numBytes: Integer;
  dbl: Double;
begin
  RowCount := FixedRows + 1;
  numBytes := 8;
  Move(FBuffer[FBufferIndex], dbl, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl),
    'Left page margin in inches (IEEE 754 floating-point value, 64-bit double precision)');
end;


procedure TBIFFGrid.ShowMergedCells;
var
  w: Word;
  numBytes: Integer;
  i, n: Integer;
begin
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  n := WordLEToN(w);  // count of merged ranges in this record

  RowCount := FixedRows + 1 + n*4;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(n),
    'Count of merged ranges in this record');

  for i:=1 to n do begin
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
      Format('Merged range #%d: First row = %d', [i, w]));
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
      Format('Merged range #%d: Last row = %d', [i, w]));
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
      Format('Merged range #%d: First column = %d', [i, w]));
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
      Format('Merged range #%d: Last column = %d', [i, w]));
  end;
end;


procedure TBIFFGrid.ShowMMS;
var
  w: Word;
  numbytes: Integer;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w), 'Reserved, MUST be ignored');
end;

{ Only those records needed for comments are deciphered. }
procedure TBIFFGrid.ShowMSODrawing;
var
  n: Integer;
  numBytes: Integer;
  w: Word;
  dw: DWord;
  recType: Word;
  recLen: DWord;
  isContainer: Boolean;
  indent: String;
  level: Integer;

  function PrepareIndent(ALevel: Integer): String;
  var
    i: Integer;
  begin
    Result := '';
    for i := 1 to ALevel do Result := Result + '        ';
  end;

  procedure DoShowHeader(out ARecType: Word; out ARecLen: DWord;
    out IsContainer: Boolean);
  var
    s: String;
    instance: Word;
    version: Word;
  begin
    numbytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    IsContainer := (w and $000F = $000F);

    if IsContainer and (FBufferIndex > 0) then
      indent := PrepareIndent(level+1);

    s := IfThen(IsContainer, '***** CONTAINER *****', '--- ATOM ---');
    ShowInRow(FCurrRow, FBufferIndex, 0, '', indent + s, TRUE);
    inc(n);

    version := w and $000F;
    instance := (w and $FFF0) shr 4;
    if Row = FCurrRow then begin
      FDetails.Add('OfficeArtDrawing Header:'#13);
      FDetails.Add(Format('Bits 3-0 = $%.1x: Version', [version]));
      FDetails.Add(Format('Bits 15-4 = $%.3x: Instance', [instance]));
    end;
    s := Format('OfficeArtDrawing Header: Version %d, instance %d', [version, instance]);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.04x', [w]), indent + s);
    inc(n);

    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    ARecType := w;
    case ARecType of
      $F00A: //, $F00B:
         case instance of
           $01: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Rectangle)';
           $02: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Rounded rectangle)';
           $03: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Ellipse)';
           $04: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Diamond)';
           $05: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Triangle)';
           $06: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Right triangle)';
           $07: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Parallelogram)';
           $08: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Trapezoid)';
           $09: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Hexagon)';
           $0A: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Pentagon)';
           $0B: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Plus)';
           $0C: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Star)';
           $0D: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Arrow)';
           $0E: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Thick arrow)';
           $0F: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Irregular pentagon)';
           $10: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Cube)';
           $11: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Speech balloon)';
           $12: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Seal)';
           $13: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Curved arc)';
           $14: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Line)';
           $15: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Plaque)';
           $16: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Cylinder)';
           $17: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Donut)';
           $CA: Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + ' (Text box)';
         end;
      $F00B:
        Cells[3, FCurrRow-1] := Cells[3, FCurrRow-1] + Format(' (i.e., %d properties)', [instance]);
    end;

    case ARecType of
      $F000: s := 'OfficeArtDggContainer (all the OfficeArt file records containing document-wide data)';
      $F001: s := 'OfficeArtBStoreContainer (all the BLIPs used in all the drawings associated with parent OfficeArtDggContainer record)';
      $F002: s := 'OfficeArtDgContainer (all file records for the objects in a drawing)';
      $F003: s := 'OfficeArtSpgrContainer (groups of shapes)';
      $F004: s := 'OfficeArtSpContainer (shape)';
      $F005: s := 'OfficeArtSolverContainer (rules applicable to the shapes contained in an OfficeArtDgContainer record)';
      $F006: s := 'OfficeArtFDGGBlock record (document-wide information about all drawings saved in the file)';
      $F007: s := 'OfficeArtFBSE record (File BLIP Store Entry (FBSE) containing information about the BLIP)';
      $F008: s := 'OfficeArtFDG record (number of shapes, drawing identifier, and shape identifier of the last shape in a drawing)';
      $F009: s := 'OfficeArtFSPGR record (coordinate system of the group shape that the anchors of the child shape are expressed in)';
      $F00A: s := 'OfficeArtFSP record (instance of a shape)';
      $F00B: s := 'OfficeArtFOPT record (table of OfficeArtRGFOPTE records)';
      $F00D: s := 'OfficeArtClientTextBox (text related data for a shape)';
      $F00F: s := 'OfficeArtChildAnchor record (anchors for the shape containing this record)';
      $F010: s := 'OfficeArtClientAnchor record (location of a shape)';
      $F011: s := 'OfficeArtClientData record (information about a shape)';
      $F012: s := 'OfficeArtFConnectorRule record (connection between two shapes by a connector shape)';
      $F014: s := 'OfficeArtFArcRule record (Specifies an arc rule. Each arc shape MUST correspond to a unique arc rule)';
      $F017: s := 'OfficeArtFCalloutRule record (Callout rule: One callout rule MUST exist per callout shape)';
      $F01A: s := 'OfficeArtBlipEMF record (BLIP file data for the enhanced metafile format (EMF))';
      $F01B: s := 'OfficeArtBlipWMF record (BLIP file data for the Windows Metafile Format (WMF))';
      $F01C: s := 'OfficeArtBlipPICT record (BLIP file data for the Macintosh PICT format)';
      $F01D: s := 'OfficeArtBlipJPEG record (BLIP file data for the Joint Photographic Experts Group (JPEG) format)';
      $F01E: s := 'OfficeArtBlipPNG record (BLIP file data for the Portable Network Graphics (PNG) format)';
      $F01F: s := 'OfficeArtBlipDIB record (BLIP file data for the device-independent bitmap (DIB) format)';
      $F029: s := 'OfficeArtBlipTIFF record (BLIP file data for the TIFF format)';
      $F118: s := 'OfficeArtFRITContainer record (container for the table of group identifiers that are used for regrouping ungrouped shapes)';
      $F119: s := 'OfficeArtFDGSL record (specifies both the selected shapes and the shape that is in focus in the drawing)';
      $F11A: s := 'OfficeArtColorMRUContainer record (most recently used custom colors)';
      $F11D: s := 'OfficeArtFPSPL record (former hierarchical position of the containing object that is either a shape or a group of shapes)';
      $F11E: s := 'OfficeArtSplitMenuColorContainer record (container for the colors that were most recently used to format shapes)';
      $F121: s := 'OfficeArtSecondaryFOPT record (table of OfficeArtRGFOPTE records)';
      $F122: s := 'OfficeArtTertiaryFOPT record (table of OfficeArtRGFOPTE records)';
      else   s := 'Not known to BIFFExplorer';
    end;
    if Row = FCurrRow then begin
      FDetails.Add('OfficeArtDrawing Record Type:'#13);
      FDetails.Add(Format('$%.4x: %s', [ARecType, s]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
      indent + 'Record type: ' + s);
    inc(n);

    numbytes := 4;
    Move(FBuffer[FbufferIndex], ARecLen, Numbytes);
    ARecLen := DWordLEToN(ARecLen);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d', [ARecLen]),
      indent + 'Record length');
    inc(n);
  end;

  procedure DoShowOfficeArtFDG;   // $F008
  // The OfficeArtFDG record specifies the number of shapes, the drawing
  // identifier, and the shape identifier of the last shape in a drawing.
  begin
    numbytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    dw := DWordLEToN(dw);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(dw),
      indent + 'Number of shapes in this drawing');
    inc(n);

    Move(FBuffer[FBufferIndex], dw, numbytes);
    dw := DWordLEToN(dw);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(dw),
      indent + 'Shape identifier of the last shape in this drawing');
    inc(n);
  end;

  procedure DoShowOfficeArtFSPGR;  // $F009
  // The OfficeArtFSPGR record specifies the coordinate system of the group
  // shape that the anchors of the child shape are expressed in. This record
  // is present only for group shapes.
  var
    rt: Word;
  begin
    numbytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(LongInt(DWordLEToN(dw))),
      indent + 'xLeft');
    inc(n);

    Move(FBuffer[FBufferIndex], dw, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(LongInt(DWordLEToN(dw))),
      indent + 'yTop');
    inc(n);

    Move(FBuffer[FBufferIndex], dw, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(LongInt(DWordLEToN(dw))),
      indent + 'xRight');
    inc(n);

    Move(FBuffer[FBufferIndex], dw, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(LongInt(DWordLEToN(dw))),
      indent + 'yBottom');
    inc(n);
  end;

  procedure DoShowOfficeArtFSP;   // §F00A
  // The OfficeArtFSP record specifies an instance of a shape.
  // The record header contains the shape type, and the record itself
  // contains the shape identifier and a set of bits that further define the shape.
  var
    rt: word;
    s: String;
  begin
    numbytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
      indent + 'Shape identifier');
    inc(n);

    numbytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    dw := DWordLEToN(dw);
    s := '';
    if Row = FCurrRow then begin
      FDetails.Add('Shape options:'#13);
      if dw and $0001 = 0
        then FDetails.Add('  Bit 1=0: Shape is NOT a group shape')
        else FDetails.Add('x Bit 1=1: Shape is a group shape');
      if dw and $0002 = 0
        then FDetails.Add('  Bit 2=0: Shape is NOT a child shape')
        else FDetails.Add('x Bit 2=1: Shape is a child shape');
      if dw and $0004 = 0
        then FDetails.Add('  Bit 3=0: Shape is NOT the top-most group shape (patriarch)')
        else FDetails.Add('x Bit 3=1: Shaoe is the top-most group shape (patriarch)');
      if dw and $0008 = 0
        then FDetails.Add('  Bit 4=0: Shape has NOT been deleted')
        else FDetails.Add('x Bit 4=1: Shape has been deleted');
      if dw and $0010 = 0
        then FDetails.Add('  Bit 5=0: Shape is NOT an OLE object')
        else FDetails.Add('x Bit 5=1: Shape is an OLE object');
      if dw and $0020 = 0
        then FDetails.Add('  Bit 6=0: Shape does NOT have a valid master in the hspMaster property')
        else FDetails.Add('x Bit 6=1: Shape has a valid master in the hspMaster property');
      if dw and $0040 = 0
        then FDetails.Add('  Bit 7=0: Shape is NOT horizontally flipped')
        else FDetails.Add('x Bit 7=1: Shape is horizontally flipped');
      if dw and $0080 = 0
        then FDetails.Add('  Bit 8=0: Shape is NOT vertically flipped')
        else FDetails.Add('x Bit 8=1: Shape is vertically flipped');
      if dw and $0100 = 0
        then FDetails.Add('  Bit 9=0: Shape is NOT a connector shape')
        else FDetails.Add('x Bit 9=1: Shape is a connector shape');
      if dw and $0200 = 0
        then FDetails.Add('  Bit 10=0: Shape doe NOT have an anchor')
        else FDetails.Add('x Bit 10=1: Shape has an anchor');
      if dw and $0400 = 0
        then FDetails.Add('  Bit 11=0: Shape is NOT a background shape')
        else FDetails.Add('x Bit 11=1: Shape is a background shape');
      if dw and $0800 = 0
        then FDetails.Add('  Bit 12=0: Shape does NOT have a shape type property')
        else FDetails.Add('x Bit 12=1: Shape has a shape type property');
      FDetails.Add('  Bits 13-32: unused');
    end;
    s := '';
    if dw and $0001 <> 0 then s := s + 'group shape, ';
    if dw and $0002 <> 0 then s := s + 'child shape, ';
    if dw and $0004 <> 0 then s := s + 'patriarch, ';
    if dw and $0008 <> 0 then s := s + 'deleted, ';
    if dw and $0010 <> 0 then s := s + 'OLE object, ';
    if dw and $0020 <> 0 then s := s + 'master, ';
    if dw and $0040 <> 0 then s := s + 'flipped hor, ';
    if dw and $0080 <> 0 then s := s + 'flipped vert, ';
    if dw and $0100 <> 0 then s := s + 'connector shape, ';
    if dw and $0200 <> 0 then s := s + 'anchor, ';
    if dw and $0400 <> 0 then s := s + 'background, ';
    if dw and $0800 <> 0 then s := s + 'shape type, ';
    if s <> '' then begin
      Delete(s, Length(s)-1, 2);
      s := 'Shape options: ' + s;
    end else
      s := 'Shape options';
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x (%d)', [dw, dw]), indent + s);
    inc(n);
  end;

  procedure DoShowOfficeArtFOPT(ARecLen: DWord);  // $F00B
  // The OfficeArtFOPT record specifies a table of OfficeArtRGFOPTE records.
  var
    startIndex: Int64;
    opid: Word;
    op: DWord;
    s: String;
  begin
    startIndex := FBufferIndex;
    while FBufferIndex < startIndex + ARecLen do
    begin
      numbytes := 2;
      Move(FBuffer[FBufferIndex], opid, numbytes);
      opid := WordLEToN(opid);
      case opid of
        $007F: s := 'OfficeArtFOPTEOPID: Protection boolean properties';
        $0080: s := 'OfficeArtFOPTEOPID: TextID';
        $0081: s := 'OfficeArtFOPTEOPID: dxTextLeft';
        $0082: s := 'OfficeArtFOPTEOPID: dyTextTop';
        $0083: s := 'OfficeArtFOPTEOPID: dxTextRight';
        $0084: s := 'OfficeArtFOPTEOPID: dyTextBottom';
        $0085: s := 'OfficeArtFOPTEOPID: Wrap text';
        $0086: s := 'OfficeArtFOPTEOPID: Unused';
        $0087: s := 'OfficeArtFOPTEOPID: Text anchor';
        $0088: s := 'OfficeArtFOPTEOPID: Text flow';
        $0089: s := 'OfficeArtFOPTEOPID: Font rotation';
        $008A: s := 'OfficeArtFOPTEOPID: Next shape in sequence of linked shapes';
        $008B: s := 'OfficeArtFOPTEOPID: Text direction';
        $008C: s := 'OfficeArtFOPTEOPID: Unused';
        $008D: s := 'OfficeArtFOPTEOPID: Unused';
        $00BF: s := 'OfficeArtFOPTEOPID: Boolean properties for the text in a shape';
        $00C0: s := 'OfficeArtFOPTEOPID: Text for this shape’s geometry text';
        $00C2: s := 'OfficeArtFOPTEOPID: Alignment of text in the shape';
        $00C3: s := 'OfficeArtFOPTEOPID: Font size, in points, of the geometry text for this shape';
        $00C4: s := 'OfficeArtFOPTEOPID: Amount of spacing between characters in the text';
        $00C5: s := 'OfficeArtFOPTEOPID: Font to use for the text';
        $0158: s := 'OfficeArtFOPTEOPID: Type of connection point';
        $0181: s := 'OfficeArtFOPTEOPID: Fill color';
        $0183: s := 'OfficeArtFOPTEOPID: Background color of the fill';
        $0185: s := 'OfficeArtFOPTEOPID: Foreground color of the fill';
        $01BF: s := 'OfficeArtFOPTEOPID: Fill style boolean properties';
        $01C0: s := 'OfficeArtFOPTEOPID: Line color';
        $01C3: s := 'OfficeArtFOPTEOPID: Line foreground color for black & white display mode';
        $01FF: s := 'OfficeArtFOPTEOPID: Line style boolean properties';
        $0201: s := 'OfficeArtFOPTEOPID: Shadow color';
        $0203: s := 'OfficeArtFOPTEOPID: Shadow primary color modifier if in black & white mode';
        $023F: s := 'OfficeArtFOPTEOPID: Shadow style boolean properties';
        $03BF: s := 'OfficeArtFOPTEOPID: Group shape boolean properties';
        else   s := 'OfficeArtFOPTEOPID that specifies the header information for this property';
      end;
      s := indent + s;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x (%d)', [opid, opid]), s);
      inc(n);

      numbytes := 4;
      Move(FBuffer[FBufferIndex], op, numbytes);
      op := DWordLEToN(op);
      s := '';
      case opid of
        $0087: case op of
                 0: s := s + ': text at top';
                 1: s := s + ': text vertically centered';
                 2: s := s + ': text at bottom';
                 3: s := s + ': text at top and centered horizontally';
                 4: s := s + ': text at center of box';
                 5: s := s + ': text at bottom and centered horizontally';
               end;
        $0089: case op of
                 0: s := s + ': horizontal';
                 1: s := s + ': 90° clockwise, vertical down';
                 2: s := s + ': horizontal, 180° rotated';
                 3: s := s + ': 90° counter-clickwise, vertical up';
               end;
        $008B: case op of
                 0: s := s + ': left-to-right';
                 1: s := s + ': right-to-left';
                 2: s := s + ': determined from text string';
               end;
        $00C2: case op of
                 0: s := s + ': stretched';
                 1: s := s + ': centered';
                 2: s := s + ': left-aligned';
                 3: s := s + ': right-aligned';
                 4: s := s + ': justified';
                 5: s := s + ': word-justified';
               end;
        $0158: case op of
                 0: s := s + ': no connection point';
                 1: s := s + ': Edit points of shape are used as connection points';
                 2: s := s + ': Custom array of connection points';
                 3: s := s + ': standard 4 connection points at the top, bottom, left, and right side centers';
               end;
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x (%d)', [op, op]),
        indent + 'Value of this property' + s);
      inc(n);
    end;
  end;

  procedure DoShowOfficeArtClientTextbox(ARecLen: Word);  // $F00D
  begin
    numBytes := ARecLen;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, '...', indent + 'Text data follow in TXO record', true);
    inc(n);
  end;

  procedure DoShowOfficeArtClientAnchorData(StructureKind: Integer); // $F010
  begin
    case StructureKind of
      0: begin   // OfficeArtClientAnchorSheet
           numBytes := 2;
           Move(FBuffer[FBufferIndex], w, numbytes);
           w := WordLEToN(w);
           if Row = FCurrRow then begin
             FDetails.Add('Move/resize flags:'#13);
             if w and $0001 = 0
               then FDetails.Add('  Bit 0=0: Shape will NOT be kept intact when the cells are moved')
               else FDetails.Add('x Bit 0=1: Shape will be kept intact when the cells are moved');
             if w and $0002 = 0
               then FDetails.Add('  Bit 1=0: Shape will NOT be kept intact when the cells are resized')
               else FDetails.Add('x Bit 1=1: Shape will be kept intact when the cells are resized');
             FDetails.Add(       '  Bit 2=0: reserved (must be 0)');
             FDetails.Add(       '  Bit 3=0: reserved (must be 0)');
             FDetails.Add(       '  Bit 4=0: reserved (must be 0)');
             FDetails.Add(       '  Bits 15-5: unused');
           end;
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w), indent + 'Move/resize flags');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'Column of the cell under the top left corner of the bounding rectangle of the shape');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'x coordinate of the top left corner of the bounding rectangle relative to the corner of the underlying cell (as 1/1024 of that cell’s width)');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'Row of the cell under the top left corner of the bounding rectangle of the shape');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'y coordinate of the top left corner of the bounding rectangle relative to the corner of the underlying cell (as 1/256 of that cell’s height');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'Column of the cell under the bottom right corner of the bounding rectangle of the shape');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'x coordinate of the bottom right corner of the bounding rectangle relative to the corner of the underlying cell (as 1/1024 of that cell’s width)');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'Row of the cell under the bottom right corner of the bounding rectangle of the shape');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'y coordinate of the bottom right corner of the bounding rectangle relative to the corner of the underlying cell (as 1/256 of that cell’s height');
           inc(n);
         end;
      1: begin  // "Small" rectangle structure
           numBytes := 2;
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'top');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'left');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'right');
           inc(n);
           Move(FBuffer[FBufferIndex], w, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
             indent + 'bottom');
           inc(n);
         end;
      2: begin   // standard rectangle structure
           numbytes := 4;
           Move(FBuffer[FBufferIndex], dw, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
             indent + 'top');
           inc(n);
           Move(FBuffer[FBufferIndex], dw, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
             indent + 'left');
           inc(n);
           Move(FBuffer[FBufferIndex], dw, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
             indent + 'right');
           inc(n);
           Move(FBuffer[FBufferIndex], dw, numbytes);
           ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
             indent + 'bottom');
           inc(n);
         end;
    end;
  end;

  procedure DoShowOfficeArtRecord(out ARecType:Word; out ARecLen: DWord;
    out AIsContainer: Boolean);
  var
    startIndex, endindex: Int64;
    totalbytes: Int64;
  begin
    totalbytes := 0;
    startIndex := FBufferIndex + 8;
    repeat
      DoShowHeader(ARecType, ARecLen, AIsContainer);
      if totalbytes = 0 then endindex := startindex + ARecLen;
      totalbytes := totalbytes + ARecLen;
      if AIsContainer then begin
        inc(level);
        indent := PrepareIndent(level);
        DoShowOfficeArtRecord(ARecType, ARecLen, AIsContainer);
      end else
        case ARecType of
          $F008: DoShowOfficeArtFDG;
          $F009: DoShowOfficeArtFSPGR;  // shape group
          $F00A: DoShowOfficeArtFSP;    // instance of a shape
          $F00B: DoShowOfficeArtFOPT(ARecLen); // table of OfficeArtRGFOPTE records
          $F00D: DoShowOfficeArtClientTextbox(ARecLen);
          $F010: DoShowOfficeArtClientAnchorData(0);  // 0 - use OfficeArtClientAnchorSheet because contained in a sheet stream
//           $F122: DoShowOfficeArtRGFOPTE(AIndent);  // (tertiary) table of OfficeArtRGFOPTE records
          else if ARecLen <> 0 then begin
                 ShowInRow(FCurrRow, FBufferIndex, ARecLen, '', indent + 'Skipping this unknown record...');
                 inc(n);
               end;
        end;
    until (FBufferIndex >= endindex) or (FBufferIndex >= Length(FBuffer));
    dec(level);
    FBufferIndex := endindex;
  end;

begin
  RowCount := FixedRows + 1000;
  n := 0;
  level := -1;
  DoShowOfficeArtRecord(recType, recLen, isContainer);
  RowCount := FixedRows + n;
end;

procedure TBIFFGrid.ShowMulBlank;
var
  w: Word;
  numbytes: Integer;
  i, nc: Integer;
begin
  nc := (Length(FBuffer) - 6) div 2;
  RowCount := FixedRows + 3 + nc;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to row');

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to first column');

  for i:=0 to nc-1 do begin
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      Format('Index to XF record #%d', [i]));
  end;

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to last column');
end;


procedure TBIFFGrid.ShowMulRK;
var
  w: Word;
  numBytes: Integer;
  i, nc: Integer;
  dw: DWord;
  encint: DWord;
  encdbl: QWord;
  dbl: Double absolute encdbl;
  s: String;
begin
  nc := (Length(FBuffer) - 6) div 6;
  RowCount := FixedRows + 3 + nc*2;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to row');

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to first column');

  for i:=0 to nc-1 do begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      Format('Index to XF record #%d', [i]));

    numBytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    dw := DWordLEToN(dw);

    if Row = FCurrRow then begin
      FDetails.Add('RK Value:'#13);
      if dw and $00000001 = 0
        then FDetails.Add('Bit 0 = 0: Value not changed')
        else FDetails.Add('Bit 0 = 1: Encoded value is multiplied by 100.');
      if dw and $00000002 = 0
        then FDetails.Add('Bit 1 = 0: Floating point value')
        else FDetails.Add('Bit 1 = 1: Signed integer value');
      if dw and $00000002 = 0 then begin
        encdbl := (QWord(dw) and QWord($FFFFFFFFFFFFFFFC)) shl 32;
        if dw and $00000001 = 1 then
          s := Format('%.2f', [dbl*0.01])
        else
          s := Format('%.0f', [dbl]);
      end
      else begin
        s := Format('$%.16x', [-59000000]);
        encint := ((dw and DWord($FFFFFFFC)) shr 2) or (dw and DWord($C0000000));
          // "arithmetic shift" = replace left-most bits by original bits
        if dw and $00000001 = 1 then
          s := FloatToStr(encint*0.01)
        else
          s := IntToStr(encint);
      end;
      FDetails.Add('Bits 31-2: Encoded value ' + s);
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes,
      Format('$%.8x', [dw]),
      Format('RK value #%d', [i])
    );
  end;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Index to last column');
end;

procedure TBIFFGrid.ShowNote;
var
  numBytes: Integer;
  w: Word = 0;
  s: String;
begin
  RowCount := IfThen(FFormat = sfExcel8, 6, 5);

  // Offset 0: Row and Col index
  ShowRowColData(FBufferIndex);

  if FFormat = sfExcel8 then begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Comment flags:'#13);
      if (w and $0002 <> 0)
        then FDetails.Add('x Bit 1=1: Comment is shown at all times')
        else FDetails.Add('  Bit 1=0: Comment is not shown at all tiems');
      if (w and $0080 <> 0)
        then FDetails.Add('x Bit 7=1: Row with comment is hidden')
        else FDetails.Add('  Bit 7=0: Row with comment is visible');
      if (w and $0100 <> 0)
        then FDetails.Add('x Bit 8=1: Column with comment is hidden')
        else FDetails.Add('  Bit 8=0: Column with comment is visible');
      FDetails.Add('All other bits are reserved and must be ignored.');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Flags');

    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Object ID');

    ExtractString(FBufferIndex, IfThen(FFormat=sfExcel8, 2, 1), FFormat=sfExcel8,
      s, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Author');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
      'Total length of comment');

    numBytes := Min(Length(FBuffer) - FBufferIndex, 2048);
    SetLength(s, numBytes);
    Move(FBuffer[FBufferIndex], s[1], numBytes);
    SetLength(s, Length(s));
    s := UTF8StringReplace(s, #10, '[\n]', [rfReplaceAll]);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Comment text');
  end;
end;


procedure TBIFFGrid.ShowNumberCell;
var
  numBytes: Integer;
  b: Byte = 0;
  w: Word = 0;
  dbl: Double;
begin
  RowCount := IfThen(FFormat = sfExcel2, FixedRows + 6, FixedRows + 4);
  // Offset 0: Row & Offsset 2: Column
  ShowRowColData(FBufferIndex);
  // Offset 4: Cell attributes (BIFF2) or XF ecord index (> BIFF2)
  if FFormat = sfExcel2 then begin
    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell protection and XF index:'#13);
      FDetails.Add(Format('x Bits 5-0 = %d: XF Index', [b and $3F]));
      case b and $40 of
        0: FDetails.Add('  Bit 6 = 0: Cell is NOT locked.');
        1: FDetails.Add('x Bit 6 = 1: Cell is locked.');
      end;
      case b and $80 of
        0: FDetails.Add('  Bit 7 = 0: Formula is NOT hidden.');
        1: FDetails.Add('x Bit 7 = 1: Formula is hidden.');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Cell protection and XF index');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Indexes to format and font records:'#13);
      FDetails.Add(Format('Bits 5-0 = %d: Index to FORMAT record', [b and $3f]));
      FDetails.Add(Format('Bits 7-6 = %d: Index to FONT record', [(b and $C0) shr 6]));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b,b]),
      'Indexes of format and font records');

    numBytes := 1;
    Move(FBuffer[FBufferIndex], b, numBytes);
    if Row = FCurrRow then begin
      FDetails.Add('Cell style:'#13);
      case b and $07 of
        0: FDetails.Add('x Bits 2-0 = 0: Horizontal alignment is GENERAL');
        1: FDetails.Add('x Bits 2-0 = 1: Horizontal alignment is LEFT');
        2: FDetails.Add('x Bits 2-0 = 2: Horizontal alignment is CENTERED');
        3: FDetails.Add('x Bits 2-0 = 3: Horizontal alignment is RIGHT');
        4: FDetails.Add('x Bits 2-0 = 4: Horizontal alignment is FILLED');
      end;
      if b and $08 = 0
        then FDetails.Add('  Bit 3 = 0: Cell has NO left border')
        else FDetails.Add('x Bit 3 = 1: Cell has left black border');
      if b and $10 = 0
        then FDetails.Add('  Bit 4 = 0: Cell has NO right border')
        else FDetails.Add('x Bit 4 = 1: Cell has right black border');
      if b and $20 = 0
        then FDetails.Add('  Bit 5 = 0: Cell has NO top border')
        else FDetails.Add('x Bit 5 = 1: Cell has top black border');
      if b and $40 = 0
        then FDetails.Add('  Bit 6 = 0: Cell has NO bottom border')
        else FDetails.Add('x Bit 6 = 1: Cell has bottom black border');
      if b and $80 = 0
        then FDetails.Add('  Bit 7 = 0: Cell has NO shaded background')
        else FDetails.Add('x Bit 7 = 1: Cell has shaded background');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.2x)', [b,b]),
      'Cell style');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrROw, FBufferIndex, numBytes, Format('%d ($%.4x)', [w, w]),
      'Index of XF record');
  end;
  // Offset  6: Double value
  numBytes := 8;
  Move(FBuffer[FBufferIndex], dbl, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%g', [dbl]), //FloatToStr(dbl),
    'IEEE 764 floating-point value');
end;


procedure TBIFFGrid.ShowObj;
var
  numBytes: Integer;
  w: Word;
  dw: DWord;
  savedBufferIndex: Integer;
  fieldType: Word;
  fieldSize: Word;
  s: String;
  i, n: Integer;
  guid: TGUID;
begin
  if FFormat = sfExcel8 then begin
    n := 0;
    RowCount := FixedRows + 1000;
    while FBufferIndex < Length(FBuffer) do begin
      savedBufferIndex := FBufferIndex;
      numBytes := 2;
      Move(FBuffer[FBufferIndex], w, numBytes);
      fieldType := WordLEToN(w);
      case fieldType of
        $00: s := 'ftEnd (End of OBJ record)';
        $01: s := '(Reserved)';
        $02: s := '(Reserved)';
        $03: s := '(Reserved)';
        $04: s := 'ftMacro (fmla-style macro)';
        $05: s := 'ftButton (Command button)';
        $06: s := 'ftGmo (Group marker)';
        $07: s := 'ftCf (Clipboard format)';
        $08: s := 'ftPioGrbit (Picture option flags)';
        $09: s := 'ftPictFmla (Picture fmla-style macro)';
        $0A: s := 'ftCbls (Checkbox link)';
        $0B: s := 'ftRbo (Radio button)';
        $0C: s := 'ftSbs (Scrollbar)';
        $0D: s := 'ftNts (Note structure)';
        $0E: s := 'ftSbsFmla (Scroll bar fmla-style macro)';
        $0F: s := 'ftGboData (Group box data)';
        $10: s := 'ftEdoData (Edit control data)';
        $11: s := 'ftRboData (Radio button data)';
        $12: s := 'ftCblsData (Check box data)';
        $13: s := 'ftLbsData (List box data)';
        $14: s := 'ftCblsFmla (Check box link fmla-style macro)';
        $15: s := 'ftCmo (Common object data)';
        else s := '(unknown)';
      end;
      ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.04x', [fieldType]),
        Format('Subrecord type: %s', [s]));
      inc(n);

      Move(FBuffer[FBufferIndex], w, numBytes);
      fieldSize := WordLEToN(w);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(fieldSize), 'Size of subrecord');
      inc(n);

      case fieldType of
        $000D:
          begin  // ftNts, from https://msdn.microsoft.com/en-us/library/office/dd951373%28v=office.12%29.aspx
            numBytes := 16;
            Move(FBuffer[FBufferIndex], guid, numBytes);
            s := GuidToString(guid);
            ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'GUID of comment');
            inc(n);

            numbytes := 2;
            Move(FBuffer[FBufferIndex], w, numBytes);
            ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
              'Shared Note (0 = false, 1 = true)');
            inc(n);

            numBytes := 4;
            Move(FBuffer[FBufferIndex], dw, numBytes);
            ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)),
              'Unused (undefined, must be ignored)');
            inc(n);
          end;

        $0015:
          begin  // common object data
             Move(FBuffer[FBufferIndex], w, numBytes);
             w := WordLEToN(w);
             case w of
               $00: s := 'Group';
               $01: s := 'Line';
               $02: s := 'Rectangle';
               $03: s := 'Oval';
               $04: s := 'Arc';
               $05: s := 'Chart';
               $06: s := 'Text';
               $07: s := 'Button';
               $08: s := 'Picture';
               $09: s := 'Polybon';
               $0A: s := '(Reserved)';
               $0B: s := 'Checkbox';
               $0C: s := 'Option button';
               $0D: s := 'Edit box';
               $0E: s := 'Label';
               $0F: s := 'Dialog box';
               $10: s := 'Spinner';
               $11: s := 'Scrollbar';
               $12: s := 'List box';
               $13: s := 'Group box';
               $14: s := 'Combobox';
               $15: s := '(Reserved)';
               $16: s := '(Reserved)';
               $17: s := '(Reserved)';
               $18: s := '(Reserved)';
               $19: s := 'Comment';
               $1A: s := '(Reserved)';
               $1B: s := '(Reserved)';
               $1C: s := '(Reserved)';
               $1D: s := '(Reserved)';
               $1E: s := 'Microsoft Office drawing';
               else s := '(Unknown)';
             end;
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.04x', [w]),
               'Object type: '+s);
             inc(n);

             Move(FBuffer[FBufferIndex], w, numbytes);
             w := WordLEToN(w);
             ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.04x', [w]),
               'Object ID number');
             inc(n);

             Move(FBuffer[FBufferIndex], w, numbytes);
             w := WordLEToN(w);
             if Row = FCurrRow then begin
               FDetails.Add('Option flags:'#13);
               if w and $0001 <> 0
                 then FDetails.Add('x Bit $0001 = 1: Object is locked when sheet is protected.')
                 else FDetails.Add('  Bit $0001 = 0: Object is NOT locked when sheet is protected.');
               if w and $000E <> 0
                 then FDetails.Add('! Bit $0002 <> 0: Reserved - must be zero - THIS SEEMS TO BE AN ERROR!')
                 else FDetails.Add('  Bit $0002 = 0: Reserved - must be zero');
               if w and $0010 <> 0
                 then FDetails.Add('x Bit $0010 = 1: Image of this object is intended to be included when printing')
                 else FDetails.Add('  Bit $0010 = 0: Image of this object is NOT intended to be included when printing');
               if w and $1FE0 <> 0
                 then FDetails.Add('! Bits 12-5 <> 0: Reserved - must be zero - THIS SEEMS TO BE AN ERROR!')
                 else FDetails.Add('  Bits 12-5 = 0: Reserved - must be zero');
               if w and $2000 <> 0
                 then FDetails.Add('x Bit $2000 = 1: Object uses automatic fill style.')
                 else FDetails.Add('  Bit $2000 = 0: Object does NOT use automatic fill style.');
               if w and $4000 <> 0
                 then FDetails.Add('x Bit $4000 = 1: Object uses automatic line style.')
                 else FDetails.Add('  Bit $4000 = 0: Object does NOT use automatic line style.');
               if w and $8000 <> 0
                 then FDetails.Add('! Bit $8000 = 1: Reserved - must be zero - THIS SEEMS TO BE AN ERROR!')
                 else FDetails.Add('  Bit $8000 = 0: Reserved - must be zero.');
             end;
             s := '';
             if w and $0001 <> 0 then s := s + 'locked, ';
             if w and $0010 <> 0 then s := s + 'print, ';
             if w and $2000 <> 0 then s := s + 'automatic fill style, ';
             if w and $4000 <> 0 then s := s + 'automatic line style, ';
             if s <> '' then begin
               Delete(s, Length(s)-1, 2);
               s := 'Option flags:' + s;
             end else
               s := 'Option flags';
             ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.04x', [w]), s);
             inc(n);
           end;
      end;
      FBufferIndex := savedBufferIndex + 4 + fieldSize;
    end;
    RowCount := FixedRows + n;
  end else
  if FFormat = sfExcel5 then begin
      // to do
  end;
end;


procedure TBIFFGrid.ShowPageSetup;
var
  numBytes: Integer;
  w: Word;
  s: String;
  dbl: Double;
begin
  if FFormat in [sfExcel5, sfExcel8] then RowCount := FixedRows + 11
    else RowCount := FixedRows + 6;

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Paper size:'#13);
    FDetails.Add(Format('%d - %s', [w, PaperSizeName(w)]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Paper size');

  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Scaling factor in percent');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLETON(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Start page number');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLETON(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Fit worksheet width to this number of pages (0 = use as many as needed)');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLETON(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Fit worksheet height to this number of pages (0 = use as many as needed)');

  { Excel4 not supported in fpspreadsheet }
  {
  if FFormat = sfExcel4 then begin
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLETON(w);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:'#13);
      if w and $0001 = 0
        then FDetails.Add('Bit $0001 = 0: Print pages in columns')
        else FDetails.Add('Bit $0001 = 1: Print pages in rows');
      if w and $0002 = 0
        then FDetails.add('Bit $0002 = 0: Landscape')
        else FDetails.Add('Bit $0002 = 1: Portrait');
      if w and $0004 = 0
        then FDetails.Add('Bit $0004 = 0: Paper size, scaling factor, and paper orientation ' +
                                          '(portrait/landscape) ARE initialised')
        else FDetails.Add('Bit $0004 = 1: Paper size, scaling factor, and paper orientation ' +
                                          '(portrait/landscape) are NOT initialised');
      if w and $0008 = 0
        then FDetails.Add('Bit $0008 = 0: Print colored')
        else FDetails.add('Bit $0008 = 1: Print black and white');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x (%d)', [w, w]),
      'Option flags');
  end else }
  if (FFormat in [sfExcel5, sfExcel8]) then begin
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:'#13);
      if w and $0001 = 0
        then FDetails.Add('Bit $0001 = 0: Print pages in columns')
        else FDetails.Add('Bit $0001 = 1: Print pages in rows');
      if w and $0002 = 0
        then FDetails.add('Bit $0002 = 0: Landscape')
        else FDetails.Add('Bit $0002 = 1: Portrait');
      if w and $0004 = 0
        then FDetails.Add('Bit $0004 = 0: Paper size, scaling factor, and paper orientation ' +
                                          '(portrait/landscape) ARE initialised')
        else FDetails.Add('Bit $0004 = 1: Paper size, scaling factor, and paper orientation ' +
                                          '(portrait/landscape) are NOT initialised');
      if w and $0008 = 0
        then FDetails.Add('Bit $0008 = 0: Print colored')
        else FDetails.add('Bit $0008 = 1: Print black and white');
      if w and $0010 = 0
        then FDetails.Add('Bit $0010 = 0: Default print quality')
        else FDetails.Add('Bit $0010 = 1: Draft quality');
      if w and $0020 = 0
        then FDetails.Add('Bit $0020 = 0: Do NOT print cell notes')
        else FDetails.Add('Bit $0020 = 0: Print cell notes');
      if w and $0040 = 0
        then FDetails.Add('Bit $0040 = 0: Use paper orientation (portrait/landscape) flag abov')
        else FDetails.Add('Bit $0040 = 1: Use default paper orientation (landscape for chart sheets, portrait otherwise)');
      if w and $0080 = 0
        then FDetails.Add('Bit $0080 = 0: Automatic page numbers')
        else FDetails.Add('Bit $0080 = 1: Use start page number above');
      if w and $0200 = 0
        then FDetails.Add('Bit $0200 = 0: Print notes as displayed')
        else FDetails.Add('Bit $0200 = 1: Print notes at end of sheet');
      case (w and $0C00) shr 10 of
        0: FDetails.Add('Bit $0C00 = 0: Print errors as displayed');
        1: FDetails.add('Bit $0C00 = 1: Do not print errors');
        2: FDetails.Add('Bit $0C00 = 2: Print errors as "--"');
        3: FDetails.Add('Bit $0C00 = 4: Print errors as "#N/A"');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x (%d)', [w, w]),
      'Option flags');

    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Print resolution in dpi');

    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Vertical print resolution in dpi');

    numBytes := 8;
    Move(FBuffer[FBufferIndex], dbl, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl), 'Header margin');

    Move(FBuffer[FBufferIndex], dbl, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl), 'Footer margin');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Number of copies to print');
  end;

end;

procedure TBIFFGrid.ShowPalette;
var
  numBytes: Integer;
  w: Word;
  dw: DWord;
  npal: Integer;
  i: Integer;
  s: String;
begin
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  npal := WordLEToN(w);

  RowCount := FixedRows + 1 + npal;

  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(npal),
    'Number of palette colors');

  for i := 0 to npal-1 do begin
    numbytes := 4;
    Move(FBuffer[FBufferIndex], dw, numBytes);
    dw := DWordLEToN(dw);
    s := Format('Palette color, index #%d ($%.2x)',[i, i]);
    case i of
      $00..$07: ;
      $08..$3F: s := s + ', user-defined palette';
      $40     : s := s + ', system window text color for border lines';
      $41     : s := s + ', system window background color for pattern background';
      $43     : s := s + ', system face color (dialogue background color)';
      $4D     : s := s + ', system window text colour for chart border lines';
      $4E     : s := s + ', system window background color for chart areas';
      $4F     : s := s + ', automatic color for chart border lines (seems to be always Black)';
      $50     : s := s + ', system tool tip background color (used in note objects)';
      $51     : s := s + ', system tool tip text color (used in note objects)';
      $7FFF   : s := s + ', system window text color for fonts';
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.8x', [dw]), s);
  end;
end;


procedure TBIFFGrid.ShowPane;
var
  numBytes: Integer;
  w: Word;
  b: Byte;
begin
  RowCount := FixedRows + IfThen(FFormat < sfExcel5, 5, 6);

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Position of vertical split (twips or columns (if frozen))');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Position of horizontal split (twips or rows (if frozen))');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to first visible row in bottom pane(s)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to first visible column in right pane(s)');

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(b),
    'Identifier of pane with active cell cursor');

  if FFormat >= sfExcel5 then begin
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBUfferIndex, numBytes, IntToStr(b), 'not used');
  end;
end;

procedure TBIFFGrid.ShowPassword;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Password verifier for sheet or workbook:'#13);
    if w = 0
      then FDetails.Add('0 = No password')
      else FDetails.Add(Format('$%.4x = Password verifier', [w]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Password verifier for sheet or workbook');
end;


procedure TBIFFGrid.ShowPrecision;
var
  numbytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Precision-as-displayed mode:'#13);
    if w = 0
      then FDetails.Add('0 = Precision-as-displayed mode selected')
      else FDetails.Add('1 = Precision-as-displayed mode NOT selected');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Precision-as-displayed mode');
end;


procedure TBIFFGrid.ShowPrintGridLines;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Print sheet grid lines:'#13);
    if w = 0
      then FDetails.Add('0 = Do not print sheet grid lines')
      else FDetails.Add('1 = Print sheet grid lines');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Print sheet grid lines');
end;


procedure TBIFFGrid.ShowPrintHeaders;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Print row/column headers');
    FDetails.Add('(the area with row numbers and column letters):'#13);
    if w = 0
      then FDetails.Add('0 = Do not print row/column headers')
      else FDetails.Add('1 = Print row/column headers');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Print row/column headers');
end;


procedure TBIFFGrid.ShowProt4Rev;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Removal of the shared workbook''s revision logs:'#13);
    if w = 0
      then FDetails.Add('0 = Removal of the shared workbook''s revision logs is allowed.')
      else FDetails.Add('1 = Removal of the shared workbook''s revision logs is NOT allowed.');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Removal of the shared workbook''s revision logs');
end;


procedure TBIFFGrid.ShowProt4RevPass;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Password verifier needed to change the PROT4REV record:'#13);
    if w = 0
      then FDetails.Add('0 = No password.')
      else FDetails.Add(Format('$%.04x = Password verifier.', [w]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.4x)', [w,w]),
    'Password verifier needed to change the PROT4REV record');
end;


procedure TBIFFGrid.ShowProtect;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Protection state of the workbook:'#13);
    if w = 0
      then FDetails.Add('  0 = Workbook is NOT protected.')
      else FDetails.Add('x 1 = Workbook is protected.');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    'Protection state of the workbook');
end;


procedure TBIFFGrid.ShowRecalc;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('"Recalculate before save" option:'#13);
    if w = 0
      then FDetails.Add('  0 = Do not recalculate')
      else FDetails.Add('x 1 = Recalculate before saving the document');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Recalculate before saving');
end;


procedure TBIFFGrid.ShowRefMode;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FbufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Cell reference mode:'#13);
    if w = 0
      then FDetails.Add('0 = RC mode (i.e. cell address shown as "R(1)C(-1)"')
      else FDetails.Add('1 = A1 mode (i.e. cell address shown as "B1")');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Cell reference mode');
end;


procedure TBIFFGrid.ShowRefreshAll;
var
  numbytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('RefreshAll record:'#13);
    if w = 0
      then FDetails.Add('  0 = Do not force refresh of external data ranges, PivotTables and XML maps on workbook load.')
      else FDetails.Add('x 1 = Force refresh of external data ranges, PivotTables and XML maps on workbook load.');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w, w]),
    'Force refresh of external data ranges, Pivot tables and XML maps on workbook load');
end;


procedure TBIFFGrid.ShowRightMargin;
var
  numBytes: Integer;
  dbl: Double;
begin
  RowCount := FixedRows + 1;
  numBytes := 8;
  Move(FBuffer[FBufferIndex], dbl, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl),
    'Right page margin in inches (IEEE 754 floating-point value, 64-bit double precision)');
end;


procedure TBIFFGrid.ShowRK;
var
  numBytes: Integer;
  w: Word;
  dw: DWord;
  encint: DWord;
  encdbl: QWord;
  dbl: Double absolute encdbl;
  s: String;
begin
  RowCount := FixedRows + 4;

  ShowRowColData(FBufferIndex);

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to XF record');

  numBytes := 4;
  Move(FBuffer[FBufferIndex], dw, numBytes);
  dw := DWordLEToN(dw);

  if dw and $00000002 = 0 then begin
    encdbl := (QWord(dw) and QWord($FFFFFFFFFFFFFFFC)) shl 32;
    if dw and $00000001 = 1 then
      s := Format('%.2f', [dbl*0.01])
    else
      s := Format('%.0f', [dbl]);
  end
  else begin
    s := Format('$%.16x', [-59000000]);
    encint := ((dw and DWord($FFFFFFFC)) shr 2) or (dw and DWord($C0000000));
      // "arithmetic shift" = replace left-most bits by original bits
    if dw and $00000001 = 1 then
      s := FloatToStr(encint*0.01)
    else
      s := IntToStr(encint);
  end;

  if Row = FCurrRow then begin
    FDetails.Add('RK Value:'#13);
    if dw and $00000001 = 0
      then FDetails.Add('Bit 0 = 0: Value not changed')
      else FDetails.Add('Bit 0 = 1: Encoded value is multiplied by 100.');
    if dw and $00000002 = 0
      then FDetails.Add('Bit 1 = 0: Floating point value')
      else FDetails.Add('Bit 1 = 1: Signed integer value');
    FDetails.Add('Bits 31-2: Encoded value ' + s);
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes,
    Format('$%.8x', [QWord(dw)]), 'RK value ['+s+']');
end;


procedure TBIFFGrid.ShowRow;
var
  numBytes: Integer;
  dw: DWord;
  w: Word;
  b: Byte;
begin
  RowCount := FixedRows + IfThen(FFormat = sfExcel2, 10, 7);

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index of this row');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to column of the first cell which is described by a cell record');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to column of the last cell which is described by a cell record, increased by 1');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Row height:'#13);
    FDetails.Add(Format('Bits 14-0 = %d: Row height in twips (1/20 pt) --> %.1f-pt',
      [w and $7FFF, (w and $7FFF)/20.0])
    );
    if w and $8000 = 0
      then FDetails.Add('Bit 15 = 0: Row has custom height')
      else FDetails.Add('Bit 15 = 1: Row has default height');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    'Bits 14-0: Height of row in twips (1/20 pt), Bit 15: Row has default height');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, '', '(not used)');

  if FFormat = sfExcel2 then begin
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
      '0=No defaults written, 1=Default row attribute field and XF index occur below');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'Relative offset to calculate stream position of the first cell record for this row');

    if b = 1 then begin
      numBytes := 1;
      b := FBuffer[FBufferIndex];
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
        'Cell protection and XF index');

      numBytes := 1;
      b := FBuffer[FBufferIndex];
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
        'Indexes to FORMAT and FONT records');

      numBytes := 1;
      b := FBuffer[FBufferIndex];
      ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
        'Cell style');
    end;
  end
  else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, '',
      'In BIFF5-BIFF8 this field is not used anymore, but the DBCELL record instead.');

    numBytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    dw := DWordLEToN(dw);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags and default row formatting:'#13);
      FDetails.Add(Format('Bits 0-2 = %d: Outline level of the row', [dw and $00000007]));
      if dw and $00000010 = 0
        then FDetails.Add('Bit 4 = 0: Outline group does not start or end here and is not collapsed')
        else FDetails.Add('Bit 4 = 1: Outline group starts or ends here and is collapsed');
      if dw and $00000020 = 0
        then FDetails.Add('Bit 5 = 0: Row is NOT hidden')
        else FDetails.Add('Bit 5 = 1: Row is hidden');
      if dw and $00000040 = 0
        then FDetails.Add('Bit 6 = 0: Row height and default font height do match.')
        else FDetails.Add('Bit 6 = 1: Row height and default font height do NOT match.');
      if dw and $00000080 = 0
        then FDetails.Add('Bit 7 = 0: Row does NOT have explicit default format.')
        else FDetails.Add('Bit 7 = 1: Row has explicit default format.');
      FDetails.Add('Bit 8 = 1: Is always 1');
      FDetails.Add(Format('Bits 16-27 = %d: Index to default XF record', [(dw and $0FFF0000) shr 16]));
      if dw and $10000000 = 0
        then FDetails.Add('Bit 28 = 0: No additional space above the row.')
        else FDetails.Add('Bit 28 = 1: Additional space above the row.');
      if dw and $20000000 = 0
        then FDetails.Add('Bit 29 = 0: No additional space below the row.')
        else FDetails.Add('Bit 29 = 1: Additional space below the row.');
      if dw and $40000000 = 0
        then FDetails.Add('Bit 30 = 0: D0 NOT show phonetic text for all cells in the row.')
        else FDetails.Add('Bit 30 = 1: Show phonetic text for all cells in the row.');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.8x', [dw]),
      'Option flags and default row formatting');
  end;
end;

procedure TBIFFGrid.ShowRowColData(var ABufIndex: LongWord);
var
  w: Word;
  numBytes: Integer;
begin
  // Row
  numBytes := 2;
  Move(FBuffer[ABufIndex], w, numBytes);
  ShowInRow(FCurrRow, ABufIndex, numBytes, IntToStr(WordLEToN(w)), 'Index to row');
  // Column
  numBytes := 2;
  Move(FBuffer[ABufIndex], w, numBytes);
  ShowInRow(FCurrRow, ABufIndex, numBytes, IntToStr(WordLEToN(w)), 'Index to column');
end;


procedure TBIFFGrid.ShowSelection;
var
  numBytes: Integer;
  w: word;
  b: Byte;
  i, n: Integer;
begin
  Move(FBuffer[FBufferIndex+7], w, 2);
  n := WordLEToN(w);

  RowCount := FixedRows + 5 + n*4;

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Pane identifier (see PANE record)');

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to row of the active cell');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index to column of the active cell');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
    'Index into the following cell range list to the entry that contains the active cell');

  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(n),
    'Number of following cell range addresses');

  numbytes := 2;
  for i:=1 to n do begin
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)), 'Index to first row');
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)), 'Index to last row');
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Index to first column');
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Index to last column');
  end;
end;


procedure TBIFFGrid.ShowSharedFormula;
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  tokenBytes: Integer;
begin
  BeginUpdate;
  RowCount := FixedRows + 1000;
  // Brute force simplification because of unknown row count at this point
  // Will be reduced at the end.

  // Index to first row
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)), 'Index to first row');

  // Index to last row
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)), 'Index to last row');

  // Index to first column
  numBytes := 1;  // 8-bit also for BIFF8!
  Move(FBuffer[FBufferIndex], b, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Index to first column');

  // Index to last column
  numBytes := 1;  // 8-bit also for BIFF8!
  Move(FBuffer[FBufferIndex], b, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Index to last column');

  // Not used
  numBytes := 1;
  Move(FBuffer[FBufferIndex], b, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Not used');

  // Number of existing FORMULA records for this shared formula
  numBytes := 1;
  Move(FBuffer[FBufferIndex], b, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b), 'Number of FORMULA records in shared formula');

  // Size of Token array (in Bytes)
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  tokenBytes := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(tokenBytes),
    'Size of formula data (in Bytes)');

  // Formula tokens
  ShowFormulaTokens(tokenBytes);

  RowCount := FCurrRow;
  EndUpdate(true);
end;


procedure TBIFFGrid.ShowSheet;
var
  numBytes: Integer;
  dw: DWord;
  b: Byte;
  s: String;
begin
  RowCount := FixedRows + 4;

  numBytes := 4;
  Move(FBuffer[FBufferIndex], dw, numBytes);
  dw := DWordLEToN(dw);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.8x)', [dw, dw]),
    'Absolute stream position of BOF record of sheet represented by this record.');

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
    'Sheet state (0=visible, 1=hidden, 2="very" hidden)');

  numBytes := 1;
  b := FBuffer[FBufferIndex];
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
    'Sheet type ($00=worksheet, $02=Chart, $06=VB module)');

  ExtractString(FBufferIndex, 1, (FFormat = sfExcel8), s, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, s, IfThen(FFormat=sfExcel8,
    'Sheet name (unicode string, 8-bit string length)',
    'Sheet name (byte string, 8-bit string length)')
  );
end;


procedure TBIFFGrid.ShowSheetPR;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Option flags:'#13);
    if w and $0001 = 0
      then FDetails.Add('  Bit $0001 = 0: Do not show automatic page breaks')
      else FDetails.Add('x Bit $0001 = 1: Show automatic page breaks');
    if w and $0010 = 0
      then FDetails.Add('  Bit $0010 = 0: Standard sheet')
      else FDetails.Add('x Bit $0010 = 1: Dialog sheet (BIFF5-BIFF8)');
    if w and $0020 = 0
      then FDetails.Add('  Bit $0020 = 0: No automatic styles in outlines')
      else FDetails.Add('x Bit $0020 = 1: Apply automatic styles to outlines');
    if w and $0040 = 0
      then FDetails.Add('  Bit $0040 = 0: Outline buttons above outline group')
      else FDetails.Add('x Bit $0040 = 1: Outline buttons below outline group');
    if w and $0080 = 0
      then FDetails.Add('  Bit $0080 = 0: Outline buttons left of outline group')
      else FDetails.Add('x Bit $0080 = 1: Outline buttons right of outline group');
    if w and $0100 = 0
      then FDetails.Add('  Bit $0100 = 0: Scale printout in percent')
      else FDetails.Add('x Bit $0100 = 1: Fit printout to number of pages');
    if w and $0200 = 0
      then FDetails.Add('  Bit $0200 = 0: Save external linked values (BIFF3-BIFF4 only)')
      else FDetails.Add('x Bit $0200 = 1: Do NOT save external linked values (BIFF3-BIFF4 only)');
    if w and $0400 = 0
      then FDetails.Add('  Bit $0400 = 0: Do not show row outline symbols')
      else FDetails.Add('x Bit $0400 = 1: Show row outline symbols');
    if w and $0800 = 0
      then FDetails.Add('  Bit $0800 = 0: Do not show column outline symbols')
      else FDetails.Add('x Bit $0800 = 1: Show column outline symbols');
    case (w and $3000) shr 12 of
      0: FDetails.Add('x Bits $3000 = $0000: Arrange windows tiled');
      1: FDetails.Add('x Bits $3000 = $1000: Arrange windows horizontal');
      2: FDetails.Add('x Bits $3000 = $2000: Arrange windows vertical');
      3: FDetails.Add('x Bits $3000 = $3000: Arrange windows cascaded');
    end;
    if w and $4000 = 0
      then FDetails.Add('x Bits $4000 = 0: Excel like expression evaluation (BIFF4-BIFF8 only)')
      else FDetails.Add('x Bits $4000 = 1: Lotus like expression evaluation (BIFF4-BIFF8 only)');
    if w and $8000 = 0
      then FDetails.Add('x Bits $8000 = 0: Excel like formula editing (BIFF4-BIFF8 only)')
      else FDetails.Add('x Bits $8000 = 1: Lotus like formula editing (BIFF4-BIFF8 only)');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x (%d)', [w, w]),
    'Option flags');
end;

procedure TBIFFGrid.ShowSST;
var
  numBytes: Integer;
  s: String;
  total1, total2: DWord;
  i: Integer;
begin
  numBytes := 4;
  Move(FBuffer[FBufferIndex], total1, numBytes);
  Move(FBuffer[FBufferIndex+4], total2, numBytes);
  total1 := DWordLEToN(total1);
  total2 := DWordLEToN(total2);

  RowCount := FixedRows + 2 + total2;

  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(total1),
    'Total number of shared strings in the workbook');
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(total2),
    'Number of following strings');

  for i:=1 to total2 do begin
    ExtractString(FBufferIndex, 2, true, s, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, s, Format('Shared string #%d', [i]));
  end;
end;


procedure TBIFFGrid.ShowStandardWidth;
var
  w: Word;
  numBytes: Integer;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d (%f characters)', [w, w/256]),
    'Default column width (overrides DFCOLWIDTH, in 1/256 of "0" width)');
end;


procedure TBIFFGrid.ShowString;
var
  numBytes: Integer;
  s: String;
begin
  RowCount := FixedRows + 1;
  case FFormat of
    sfExcel2:
      begin
        ExtractString(FBufferIndex, 1, false, s, numBytes);
        ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Byte string, 8-bit string length');
      end;
    sfExcel5:
      begin
        ExtractString(FBufferIndex, 2, false, s, numBytes);
        ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Byte string, 16-bit string length');
      end;
    sfExcel8:
      begin
        ExtractString(FBufferIndex, 2, true, s, numBytes);
        ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Unicode string, 16-bit string length');
      end;
  end;
end;


procedure TBIFFGrid.ShowStyle;
var
  numBytes: Integer;
  b: Byte;
  w: Word;
  s: String;
  isRowLevel: Boolean;
  isColLevel: Boolean;
begin
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if w and $8000 = 0 then
    RowCount := FixedRows + 2
  else
    RowCount := FixedRows + 3;
  if Row = FCurrRow then begin
    FDetails.Add('Style:'#13);
    FDetails.Add(Format('Bits 0-11 = %d: Index to style XF record', [w and $0FFFF]));
    if w and $8000 = 0
      then FDetails.Add('Bit 15 = 0: user-defined style')
      else FDetails.Add('Bit 15 = 1: built-in style');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]), 'Style index and type');

  if w and $8000 = 0 then begin
    if FFormat = sfExcel8 then begin
      ExtractString(FBufferIndex, 2, true, s, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Style name (Unicode string, 16-bit string length)');
    end else begin
      ExtractString(FBufferIndex, 1, false, s, numBytes);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'Style name (Byte string, 8-bit string length)');
    end;
  end else begin
    numbytes := 1;
    b := FBuffer[FBufferIndex];
    isRowLevel := (b = 1);
    isColLevel := (b = 2);
    if FCurrRow = Row then begin
      FDetails.Add('Identifier for built-in cell style:'#13);
      case b of
        0: FDetails.Add('0 = normal');
        1: FDetails.Add('1 = RowLevel (see next field)');
        2: FDetails.Add('2 = ColLevel (see next field)');
        3: FDetails.Add('3 = Comma');
        4: FDetails.Add('4 = Currency');
        5: FDetails.Add('5 = Percent');
        6: FDetails.Add('6 = Comma [0]');
        7: FDetails.Add('7 = Currency [0]');
        8: FDetails.Add('8 = Hyperlink');
        9: FDetails.Add('9 = Followed hyperlink');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
      'Identifier for built-in cell style');

    b := FBuffer[FBufferIndex];
    if FCurrRow = Row then begin
      FDetails.Add('Level for RowLevel or ColLevel style (zero-based):'#13);
      if b = $FF then
        FDetails.Add('$FF = no RowLevel or ColLevel style')
      else
      if isRowLevel then
        FDetails.Add('RowLevel = ' + IntToStr(b))
      else if isColLevel then
        FDetails.Add('ColLevel = ' + IntToStr(b));
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      'Level for RowLevel or ColLevel style (if available)');
  end;
end;


procedure TBIFFGrid.ShowStyleExt;
var
  numBytes: Integer;
  w: Word;
  b: Byte;
  bs: Byte;
  s: String;
begin
  RowCount := FixedRows + 11;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [wordLEToN(w)]),
    'Future record type');
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Attributes:'#13);
    if w and $0001 = 0
      then FDetails.Add('  Bit 0 = 0: The containing record does not specify a range of cells.')
      else FDetails.Add('x Bit 0 = 1: The containing record specifies a range of cells.');
    FDetails.Add('Bit 1: specifies wether to alert the user of possible problems '+
      'when saving the file whithout having reckognized this record.');
    FDetails.Add('Bits 2-15: reserved (MUST be zero, MUST be ignored)');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
    'Attributes');
  numbytes := 8;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, '', 'Reserved');

  numbytes := 1;
  b := FBuffer[FBufferIndex];
  if Row = FCurrRow then begin
    FDetails.Add('Flags:'#13);
    if b and $01 = 0
      then FDetails.Add('  Bit 0 = 0: no built-in style')
      else FDetails.Add('x Bit 0 = 1: built-in style');
    if b and $02 = 0
      then FDetails.Add('  Bit 1 = 0: NOT hidden')
      else FDetails.Add('x Bit 1 = 1: hidden (i.e. is displayed in user interface)');
    FDetails.Add('Bit 2: specifies whether the built-in cell style was modified '+
      'by the user and thus has a custom definition.');
    FDetails.Add('Bit 3-7: Reserved');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]), 'Flags');

  numbytes := 1;
  b := FBuffer[FBufferIndex];
  if Row = FCurrRow then begin
    FDetails.Add('Category:'#13);
    case b of
      0: FDetails.Add('Bits 0-7 = 0: Custom style');
      1: FDetails.Add('Bits 0-7 = 1: Good, bad, neutral style');
      2: FDetails.Add('Bits 0-7 = 2: Data model style');
      3: FDetails.Add('Bits 0-7 = 3: Title and heading style');
      4: FDetails.Add('Bits 0-7 = 4: Themed cell style');
      5: FDetails.Add('Bits 0-7 = 5: Number format style');
    end;
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]), 'Category');

  numbytes := 1;
  b := FBuffer[FBufferIndex];
  if Row = FCurrRow then begin
    FDetails.Add('Built-in style:'#13);
    FDetails.Add('An unsigned integer that specifies the type of the built-in cell style.');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]), 'Built-in style');
  bs := b;

  numbytes := 1;
  b := FBuffer[FBufferindex];
  if Row = FCurrRow then begin
    FDetails.Add('Outline depth level:'#13);
    FDetails.Add('An unsigned integer that specifies the depth level of row/column automatic outlining.');
    if (bs in [1, 2]) then
      FDetails.Add(Format('Bits 0-7 = %d: Outline level is %d', [b, b+1]))
    else
      FDetails.Add(Format('Bits 0-7 = $%.2x: MUST be $FF, MUST be ignoried', [b]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [b]), 'Outline depth level');

  ExtractString(FBufferIndex, 1, true, s, numBytes, true);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, s, 'Name of the style name to extend (Unicode string, 8-bit string length)');

  numbytes := 2;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, '', 'XFProps (reserved)');

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w), 'Count of XFProp structures to follow in array');


end;


procedure TBiffGrid.ShowTabID;
var
  numbytes: Integer;
  w: word;
  i, n: Integer;
begin
  numbytes := 2;
  n := Length(FBuffer) div numbytes;
  RowCount := FixedRows + n;
  for i := 1 to n do begin
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w, w]),
      'Unique sheet identifier');
  end;
end;


procedure TBIFFGrid.ShowTopMargin;
var
  numBytes: Integer;
  dbl: Double;
begin
  RowCount := FixedRows + 1;
  numBytes := 8;
  Move(FBuffer[FBufferIndex], dbl, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numBytes, FloatToStr(dbl),
    'Top page margin in inches (IEEE 754 floating-point value, 64-bit double precision)');
end;


procedure TBIFFGrid.ShowTXO;
var
  numbytes: Word;
  w: Word;
  s, sh, sv, sl: String;
begin
  RowCount := FixedRows + 9;
  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  sh := '';
  case (w and $000E) shr 1 of
    1: sh := sh + 'left-aligned, ';
    2: sh := sh + 'centered, ';
    3: sh := sh + 'right-aligned, ';
    4: sh := sh + 'justified, ';
  end;
  sv := '';
  case (w and $0070) shr 4 of
    1: sv := sv + 'top, ';
    2: sv := sv + 'middle, ';
    3: sv := sv + 'bottom, ';
    4: sv := sv + 'justify, ';
  end;
  s := sh + sv;
  if (w and $0200) shr 9 <> 0 then s := s + 'lock text, ';
  if Row = FCurrRow then begin
    FDetails.Add(     'Option flags:'#13);
    FDetails.Add(     'Bit 0: Reserved');
    if sh = '' then sh := 'none';
    FDetails.Add(     'Bits 1-3: 0 = Horizontal text alignment: ' + sh);
    if sv = '' then sv := 'none';
    FDetails.Add(     'Bits 4-6: 0 = Vertical text alignment: ' + sv);
    FDetails.Add(     'Bits 7-8: Reserved');
    case (w and $0200) shr 9 of
      0: FDetails.Add('Bit 9: Lock Text Option is off.');
      1: FDetails.Add('Bit 9: Lock Text Option is on');
    end;
    FDetails.Add(     'Bits 10-15: Reserved');
  end;
  if s <> '' then begin
    Delete(s, Length(s)-1, 2);
    s := 'Option flags: ' + s;
  end else
    s := 'Option flags';
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]), s);

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  case w of
    0: s := 'No rotation (text appears left to right';
    1: s := 'Text appears top to bottom; letters are upright';
    2: s := 'Text is rotated 90 degrees counterclockwise';
    3: s := 'Text is rotated 90 degrees clockwise';
  end;
  if Row = FCurrRow then begin
    FDetails.Add('Orientation of text with the object boundary:'#13);
    FDetails.Add(Format('%d = %s', [w, s]));
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(w),
    'Orientation of text within the object boundary: ' + s);

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Reserved (must be 0)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Reserved (must be 0)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Reserved (must be 0)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Length (in characters) of text (in first following CONTINUE record)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Length of formatting runs (in second following CONTINUE record)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Reserved (must be 0)');

  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Reserved (must be 0)');
end;

procedure TBIFFGrid.ShowWindow1;
var
  numBytes: Word;
  b: Byte;
  w: word;
begin
  RowCount := FixedRows + IfThen(FFormat < sfExcel5, 5, 9);
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Horizontal position of the document window (in twips = 1/20 pt)');

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Vertical position of the document window (in twips = 1/20 pt)');

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Width of the document window (in twips = 1/20 pt)');

  Move(FBuffer[FBufferIndex], w, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
    'Height of the document window (in twips = 1/20 pt)');

  if FFormat < sfExcel5 then begin
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(b)),
      '0 = Window is visible; 1 = window is hidden');
  end else begin
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:');
      if w and $0001 = 0
        then FDetails.Add('  Bit $0001 = 0: Window is visible')
        else FDetails.Add('x Bit $0001 = 1: Window is hidden');
      if w and $0002 = 0
        then FDetails.Add('  Bit $0002 = 0: Window is open')
        else FDetails.Add('x Bit $0002 = 1: Window is minimized');
      if w and $0008 = 0
        then FDetails.Add('  Bit $0008 = 0: Horizontal scrollbar hidden')
        else FDetails.Add('x Bit $0008 = 1: Horizontal scrollbar visible');
      if w and $0010 = 0
        then FDetails.Add('  Bit $0010 = 0: Vertical scrollbar hidden')
        else FDetails.Add('x Bit $0010 = 1: Vertical scrollbar visible');
      if w and $0020 = 0
        then FDetails.Add('  Bit $0020 = 0: Worksheet tab bar hidden')
        else FDetails.Add('x Bit $0020 = 1: Worksheet tab bar visible');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w, w]),
      'Option flags');

    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Index to active (displayed) worksheet');

    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Index of first visible tab in the worksheet tab bar');

    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Number of selected worksheets (highlighted in the worksheet tab bar)');

    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Width of worksheet tab bar (in 1/1000 of window width). '+
      'The remaining space is used by the horizontal scrollbar.');
  end;
end;


procedure TBIFFGrid.ShowWindow2;
var
  numBytes: Word;
  b: Byte;
  w: word;
  dw : DWord;
begin
  if FFormat = sfExcel2 then begin
    RowCount := FixedRows + 9;
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      '0 = Show formula results; 1 = Show formulas');
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      '0 = Do not show grid lines; 1 = Show grid lines');
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      '0 = Do not show sheet headers; 1 = Show sheet headers');
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      '0 = Panes are not frozen; 1 = Panes are frozen');
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      '0 = Show zero values as empty cells; 1 = Show zero values');
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Index to first visible row');
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Index to first visible column');
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b),
      '0 = Use manual grid line colour (below); 1 = Use automatic grid line colour');
    numbytes := 4;
    Move(FBuffer[FBufferIndex], dw, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [DWordLEToN(dw)]),
      'Grid line RGB color');
  end else begin
    RowCount := FixedRows + IfThen(FFormat = sfExcel5, 4, 8);
    numbytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('Option flags:');
      if w and $0001 = 0
        then FDetails.Add('  Bit $0001 = 0: Show formula results')
        else FDetails.Add('x Bit $0001 = 1: Show formulas');
      if w and $0002 = 0
        then FDetails.Add('  Bit $0002 = 0: Do not show grid lines')
        else FDetails.Add('x Bit $0002 = 1: Show grid lines');
      if w and $0004 = 0
        then FDetails.Add('  Bit $0004 = 0: Do not show sheet headers')
        else FDetails.Add('x Bit $0004 = 1: Show sheet headers');
      if w and $0008 = 0
        then FDetails.Add('  Bit $0008 = 0: Panes are not frozen')
        else FDetails.Add('x Bit $0008 = 1: Panes are frozen');
      if w and $0010 = 0
        then FDetails.Add('  Bit $0010 = 0: Show zero values as empty cells')
        else FDetails.Add('x Bit $0010 = 1: Show zero values');
      if w and $0020 = 0
        then FDetails.Add('  Bit $0020 = 0: Manual grid line color')
        else FDetails.Add('x Bit $0020 = 1: Automatic grid line color');
      if w and $0040 = 0
        then FDetails.Add('  Bit $0040 = 0: Columns from left to right')
        else FDetails.Add('x Bit $0040 = 1: Columns from right to left');
      if w and $0080 = 0
        then FDetails.Add('  Bit $0080 = 0: Do not show outline symbols')
        else FDetails.Add('x Bit $0080 = 1: Show outline symbols');
      if w and $0100 = 0
        then FDetails.Add('  Bit $0100 = 0: Keep splits if pane freeze is removed')
        else FDetails.Add('x Bit $0100 = 1: Remove splits if pane freeze is removed');
      if w and $0200 = 0
        then FDetails.Add('  Bit $0200 = 0: Sheet not selected')
        else FDetails.Add('x Bit $0200 = 1: Sheet selected');
      if w and $0400 = 0
        then FDetails.Add('  Bit $0400 = 0: Sheet not active')
        else FDetails.Add('x Bit $0400 = 1: Sheet active');
      if w and $0800 = 0
        then FDetails.Add('  Bit $0800 = 0: Show in normal view')
        else FDetails.Add('x Bit $0800 = 1: Show in page break preview');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w, w]),
      'Option flags');

    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Index to first visible row');
    Move(FBuffer[FBufferIndex], w, numbytes);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
      'Index to first visible column');

    if FFormat =sfExcel5 then begin
      numbytes := 4;
      Move(FBuffer[FBufferIndex], dw, numbytes);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [DWordLEToN(dw)]),
        'Grid line RGB color');
    end else begin
      numBytes := 2;
      Move(FBuffer[FBufferIndex], w, numbytes);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
        'Color index of grid line color');

      ShowInRow(FCurrRow, FBufferIndex, numbytes, '', 'Not used');

      Move(FBuffer[FBufferIndex], w, numbytes);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
        'Cached magnification factor in page break preview (in percent); 0 = Default (60%)');

      Move(FBuffer[FBufferIndex], w, numbytes);
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(WordLEToN(w)),
        'Cached magnification factor in normal view (in percent); 0 = Default (100%)');

      numBytes := 4;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, '', 'Not used');
    end;
  end;
end;


procedure TBIFFGrid.ShowWindowProtect;
var
  numBytes: Integer;
  w: Word;
begin
  RowCount := FixedRows + 1;
  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Protection state of workbook windows:'#13);
    if w = 0
      then FDetails.Add('0 = The workbook windows can be resized or moved and the window state can be changed.')
      else FDetails.Add('1 = The workbook windows cannot be resized or moved and the window state cannot be changed.');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
    'Protection state of the workbook windows');
end;


procedure TBIFFGrid.ShowWriteAccess;
var
  numbytes: Integer;
  s: String;
begin
  RowCount := FixedRows + 1;
  ExtractString(FBufferIndex, IfThen(FFormat=sfExcel8, 2, 1), FFormat=sfExcel8, s, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, s, 'User name (i.e., the name that you type when you install Microsoft Excel');
end;


procedure TBIFFGrid.ShowWriteProt;
begin
  RowCount := FixedRows + 2;
  ShowInRow(FCurrRow, FBufferIndex, 0, '', 'Write protect: if present file is write-protected');
  ShowInRow(FCurrRow, FBufferIndex, 0, '', 'Write protection password is in FILESHARING record');
end;


procedure TBIFFGrid.ShowXF;
var
  numBytes: Word;
  b: Byte;
  w: word;
  dw : DWord;
begin
  if FFormat = sfExcel2 then begin
    RowCount := FixedRows + 4;
    numBytes := 1;
    b := FBuffer[FBufferIndex];
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(b),
      'Index to font record');
    ShowInRow(FCurrRow, FBufferIndex, numBytes, '',
      '(not used)');

    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Number format and cell flags:'#13);
      FDetails.Add(Format('Bits 0-5 = %d: Index to FORMAT record', [b and $3F]));
      if b and $40 = 0
        then FDetails.Add('Bit 6    = 0: Cell is not locked')
        else FDetails.Add('Bit 6    = 1: Cell is locked');
      if b and $80 = 0
        then FDetails.Add('Bit 7    = 0: Formula is not hidden')
        else FDetails.Add('Bit 7    = 1: Formula is hidden');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('%d ($%.2x)', [b, b]),
      'Number format and cell flags');

    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Horizontal alignment, border style and background:'#13);
      case b and $07 of
        0: FDetails.Add('Bits $07 = 0: Horizontal alignment "General"');
        1: FDetails.Add('Bits $07 = 1: Horizontal alignment "Left"');
        2: FDetails.Add('Bits $07 = 2: Horizontal alignemnt "Centered"');
        3: FDetails.Add('Bits $07 = 3: Horizontal alignment "Right"');
        4: FDetails.Add('Bits $07 = 4: Horizontal alignment "Filled"');
        5: FDetails.Add('Bits $07 = 5: Horizontal alignment "Justified"');
        6: FDetails.Add('Bits $07 = 6: Horizontal alignment "Centred across selection"');
        7: FDetails.Add('Bits $07 = 7: Horizontal alignment "Distributed"');
      end;
      if b and $08 = 0
        then FDetails.Add('Bit $08 = 0: Cell has no left border')
        else FDetails.Add('Bit $08 = 1: Cell has left black border');
      if b and $10 = 0
        then FDetails.Add('Bit $10 = 0: Cell has no right border')
        else FDetails.Add('Bit $10 = 1: Cell has right black border');
      if b and $20 = 0
        then FDetails.Add('Bit $20 = 0: Cell has no top border')
        else FDetails.Add('Bit $20 = 1: Cell has top black border');
      if b and $40 = 0
        then FDetails.Add('Bit $40 = 0: Cell has no bottom border')
        else FDetails.Add('Bit $40 = 1: Cell has bottom black border');
      if b and $80 = 0
        then FDetails.Add('Bit $80 = 0: Cell has no shaded background')
        else FDetails.Add('Bit $80 = 1: Cell has shaded background');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
      'Horizontal alignment, border style, and background');
  end
  else
  begin // XF (BIFF5 and BIFF8)
    RowCount := FixedRows + IfThen(FFormat=sfExcel5, 7, 10);
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
      'Index to font record');
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    ShowInRow(FCurrRow, FBufferIndex, numBytes, IntToStr(WordLEToN(w)),
       'Index to format record');
    numBytes := 2;
    Move(FBuffer[FBufferIndex], w, numBytes);
    w := WordLEToN(w);
    if Row = FCurrRow then begin
      FDetails.Add('XFType, cell protection, parent style XF:'#13);
      if w and $0001 = 0
        then FDetails.Add('Bit $0001 = 0: Cell is not locked')
        else FDetails.Add('Bit $0001 = 1: Cell is locked');
      if w and $0002 = 0
        then FDetails.Add('Bit $0002 = 0: Formula is not hidden')
        else FDetails.Add('Bit $0002 = 1: Formula is hidden');
      if w and $0004 = 0
        then FDetails.Add('Bit $0004 = 0: Cell XF')
        else FDetails.Add('Bit $0004 = 1: Style XF');
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.4x', [w]),
      'XFType, cell protection, parent style XF');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    if Row = FCurrRow then begin
      FDetails.Add('Alignment and text break:'#13);
      case b and $03 of
        0: FDetails.Add('Bits 0-2 = 0: Horizontal alignment "General"');
        1: FDetails.Add('Bits 0-2 = 1: Horizontal alignment "Left"');
        2: FDetails.Add('Bits 0-2 = 2: Horizontal alignemnt "Centered"');
        3: FDetails.Add('Bits 0-2 = 3: Horizontal alignment "Right"');
        4: FDetails.Add('Bits 0-2 = 4: Horizontal alignment "Filled"');
        5: FDetails.Add('Bits 0-2 = 5: Horizontal alignment "Justified"');
        6: FDetails.Add('Bits 0-2 = 6: Horizontal alignment "Centred across selection"');
        7: if FFormat = sfExcel8 then
             FDetails.Add('x Bits 0-2 = 7: Horizontal alignment "Distributed"');
      end;
      if b and $08 = 0
        then FDetails.Add('Bit 3    = 0: Text is not wrapped.')
        else FDetails.Add('Bit 3    = 1: Text is wrapped at right border.');
      case (b and $70) shr 4 of
        0: FDetails.Add('Bits 4-6 = 0: Vertical alignment "Top"');
        1: FDetails.Add('Bits 4-6 = 1: Vertical alignment "Centered"');
        2: FDetails.Add('Bits 4-6 = 2: Vertical alignment "Bottom"');
        3: FDetails.Add('Bits 4-6 = 3: Vertical alignment "Justified"');
        4: if FFormat = sfExcel8 then
             FDetails.Add('Bits 4-6 = 4: Vertical alignment "Distributed"');
      end;
      if FFormat = sfExcel8 then begin
        if b and $80 = 0
          then FDetails.Add('Bit 3    = 0: Do NOT justify last line in justified or distibuted text')
          else FDetails.Add('Bit 3    = 1: Justify last line in justified or distibuted text');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
      'Alignment and text break');

    numBytes := 1;
    b := FBuffer[FBufferIndex];
    if FFormat = sfExcel5 then begin
      if Row = FCurrRow then begin
        FDetails.Add('Text orientation and flags for used attribute groups:'#13);
        case (b and $03) of
          0: FDetails.Add('Bits $03 = 0: not rotated');
          1: FDetails.Add('Bits $03 = 1: not rotated, letters stacked top-to-bottom');
          2: FDetails.Add('Bits $03 = 2: text rotated 90° counter-clockwise');
          3: FDetails.Add('Bits $03 = 3: text rotated 90° clockwise');
        end;
        if b and $04 = 0
          then FDetails.Add('Bit $04 = 0: No flag for number format')
          else FDetails.Add('Bit $04 = 1: Flag for number format');
        if b and $08 = 0
          then FDetails.Add('Bit $08 = 0: No flag for font')
          else FDetails.Add('Bit $08 = 2: Flag for font');
        if b and $10 = 0
          then FDetails.Add('Bit $10 = 0: No flag for hor/vert alignment, text wrap, indentation, orientation, rotation, and text direction')
          else FDetails.Add('Bit $10 = 1: Flag for hor/vert alignment, text wrap, indentation, orientation, rotation, and text direction');
        if b and $20 = 0
          then FDetails.Add('Bit $20 = 0: No flag for border lines')
          else FDetails.Add('Bit $20 = 1: Flag for border lines');
        if b and $40 = 0
          then FDetails.Add('Bit $40 = 0: No flag for background area style')
          else FDetails.Add('Bit $40 = 1: Flag for background area style');
        if b and $80 = 0
          then FDetails.Add('Bit $80 = 0: No flag for cell protection (cell locked and formula hidden)')
          else FDetails.Add('Bit $80 = 1: Flag for cell protection (cell locked and formula hidden)');
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [b]),
        'Text orientation and flags for used attribute groups');
    end else
    begin  // sfExcel8
      if Row = FCurrRow then begin
        FDetails.Add('Text rotation angle:'#13);
        if b = 0 then
          FDetails.Add('not rotated')
        else if b <= 90 then
          FDetails.Add(Format('%d degrees counter-clockwise', [b]))
        else if b <= 180 then
          FDetails.Add(Format('%d degrees clockwize', [b-90]))
        else if b = 255 then
          FDetails.Add('not rotated, letters stacked top-to-bottom');
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(b), 'Text rotation angle');
    end;

    if FFormat = sfExcel8 then begin
      numBytes := 1;
      b := FBuffer[FBufferIndex];
      if Row = FCurrRow then begin
        FDetails.Add('Indentation, shrink to cell size, and text direction:'#13);
        FDetails.Add(Format('Bits 0-3: Indent level = %d', [b and $0F]));
        if b and $10 = 0
          then FDetails.Add('Bit $10 = 0: Don''t shrink content to fit into cell')
          else FDetails.Add('Bit $10 = 1: Shrink content to fit into cell');
        if b and $20 = 0
          then FDetails.Add('Bit $20 = 0: Merge Cell option is OFF')
          else FDetails.Add('Bit $20 = 1: Merge Cell option is ON');
        case (b and $C0) shr 6 of
          0: FDetails.Add('Bits 6-7 = 0: Text direction according to context');
          1: FDetails.Add('Bits 6-7 = 1: Text direction left-to-right');
          2: FDetails.Add('Bits 6-7 = 2: Text direction right-to-left');
        end;
      end;
      ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.2x', [b]),
        'Indentation, shrink to cell size, and text direction');

      numBytes := 1;
      b := FBuffer[FBufferIndex];
      if Row = FCurrRow then begin
        FDetails.Add('Flags for used attribute groups:'#13);
        if b and $04 = 0
          then FDetails.Add('Bit $04 = 0: No flag for number format')
          else FDetails.Add('Bit $04 = 1: Flag for number format');
        if b and $08 = 0
          then FDetails.Add('Bit $08 = 0: No flag for font')
          else FDetails.Add('Bit $08 = 2: Flag for font');
        if b and $10 = 0
          then FDetails.Add('Bit $10 = 0: No flag for hor/vert alignment, text wrap, indentation, orientation, rotation, and text direction')
          else FDetails.Add('Bit $10 = 1: Flag for hor/vert alignment, text wrap, indentation, orientation, rotation, and text direction');
        if b and $20 = 0
          then FDetails.Add('Bit $20 = 0: No flag for border lines')
          else FDetails.Add('Bit $20 = 1: Flag for border lines');
        if b and $40 = 0
          then FDetails.Add('Bit $40 = 0: No flag for background area style')
          else FDetails.Add('Bit $40 = 1: Flag for background area style');
        if b and $80 = 0
          then FDetails.Add('Bit $80 = 0: No flag for cell protection (cell locked and formula hidden)')
          else FDetails.Add('Bit $80 = 1: Flag for cell protection (cell locked and formula hidden)');
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.2x', [b]),
        'Flags for used attribute groups');

      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numbytes);
      dw := DWordLEToN(dw);
      if Row = FCurrRow then begin
        FDetails.Add('Cell border lines and background area:'#13);
        case dw and $0000000F of
          $0000: FDetails.Add('Bits 0-3 = 0: Left border = No line');
          $0001: FDetails.Add('Bits 0-3 = 1: Left border = thin solid line');
          $0002: FDetails.Add('Bits 0-3 = 2: Left border = medium solid line');
          $0003: FDetails.Add('Bits 0-3 = 3: Left border = dashed line');
          $0004: FDetails.Add('Bits 0-3 = 4: Left border = dotted line');
          $0005: FDetails.Add('Bits 0-3 = 5: Left border = thick solid line');
          $0006: FDetails.Add('Bits 0-3 = 6: Left border = double solid line');
          $0007: FDetails.Add('Bits 0-3 = 7: Left border = hair line');
          $0008: FDetails.Add('Bits 0-3 = 8: Left border = medium dashed');
          $0009: FDetails.Add('Bits 0-3 = 9: Left border = thin dash-dotted');
          $000A: FDetails.Add('Bits 0-3 = 10: Left border = medium dash-dotted');
          $000B: FDetails.Add('Bits 0-3 = 11: Left border = thin dash-dot-dotted');
          $000C: FDetails.Add('Bits 0-3 = 12: Left border = medium dash-dot-dotted');
          $000D: FDetails.Add('Bits 0-3 = 13: Left border = slanted medium dash-dotted');
        end;
        case dw and $000000F0 of
          $0000: FDetails.Add('Bits 4-7 = 0: Right border = No line');
          $0010: FDetails.Add('Bits 4-7 = 1: Right border = thin solid line');
          $0020: FDetails.Add('Bits 4-7 = 2: Right border = medium solid line');
          $0030: FDetails.Add('Bits 4-7 = 3: Right border = dashed line');
          $0040: FDetails.Add('Bits 4-7 = 4: Right border = dotted line');
          $0050: FDetails.Add('Bits 4-7 = 5: Right border = thick solid line');
          $0060: FDetails.Add('Bits 4-7 = 6: Right border = double solid line');
          $0070: FDetails.Add('Bits 4-7 = 7: Right border = hair line');
          $0080: FDetails.Add('Bits 4-7 = 8: Right border = medium dashed');
          $0090: FDetails.Add('Bits 4-7 = 9: Right border = thin dash-dotted');
          $00A0: FDetails.Add('Bits 4-7 = 10: Right border = medium dash-dotted');
          $00B0: FDetails.Add('Bits 4-7 = 11: Right border = thin dash-dot-dotted');
          $00C0: FDetails.Add('Bits 4-7 = 12: Right border = medium dash-dot-dotted');
          $00D0: FDetails.Add('Bits 4-7 = 13: Right border = slanted medium dash-dotted');
        end;
        case dw and $00000F00 of
          $0000: FDetails.Add('Bits 8-11 = 0: Top border = No line');
          $0100: FDetails.Add('Bits 8-11 = 1: Top border = thin solid line');
          $0200: FDetails.Add('Bits 8-11 = 2: Top border = medium solid line');
          $0300: FDetails.Add('Bits 8-11 = 3: Top border = dashed line');
          $0400: FDetails.Add('Bits 8-11 = 4: Top border = dotted line');
          $0500: FDetails.Add('Bits 8-11 = 5: Top border = thick solid line');
          $0600: FDetails.Add('Bits 8-11 = 6: Top border = double solid line');
          $0700: FDetails.Add('Bits 8-11 = 7: Top border = hair line');
          $0800: FDetails.Add('Bits 8-11 = 8: Top border = medium dashed');
          $0900: FDetails.Add('Bits 8-11 = 9: Top border = thin dash-dotted');
          $0A00: FDetails.Add('Bits 8-11 = 10: Top border = medium dash-dotted');
          $0B00: FDetails.Add('Bits 8-11 = 11: Top border = thin dash-dot-dotted');
          $0C00: FDetails.Add('Bits 8-11 = 12: Top border = medium dash-dot-dotted');
          $0D00: FDetails.Add('Bits 8-11 = 13: Top border = slanted medium dash-dotted');
        end;
        case dw and $0000F000 of
          $0000: FDetails.Add('Bits 12-15 = 0: Bottom border = No line');
          $1000: FDetails.Add('Bits 12-15 = 1: Bottom border = thin solid line');
          $2000: FDetails.Add('Bits 12-15 = 2: Bottom border = medium solid line');
          $3000: FDetails.Add('Bits 12-15 = 3: Bottom border = dashed line');
          $4000: FDetails.Add('Bits 12-15 = 4: Bottom border = dotted line');
          $5000: FDetails.Add('Bits 12-15 = 5: Bottom border = thick solid line');
          $6000: FDetails.Add('Bits 12-15 = 6: Bottom border = double solid line');
          $7000: FDetails.Add('Bits 12-15 = 7: Bottom border = hair line');
          $8000: FDetails.Add('Bits 12-15 = 8: Bottom border = medium dashed');
          $9000: FDetails.Add('Bits 12-15 = 9: Bottom border = thin dash-dotted');
          $A000: FDetails.Add('Bits 12-15 = 10: Bottom border = medium dash-dotted');
          $B000: FDetails.Add('Bits 12-15 = 11: Bottom border = thin dash-dot-dotted');
          $C000: FDetails.Add('Bits 12-15 = 12: Bottom border = medium dash-dot-dotted');
          $D000: FDetails.Add('Bits 12-15 = 13: Bottom border = slanted medium dash-dotted');
        end;
        FDetails.Add(Format('Bits 16-22 = %d: Color index for left line color',  [(dw and $007F0000) shr 16]));
        FDetails.Add(Format('Bits 23-29 = %d: Color index for right line color', [(dw and $3F800000) shr 23]));
        if dw and $40000000 = 0
          then FDetails.Add('Bit 30 = 0: No diagonal line from top left to right bottom')
          else FDetails.Add('Bit 30 = 1: Diagonal line from top left to right bottom');
        if dw and $80000000 = 0
          then FDetails.Add('Bit 31 = 0: No diagonal line from bottom left to right top')
          else FDetails.Add('Bit 31 = 1: Diagonal line from bottom left to right top');
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [dw]),
        'Cell border lines and background area');

      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numbytes);
      if Row = FCurrRow then begin
        FDetails.Add('Cell border lines and background area (cont''d):'#13);
        FDetails.Add(Format('Bits 0-6 = %d: Color index for top line color',  [(dw and $0000007F)]));
        FDetails.Add(Format('Bits 7-13 = %d: Color index for bottom line color', [(dw and $00003F80) shr 7]));
        FDetails.Add(Format('Bits 14-20 = %d: Color index for diagonal line color', [(dw and $001FC000) shr 14]));
        case dw and $01E00000 shr 17 of
          $0: FDetails.Add('Bits 21-24 = 0: Diagonal line style = No line');
          $1: FDetails.Add('Bits 21-24 = 1: Diagonal line style = thin solid line');
          $2: FDetails.Add('Bits 21-24 = 2: Diagonal line style = medium solid line');
          $3: FDetails.Add('Bits 21-24 = 3: Diagonal line style = dashed line');
          $4: FDetails.Add('Bits 21-24 = 4: Diagonal line style = dotted line');
          $5: FDetails.Add('Bits 21-24 = 5: Diagonal line style = thick solid line');
          $6: FDetails.Add('Bits 21-24 = 6: Diagonal line style = double solid line');
          $7: FDetails.Add('Bits 21-24 = 7: Diagonal line style = hair line');
          $8: FDetails.Add('Bits 21-24 = 8: Diagonal line style = medium dashed');
          $9: FDetails.Add('Bits 21-24 = 9: Diagonal line style = thin dash-dotted');
          $A: FDetails.Add('Bits 21-24 = 10: Diagonal line style = medium dash-dotted');
          $B: FDetails.Add('Bits 21-24 = 11: Diagonal line style = thin dash-dot-dotted');
          $C: FDetails.Add('Bits 21-24 = 12: Diagonal line style = medium dash-dot-dotted');
          $D: FDetails.Add('Bits 21-24 = 13: Diagonal line style = slanted medium dash-dotted');
        end;
        case (dw and $FC000000) shr 26 of
          $00: FDetails.Add('Bits 26-31 = 0: Fill pattern = No fill');
          $01: FDetails.Add('Bits 26-31 = 1: Fill pattern = solid fill');
          $02: FDetails.Add('Bits 26-31 = 2: Fill pattern = medium fill');
          $03: FDetails.Add('Bits 26-31 = 3: Fill pattern = dense fill');
          $04: FDetails.Add('Bits 26-31 = 4: Fill pattern = sparse fill');
          $05: FDetails.Add('Bits 26-31 = 5: Fill pattern = horizontal fill');
          $06: FDetails.Add('Bits 26-31 = 6: Fill pattern = vertical fill');
          $07: FDetails.Add('Bits 26-31 = 7: Fill pattern = backslash fill');
          $08: FDetails.Add('Bits 26-31 = 8: Fill pattern = slash fill');
          $09: FDetails.Add('Bits 26-31 = 9: Fill pattern = coarse medium fill');
          $0A: FDetails.Add('Bits 26-31 = 10: Fill pattern = coarse medium horiz fill');
          $0B: FDetails.Add('Bits 26-31 = 11: Fill pattern = sparse horizontal fill');
          $0C: FDetails.Add('Bits 26-31 = 12: Fill pattern = sparse vertical fill');
          $0D: FDetails.Add('Bits 26-31 = 13: Fill pattern = sparse backslash fill');
          $0E: FDetails.Add('Bits 26-31 = 14: Fill pattern = sparse slash fill');
          $0F: FDetails.Add('Bits 26-31 = 15: Fill pattern = cross fill');
          $10: FDetails.Add('Bits 26-31 = 16: Fill pattern = dense backslash fill');
          $11: FDetails.Add('Bits 26-31 = 17: Fill pattern = very sparse fill');
          $12: FDetails.Add('Bits 26-31 = 18: Fill pattern = extremely sparse fill');
        end;
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [dw]),
        'Cell border lines and background area (cont''d)');

      numBytes := 2;
      Move(FBuffer[FBufferIndex], w, numbytes);
      w := WordLEToN(w);
      if Row = FCurrRow then begin
        FDetails.Add('Cell border lines and background area (cont''d):'#13);
        FDetails.Add(Format('Bits 0-6 = %d: Color index for pattern color',  [(w and $007F)]));
        FDetails.Add(Format('Bits 7-13 = %d: Color index for pattern background color', [(w and $3F80) shr 7]));
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
        'Cell border lines and background area (cont''d)');
    end;

    if FFormat = sfExcel5 then begin
      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numBytes);
      dw := DWordLEToN(dw);
      if Row = FCurrRow then begin
        FDetails.Add('Cell border lines and background area:'#13);
        FDetails.Add(Format('Bits 0-6 = %d: Color index for pattern color',  [(dw and $007F)]));
        FDetails.Add(Format('Bits 7-13 = %d: Color index for pattern background color', [(dw and $3F80) shr 7]));
        case (dw and $003F0000) shr 16 of
          $00: FDetails.Add('Bits 16-21 = 0: Fill pattern = No fill');
          $01: FDetails.Add('Bits 16-21 = 1: Fill pattern = solid fill');
          $02: FDetails.Add('Bits 16-21 = 2: Fill pattern = medium fill');
          $03: FDetails.Add('Bits 16-21 = 3: Fill pattern = dense fill');
          $04: FDetails.Add('Bits 16-21 = 4: Fill pattern = sparse fill');
          $05: FDetails.Add('Bits 16-21 = 5: Fill pattern = horizontal fill');
          $06: FDetails.Add('Bits 16-21 = 6: Fill pattern = vertical fill');
          $07: FDetails.Add('Bits 16-21 = 7: Fill pattern = backslash fill');
          $08: FDetails.Add('Bits 16-21 = 8: Fill pattern = slash fill');
          $09: FDetails.Add('Bits 16-21 = 9: Fill pattern = coarse medium fill');
          $0A: FDetails.Add('Bits 16-21 = 10: Fill pattern = coarse medium horiz fill');
          $0B: FDetails.Add('Bits 16-21 = 11: Fill pattern = sparse horizontal fill');
          $0C: FDetails.Add('Bits 16-21 = 12: Fill pattern = sparse vertical fill');
          $0D: FDetails.Add('Bits 16-21 = 13: Fill pattern = sparse backslash fill');
          $0E: FDetails.Add('Bits 16-21 = 14: Fill pattern = sparse slash fill');
          $0F: FDetails.Add('Bits 16-21 = 15: Fill pattern = cross fill');
          $10: FDetails.Add('Bits 16-21 = 16: Fill pattern = dense backslash fill');
          $11: FDetails.Add('Bits 16-21 = 17: Fill pattern = very sparse fill');
          $12: FDetails.Add('Bits 16-21 = 18: Fill pattern = extremely sparse fill');
        end;
        case dw and $01C00000 shr 22 of
          $0: FDetails.Add('Bits 22-24 = 0: Bottom line style = No line');
          $1: FDetails.Add('Bits 22-24 = 1: Bottom line style = thin solid line');
          $2: FDetails.Add('Bits 22-24 = 2: Bottom line style = medium solid line');
          $3: FDetails.Add('Bits 22-24 = 3: Bottom line style = dashed line');
          $4: FDetails.Add('Bits 22-24 = 4: Bottom line style = dotted line');
          $5: FDetails.Add('Bits 22-24 = 5: Bottom line style = thick solid line');
          $6: FDetails.Add('Bits 22-24 = 6: Bottom line style = double solid line');
          $7: FDetails.Add('Bits 22-24 = 7: Bottom line style = hair line');
        end;
        FDetails.Add(Format('Bits 25-31 = %d: Color index for bottom line color', [(dw and $FE000000) shr 25]));
      end;
      ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.8x', [dw]),
        'Cell border lines & background area');

      numBytes := 4;
      Move(FBuffer[FBufferIndex], dw, numBytes);
      dw := DWOrdLEToN(dw);
      if Row = FCurrRow then begin
        FDetails.Add('Cell border lines (cont''d):'#13);
        case dw and $00000007 of
          $00: FDetails.Add('Bits 0-2 = 0: Top border = No line');
          $01: FDetails.Add('Bits 0-2 = 1: Top border = thin solid line');
          $02: FDetails.Add('Bits 0-2 = 2: Top border = medium solid line');
          $03: FDetails.Add('Bits 0-2 = 3: Top border = dashed line');
          $04: FDetails.Add('Bits 0-2 = 4: Top border = dotted line');
          $05: FDetails.Add('Bits 0-2 = 5: Top border = thick solid line');
          $06: FDetails.Add('Bits 0-2 = 6: Top border = double solid line');
          $07: FDetails.Add('Bits 0-2 = 7: Top border = hair line');
        end;
        case (dw and $00000038) shr 3 of
          $0000: FDetails.Add('Bits 3-5 = 0: Left border = No line');
          $0001: FDetails.Add('Bits 3-5 = 1: Left border = thin solid line');
          $0002: FDetails.Add('Bits 3-5 = 2: Left border = medium solid line');
          $0003: FDetails.Add('Bits 3-5 = 3: Left border = dashed line');
          $0004: FDetails.Add('Bits 3-5 = 4: Left border = dotted line');
          $0005: FDetails.Add('Bits 3-5 = 5: Left border = thick solid line');
          $0006: FDetails.Add('Bits 3-5 = 6: Left border = double solid line');
          $0007: FDetails.Add('Bits 3-5 = 7: Left border = hair line');
        end;
        case (dw and $000001C0) shr 6 of
          $0000: FDetails.Add('Bits 6-8 = 0: Right border = No line');
          $0010: FDetails.Add('Bits 6-8 = 1: Right border = thin solid line');
          $0020: FDetails.Add('Bits 6-8 = 2: Right border = medium solid line');
          $0030: FDetails.Add('Bits 6-8 = 3: Right border = dashed line');
          $0040: FDetails.Add('Bits 6-8 = 4: Right border = dotted line');
          $0050: FDetails.Add('Bits 6-8 = 5: Right border = thick solid line');
          $0060: FDetails.Add('Bits 6-8 = 6: Right border = double solid line');
          $0070: FDetails.Add('Bits 6-8 = 7: Right border = hair line');
        end;
        FDetails.Add(Format('Bits 9-15 = %d: Color index for top line color', [(dw and $0000FE00) shr 7]));
        FDetails.Add(Format('Bits 16-22 = %d: Color index for left line color', [(dw and $007F0000) shr 16]));
        FDetails.Add(Format('Bits 23-29 = %d: Color index for right line color', [(dw and $3F800000) shr 23]));
      end;
      ShowInRow(FCurrRow, FBufferIndex, numBytes, Format('$%.8x', [dw]),
        'Cell border lines (cont''d)');
    end;
  end;
end;


procedure TBIFFGrid.ShowXFCRC;
var
  numBytes: Integer;
  w: Word;
  dw: DWord;
begin
  RowCount := FixedRows + 7;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [wordLEToN(w)]),
    'Future record type');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Attributes:'#13);
    if w and $0001 = 0
      then FDetails.Add('  Bit 0 = 0: The containing record does not specify a range of cells.')
      else FDetails.Add('x Bit 0 = 1: The containing record specifies a range of cells.');
    FDetails.Add('Bit 1: specifies wether to alert the user of possible problems '+
      'when saving the file whithout having reckognized this record.');
    FDetails.Add('Bits 2-15: reserved (MUST be zero, MUST be ignored)');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
    'Attributes');

  numbytes := 4;
  Move(FBuffer[FBufferIndex], dw, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)), 'Reserved');
  Move(FBuffer[FBufferIndex], dw, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)), 'Reserved');

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]), 'Reserved');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Count of XF records');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w,w]),
    'Checksum of XF records');
end;


procedure TBIFFGrid.ShowXFEXT;
var
  numBytes: Integer;
  w: Word;
  dw: DWord;
  i, n: Integer;
  et: Word;
  es: Word;
  ct: Word;
  buffidx: Cardinal;
  s: String;
begin
  BeginUpdate;

  RowCount := FixedRows + 100;

  numBytes := 2;
  Move(FBuffer[FBufferIndex], w, numBytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [wordLEToN(w)]),
    'Future record type');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  if Row = FCurrRow then begin
    FDetails.Add('Attributes:'#13);
    if w and $0001 = 0
      then FDetails.Add('  Bit 0 = 0: The containing record does not specify a range of cells.')
      else FDetails.Add('x Bit 0 = 1: The containing record specifies a range of cells.');
    FDetails.Add('Bit 1: specifies wether to alert the user of possible problems '+
      'when saving the file whithout having reckognized this record.');
    FDetails.Add('Bits 2-15: reserved (MUST be zero, MUST be ignored)');
  end;
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]),
    'Attributes');

  numbytes := 4;
  Move(FBuffer[FBufferIndex], dw, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)), 'Reserved');
  Move(FBuffer[FBufferIndex], dw, numbytes);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(DWordLEToN(dw)), 'Reserved');

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]), 'Reserved');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w,w]),
    'XF index');

  numbytes := 2;
  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('$%.4x', [w]), 'Reserved');

  Move(FBuffer[FBufferIndex], w, numbytes);
  w := WordLEToN(w);
  ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
    'Number of extension properties');
  n := w;

  for i:=1 to n do begin
    buffidx := FBufferIndex;
    numbytes := 2;
    Move(FBuffer[FBufferIndex], et, numbytes);
    et := WordLEToN(et);
    if Row = FCurrRow then begin
      FDetails.Add('Type:'#13);
      case et of
        $04: FDetails.Add('Full color extension that specifies the cell interior foreground color.');
        $05: FDetails.Add('Full color extension that specifies the cell interior background color.');
        $06: FDetails.Add('Gradient extension that specifies a cell interior gradient fill.');
        $07: FDetails.Add('Full color extension that specifies the top cell border color.');
        $08: FDetails.Add('Full color extension that specifies the bottom cell border color.');
        $09: FDetails.Add('Full color extension that specifies the left cell border color.');
        $0A: FDetails.Add('Full color extension that specifies the right cell border color.');
        $0B: FDetails.Add('Full color extension that specifies the diagonal cell border color.');
        $0D: FDetails.Add('Full color extension that specifies the cell text color.');
        $0E: FDetails.Add('2-byte unsigned integer that specifies a font scheme.');
        $0F: FDetails.Add('2-byte unsigned integer that specifies the text indentation level (MUST be <= 250).');
      end;
    end;
    ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [et, et]),
      Format('Extension property #%d: Type', [i]));
    Move(FBuffer[FBufferIndex], es, numbytes);
    es := WordLEToN(es);
    ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(es),
      Format('Extension property #%d: Data size', [i]));

    case et of
      $04, $05, $07..$0D:  // FullColorExt
        begin
          numbytes := 2;
          Move(FBuffer[FBufferIndex], ct, numbytes);
          ct := WordLEToN(ct);
          if Row = FCurrRow then begin
            FDetails.Add('Full color extension - Color type:'#13);
            case ct of
              0: FDetails.Add('0 - Automatic color');
              1: FDetails.Add('1 - Indexed color');
              2: FDetails.Add('2 - RGB color');
              3: FDetails.Add('3 - Theme color');
              4: FDetails.Add('4 - Color not set');
            end;
          end;
          ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(ct),
            Format('Extension property #%d (Full color extension): Color type', [i]));
          numbytes := 2;
          Move(FBuffer[FBufferIndex], w, numbytes);
          w := WordLEToN(w);
          ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(integer(w)),
            Format('Extension property #%d (Full color extension): Color tint', [i]));
          numbytes := 4;
          Move(FBuffer[FBufferIndex], dw, numbytes);
          dw := DWordLEToN(dw);
          case ct of
            0: s := '(dummy - MUST be 0)';
            1: s := '(index)';
            2: s := '(RGB value)';
            3: s := '(theme)';
            else s := '';
          end;
          ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.8x)', [QWord(dw), QWord(dw)]),
            Format('Extension property #%d (Full color extension): value %s', [i, s]));
          numbytes := 4;
          Move(FBuffer[FBufferIndex], dw, numbytes);
          dw := DWordLEToN(dw);
          ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.8x)', [QWord(dw), QWord(dw)]),
            Format('Extension property #%d (Full color extension): Reserved', [i]));
          Move(FBuffer[FBufferIndex], dw, numbytes);
          dw := DWordLEToN(dw);
          ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.8x)', [QWord(dw), QWord(dw)]),
            Format('Extension property #%d (Full color extension): Reserved', [i]));
        end;

      $06:  // Gradient
        begin
          ShowInRow(FCurrRow, FBufferIndex, es, '(var)',
            Format('Extension property #%d (Gradient): - not interpreted here -', [i]));
        end;

      $0E:  // Font scheme
        begin
          numbytes := 2;
          Move(FBuffer[FBufferIndex], w, numbytes);
          w := WordLEToN(w);
          if Row = FCurrRow then begin
            FDetails.Add('Font scheme:'#13);
            case w of
              0: FDetails.Add('0 - No font scheme');
              1: FDetails.Add('1 - Major scheme');
              2: FDetails.Add('2 - Minor scheme');
              3: FDetails.Add('3 - Ninched scheme');
            end;
          end;
          ShowInRow(FCurrRow, FBufferIndex, numbytes, Format('%d ($%.4x)', [w,w]),
            Format('Extension property #%d Font scheme', [i]));
        end;

      $0F:   // Text indentation level
        begin
          numbytes := 2;
          Move(FBuffer[FBufferIndex], w, numbytes);
          w := WordLEToN(w);
          ShowInRow(FCurrRow, FBufferIndex, numbytes, IntToStr(w),
            Format('Extension property #%d Text indentation level', [i]));
        end;
    end;
    FBufferIndex := buffidx + es;
  end;
  RowCount := FCurrRow;

  EndUpdate(true);
end;


end.

