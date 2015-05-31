{ fpspreadsheet }

{@@ ----------------------------------------------------------------------------
  Unit fpspreadsheet reads and writes spreadsheet documents.

  AUTHORS: Felipe Monteiro de Carvalho, Reinier Olislagers, Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpspreadsheet;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

{$include fps.inc}

interface

uses
 {$ifdef UNIX}{$ifndef DARWIN}{$ifndef FPS_DONT_USE_CLOCALE}
  clocale,
 {$endif}{$endif}{$endif}
  Classes, SysUtils, fpimage, AVL_Tree, avglvltree, lconvencoding,
  fpsTypes, fpsClasses;

type
  { Forward declarations }
  TsWorksheet = class;
  TsWorkbook = class;
  TsBasicSpreadReader = class;
  TsBasicSpreadWriter = class;

  {@@ The record TRow contains information about a spreadsheet row:
    @param Row   The index of the row (beginning with 0)
    @param Height  The height of the row (expressed as lines count of the default font)
   Only rows with heights that cannot be derived from the font height have a
   row record. }
  TRow = record
    Row: Cardinal;
    Height: Single;  // in "lines"
  end;

  {@@ Pointer to a TRow record }
  PRow = ^TRow;

  {@@ The record TCol contains information about a spreadsheet column:
   @param Col    The index of the column (beginning with 0)
   @param Width  The width of the column (expressed in character count of the "0" character of the default font.
   Only columns with non-default widths have a column record. }
  TCol = record
    Col: Cardinal;
    Width: Single; // in "characters". Excel uses the width of char "0" in 1st font
  end;

  {@@ Pointer to a TCol record }
  PCol = ^TCol;

  {@@ Worksheet user interface options:
    @param soShowGridLines  Show or hide the grid lines in the spreadsheet
    @param soShowHeaders    Show or hide the column or row headers of the spreadsheet
    @param soHasFrozenPanes If set a number of rows and columns of the spreadsheet
                            is fixed and does not scroll. The number is defined by
                            LeftPaneWidth and TopPaneHeight. }
  TsSheetOption = (soShowGridLines, soShowHeaders, soHasFrozenPanes);

  {@@ Set of user interface options
    @ see TsSheetOption }
  TsSheetOptions = set of TsSheetOption;


  { TsWorksheet }

  {@@ This event fires whenever a cell value or cell formatting changes. It is
    handled by TsWorkbookLink to update the listening controls. }
  TsCellEvent = procedure (Sender: TObject; ARow, ACol: Cardinal) of object;

  {@@ This event can be used to override the built-in comparing function which
    is called when cells are sorted. }
  TsCellCompareEvent = procedure (Sender: TObject; ACell1, ACell2: PCell;
    var AResult: Integer) of object;

  {@@ The worksheet contains a list of cells and provides a variety of methods
    to read or write data to the cells, or to change their formatting. }
  TsWorksheet = class
  private
    FWorkbook: TsWorkbook;
    FName: String;  // Name of the worksheet (displayed at the tab)
    FCells: TsCells;
    FComments: TsComments;
    FMergedCells: TsMergedCells;
    FHyperlinks: TsHyperlinks;
    FRows, FCols: TIndexedAVLTree; // This lists contain only rows or cols with styles different from default
    FActiveCellRow: Cardinal;
    FActiveCellCol: Cardinal;
    FSelection: TsCellRangeArray;
    FLeftPaneWidth: Integer;
    FTopPaneHeight: Integer;
    FOptions: TsSheetOptions;
    FFirstRowIndex: Cardinal;
    FFirstColIndex: Cardinal;
    FLastRowIndex: Cardinal;
    FLastColIndex: Cardinal;
    FDefaultColWidth: Single;   // in "characters". Excel uses the width of char "0" in 1st font
    FDefaultRowHeight: Single;  // in "character heights", i.e. line count
    FSortParams: TsSortParams;  // Parameters of the current sorting operation
    FOnChangeCell: TsCellEvent;
    FOnChangeFont: TsCellEvent;
    FOnCompareCells: TsCellCompareEvent;
    FOnSelectCell: TsCellEvent;

    { Setter/Getter }
    function GetFormatSettings: TFormatSettings;
    procedure SetName(const AName: String);

    { Callback procedures called when iterating through all cells }
    procedure DeleteColCallback(data, arg: Pointer);
    procedure DeleteRowCallback(data, arg: Pointer);
    procedure InsertColCallback(data, arg: Pointer);
    procedure InsertRowCallback(data, arg: Pointer);

  protected
    function CellUsedInFormula(ARow, ACol: Cardinal): Boolean;

    // Remove and delete cells
    function RemoveCell(ARow, ACol: Cardinal): PCell;
    procedure RemoveAndFreeCell(ARow, ACol: Cardinal);

    // Sorting
    function DoCompareCells(ARow1, ACol1, ARow2, ACol2: Cardinal;
      ASortOptions: TsSortOptions): Integer;
    function DoInternalCompareCells(ACell1, ACell2: PCell;
      ASortOptions: TsSortOptions): Integer;
    procedure DoExchangeColRow(AIsColumn: Boolean; AIndex, WithIndex: Cardinal;
      AFromIndex, AToIndex: Cardinal);
    procedure ExchangeCells(ARow1, ACol1, ARow2, ACol2: Cardinal);

  public
    {@@ Page layout parameters for printing }
    PageLayout: TsPageLayout;

    { Base methods }
    constructor Create;
    destructor Destroy; override;

    { Utils }
    class function CellInRange(ARow, ACol: Cardinal; ARange: TsCellRange): Boolean;
    class function CellPosToText(ARow, ACol: Cardinal): string;
//    procedure RemoveAllCells;
    procedure UpdateCaches;

    { Reading of values }
    function  ReadAsUTF8Text(ARow, ACol: Cardinal): string; overload;
    function  ReadAsUTF8Text(ACell: PCell): string; overload;
    function  ReadAsUTF8Text(ACell: PCell; AFormatSettings: TFormatSettings): string; overload;
    function  ReadAsNumber(ARow, ACol: Cardinal): Double; overload;
    function  ReadAsNumber(ACell: PCell): Double; overload;
    function  ReadAsDateTime(ARow, ACol: Cardinal; out AResult: TDateTime): Boolean; overload;
    function  ReadAsDateTime(ACell: PCell; out AResult: TDateTime): Boolean; overload;
    function  ReadFormulaAsString(ACell: PCell; ALocalized: Boolean = false): String;
    function  ReadNumericValue(ACell: PCell; out AValue: Double): Boolean;

    { Reading of cell attributes }
    function GetDisplayedDecimals(ACell: PCell): Byte;
    function GetNumberFormatAttributes(ACell: PCell; out ADecimals: Byte;
      out ACurrencySymbol: String): Boolean;

    function  ReadUsedFormatting(ACell: PCell): TsUsedFormattingFields;
    function  ReadBackground(ACell: PCell): TsFillPattern;
    function  ReadBackgroundColor(ACell: PCell): TsColor;
    function  ReadCellBorders(ACell: PCell): TsCellBorders;
    function  ReadCellBorderStyle(ACell: PCell; ABorder: TsCellBorder): TsCellBorderStyle;
    function  ReadCellBorderStyles(ACell: PCell): TsCellBorderStyles;
    function  ReadCellFont(ACell: PCell): TsFont;
    function  ReadCellFormat(ACell: PCell): TsCellFormat;
    function  ReadHorAlignment(ACell: PCell): TsHorAlignment;
    procedure ReadNumFormat(ACell: PCell; out ANumFormat: TsNumberFormat;
      out ANumFormatStr: String);
    function  ReadTextRotation(ACell: PCell): TsTextRotation;
    function  ReadVertAlignment(ACell: PCell): TsVertAlignment;
    function  ReadWordwrap(ACell: PCell): boolean;

    { Writing of values }
    function WriteBlank(ARow, ACol: Cardinal): PCell; overload;
    procedure WriteBlank(ACell: PCell); overload;

    function WriteBoolValue(ARow, ACol: Cardinal; AValue: Boolean): PCell; overload;
    procedure WriteBoolValue(ACell: PCell; AValue: Boolean); overload;

    function WriteCellValueAsString(ARow, ACol: Cardinal; AValue: String): PCell; overload;
    procedure WriteCellValueAsString(ACell: PCell; AValue: String); overload;

    function WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
      ANumFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = 2;
      ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
      ANegCurrFormat: Integer = -1): PCell; overload;
    procedure WriteCurrency(ACell: PCell; AValue: Double;
      ANumFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = -1;
      ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
      ANegCurrFormat: Integer = -1); overload;
    function WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
      ANumFormat: TsNumberFormat; ANumFormatString: String): PCell; overload;
    procedure WriteCurrency(ACell: PCell; AValue: Double;
      ANumFormat: TsNumberFormat; ANumFormatString: String); overload;

    function WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime): PCell; overload;
    procedure WriteDateTime(ACell: PCell; AValue: TDateTime); overload;
    function WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
      ANumFormat: TsNumberFormat; ANumFormatStr: String = ''): PCell; overload;
    procedure WriteDateTime(ACell: PCell; AValue: TDateTime;
      ANumFormat: TsNumberFormat; ANumFormatStr: String = ''); overload;
    function WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
      ANumFormatStr: String): PCell; overload;
    procedure WriteDateTime(ACell: PCell; AValue: TDateTime;
      ANumFormatStr: String); overload;

    function WriteErrorValue(ARow, ACol: Cardinal; AValue: TsErrorValue): PCell; overload;
    procedure WriteErrorValue(ACell: PCell; AValue: TsErrorValue); overload;

    function WriteFormula(ARow, ACol: Cardinal; AFormula: String;
      ALocalized: Boolean = false): PCell; overload;
    procedure WriteFormula(ACell: PCell; AFormula: String;
      ALocalized: Boolean = false); overload;

    function WriteNumber(ARow, ACol: Cardinal; ANumber: double): PCell; overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double); overload;
    function WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      ANumFormat: TsNumberFormat; ADecimals: Byte = 2): PCell; overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double;
      ANumFormat: TsNumberFormat; ADecimals: Byte = 2); overload;
    function WriteNumber(ARow, ACol: Cardinal; ANumber: double;
      ANumFormat: TsNumberFormat; ANumFormatString: String): PCell; overload;
    procedure WriteNumber(ACell: PCell; ANumber: Double;
      ANumFormat: TsNumberFormat; ANumFormatString: String); overload;

    function WriteRPNFormula(ARow, ACol: Cardinal;
      AFormula: TsRPNFormula): PCell; overload;
    procedure WriteRPNFormula(ACell: PCell;
      AFormula: TsRPNFormula); overload;

    function WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring): PCell; overload;
    procedure WriteUTF8Text(ACell: PCell; AText: ansistring); overload;

    { Writing of cell attributes }
    function WriteBackground(ARow, ACol: Cardinal; AStyle: TsFillStyle;
      APatternColor: TsColor = scTransparent;
      ABackgroundColor: TsColor = scTransparent): PCell; overload;
    procedure WriteBackground(ACell: PCell; AStyle: TsFillStyle;
      APatternColor: TsColor = scTransparent;
      ABackgroundColor: TsColor = scTransparent); overload;
    function WriteBackgroundColor(ARow, ACol: Cardinal; AColor: TsColor): PCell; overload;
    procedure WriteBackgroundColor(ACell: PCell; AColor: TsColor); overload;

    function WriteBorderColor(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      AColor: TsColor): PCell; overload;
    procedure WriteBorderColor(ACell: PCell; ABorder: TsCellBorder;
      AColor: TsColor); overload;
    function WriteBorderLineStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle): PCell; overload;
    procedure WriteBorderLineStyle(ACell: PCell; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle); overload;
    function WriteBorders(ARow, ACol: Cardinal;
      ABorders: TsCellBorders): PCell; overload;
    procedure WriteBorders(ACell: PCell; ABorders: TsCellBorders); overload;
    function WriteBorderStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      AStyle: TsCellBorderStyle): PCell; overload;
    procedure WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
      AStyle: TsCellBorderStyle); overload;
    function WriteBorderStyle(ARow, ACol: Cardinal; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle; AColor: TsColor): PCell; overload;
    procedure WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
      ALineStyle: TsLineStyle; AColor: TsColor); overload;
    function WriteBorderStyles(ARow, ACol: Cardinal;
      const AStyles: TsCellBorderStyles): PCell; overload;
    procedure WriteBorderStyles(ACell: PCell;
      const AStyles: TsCellBorderStyles); overload;

    procedure WriteCellFormat(ACell: PCell; const ACellFormat: TsCellFormat);

    function WriteDateTimeFormat(ARow, ACol: Cardinal; ANumFormat: TsNumberFormat;
      const ANumFormatString: String = ''): PCell; overload;
    procedure WriteDateTimeFormat(ACell: PCell; ANumFormat: TsNumberFormat;
      const ANumFormatString: String = ''); overload;

    function WriteDecimals(ARow, ACol: Cardinal; ADecimals: byte): PCell; overload;
    procedure WriteDecimals(ACell: PCell; ADecimals: Byte); overload;

    function  WriteFont(ARow, ACol: Cardinal; const AFontName: String;
      AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer; overload;
    function  WriteFont(ACell: PCell; const AFontName: String;
      AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer; overload;
    function WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer): PCell; overload;
    procedure WriteFont(ACell: PCell; AFontIndex: Integer); overload;
    function WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer; overload;
    function WriteFontColor(ACell: PCell; AFontColor: TsColor): Integer; overload;
    function WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer; overload;
    function WriteFontName(ACell: PCell; AFontName: String): Integer; overload;
    function WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer; overload;
    function WriteFontSize(ACell: PCell; ASize: Single): Integer; overload;
    function WriteFontStyle(ARow, ACol: Cardinal; AStyle: TsFontStyles): Integer; overload;
    function WriteFontStyle(ACell: PCell; AStyle: TsFontStyles): Integer; overload;

    function WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment): PCell; overload;
    procedure WriteHorAlignment(ACell: PCell; AValue: TsHorAlignment); overload;

    function WriteNumberFormat(ARow, ACol: Cardinal; ANumFormat: TsNumberFormat;
      const ANumFormatString: String = ''): PCell; overload;
    procedure WriteNumberFormat(ACell: PCell; ANumFormat: TsNumberFormat;
      const ANumFormatString: String = ''); overload;
    function WriteNumberFormat(ARow, ACol: Cardinal; ANumFormat: TsNumberFormat;
      ADecimals: Integer; ACurrencySymbol: String = ''; APosCurrFormat: Integer = -1;
      ANegCurrFormat: Integer = -1): PCell; overload;
    procedure WriteNumberFormat(ACell: PCell; ANumFormat: TsNumberFormat;
      ADecimals: Integer; ACurrencySymbol: String = '';
      APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1); overload;
    function WriteFractionFormat(ARow, ACol: Cardinal; AMixedFraction: Boolean;
      ANumeratorDigits, ADenominatorDigits: Integer): PCell; overload;
    procedure WriteFractionFormat(ACell: PCell; AMixedFraction: Boolean;
      ANumeratorDigits, ADenominatorDigits: Integer); overload;

    function WriteTextRotation(ARow, ACol: Cardinal; ARotation: TsTextRotation): PCell; overload;
    procedure WriteTextRotation(ACell: PCell; ARotation: TsTextRotation); overload;

    function WriteUsedFormatting(ARow, ACol: Cardinal;
      AUsedFormatting: TsUsedFormattingFields): PCell; overload;
    procedure WriteUsedFormatting(ACell: PCell;
      AUsedFormatting: TsUsedFormattingFields); overload;

    function WriteVertAlignment(ARow, ACol: Cardinal; AValue: TsVertAlignment): PCell; overload;
    procedure WriteVertAlignment(ACell: PCell; AValue: TsVertAlignment); overload;

    function WriteWordwrap(ARow, ACol: Cardinal; AValue: boolean): PCell; overload;
    procedure WriteWordwrap(ACell: PCell; AValue: boolean); overload;

    { Formulas }
    function BuildRPNFormula(ACell: PCell; ADestCell: PCell = nil): TsRPNFormula;
    procedure CalcFormula(ACell: PCell);
    procedure CalcFormulas;
    function ConvertRPNFormulaToStringFormula(const AFormula: TsRPNFormula): String;
    function GetCalcState(ACell: PCell): TsCalcState;
    procedure SetCalcState(ACell: PCell; AValue: TsCalcState);

    { Data manipulation methods - For Cells }
    procedure CopyCell(AFromCell, AToCell: PCell); overload;
    procedure CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal;
      AFromWorksheet: TsWorksheet = nil); overload;
    procedure CopyFormat(AFromCell, AToCell: PCell); overload;
    procedure CopyFormat(AFormatCell: PCell; AToRow, AToCol: Cardinal); overload;
    procedure CopyFormula(AFromCell, AToCell: PCell); overload;
    procedure CopyFormula(AFormulaCell: PCell; AToRow, AToCol: Cardinal); overload;
    procedure CopyValue(AFromCell, AToCell: PCell); overload;
    procedure CopyValue(AValueCell: PCell; AToRow, AToCol: Cardinal); overload;

    procedure DeleteCell(ACell: PCell);
    procedure EraseCell(ACell: PCell);

    function  AddCell(ARow, ACol: Cardinal): PCell;
    function  FindCell(ARow, ACol: Cardinal): PCell; overload;
    function  FindCell(AddressStr: String): PCell; overload;
    function  GetCell(ARow, ACol: Cardinal): PCell; overload;
    function  GetCell(AddressStr: String): PCell; overload;
    function  GetCellCount: Cardinal;

    function  GetFirstColIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastColIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastColNumber: Cardinal; deprecated 'Use GetLastColIndex';
    function  GetLastOccupiedColIndex: Cardinal;
    function  GetFirstRowIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastOccupiedRowIndex: Cardinal;
    function  GetLastRowIndex(AForceCalculation: Boolean = false): Cardinal;
    function  GetLastRowNumber: Cardinal; deprecated 'Use GetLastRowIndex';

    { Data manipulation methods - For Rows and Cols }
    function  CalcAutoRowHeight(ARow: Cardinal): Single;
    function  FindRow(ARow: Cardinal): PRow;
    function  FindCol(ACol: Cardinal): PCol;
    function  GetCellCountInRow(ARow: Cardinal): Cardinal;
    function  GetCellCountInCol(ACol: Cardinal): Cardinal;
    function  GetRow(ARow: Cardinal): PRow;
    function  GetRowHeight(ARow: Cardinal): Single;
    function  GetCol(ACol: Cardinal): PCol;
    function  GetColWidth(ACol: Cardinal): Single;
    procedure DeleteCol(ACol: Cardinal);
    procedure DeleteRow(ARow: Cardinal);
    procedure InsertCol(ACol: Cardinal);
    procedure InsertRow(ARow: Cardinal);
    procedure RemoveAllRows;
    procedure RemoveAllCols;
    procedure RemoveCol(ACol: Cardinal);
    procedure RemoveRow(ARow: Cardinal);
    procedure WriteRowInfo(ARow: Cardinal; AData: TRow);
    procedure WriteRowHeight(ARow: Cardinal; AHeight: Single);
    procedure WriteColInfo(ACol: Cardinal; AData: TCol);
    procedure WriteColWidth(ACol: Cardinal; AWidth: Single);

    // Sorting
    procedure Sort(const ASortParams: TsSortParams;
      ARowFrom, AColFrom, ARowTo, AColTo: Cardinal); overload;
    procedure Sort(ASortParams: TsSortParams; ARange: String); overload;

    // Selected cell and ranges
    procedure SelectCell(ARow, ACol: Cardinal);
    procedure ClearSelection;
    function GetSelection: TsCellRangeArray;
    function GetSelectionAsString: String;
    function GetSelectionCount: Integer;
    procedure SetSelection(const ASelection: TsCellRangeArray);

    // Comments
    function FindComment(ACell: PCell): PsComment;
    function HasComment(ACell: PCell): Boolean;
    function ReadComment(ARow, ACol: Cardinal): String; overload;
    function ReadComment(ACell: PCell): string; overload;
    procedure RemoveComment(ACell: PCell);
    function WriteComment(ARow, ACol: Cardinal; AText: String): PCell; overload;
    procedure WriteComment(ACell: PCell; AText: String); overload;

    // Hyperlinks
    function FindHyperlink(ACell: PCell): PsHyperlink;
    function HasHyperlink(ACell: PCell): Boolean;
    function ReadHyperlink(ACell: PCell): TsHyperlink;
    procedure RemoveHyperlink(ACell: PCell);
    function ValidHyperlink(AValue: String; out AErrMsg: String): Boolean;
    function WriteHyperlink(ARow, ACol: Cardinal; ATarget: String;
      ATooltip: String = ''): PCell; overload;
    procedure WriteHyperlink(ACell: PCell; ATarget: String;
      ATooltip: String = ''); overload;

    { Merged cells }
    function FindMergeBase(ACell: PCell): PCell;
    function FindMergedRange(ACell: PCell; out ARow1, ACol1, ARow2, ACol2: Cardinal): Boolean;
    procedure MergeCells(ARow1, ACol1, ARow2, ACol2: Cardinal); overload;
    procedure MergeCells(ARange: String); overload;
    function InSameMergedRange(ACell1, ACell2: PCell): Boolean;
    function IsMergeBase(ACell: PCell): Boolean;
    function IsMerged(ACell: PCell): Boolean;
    procedure UnmergeCells(ARow, ACol: Cardinal); overload;
    procedure UnmergeCells(ARange: String); overload;

    // Notification of changed cells
    procedure ChangedCell(ARow, ACol: Cardinal);
    procedure ChangedFont(ARow, ACol: Cardinal);

    { Properties }

    {@@ List of cells of the worksheet. Only cells with contents or with formatting
        are listed }
    property  Cells: TsCells read FCells;
    {@@ List of all column records of the worksheet having a non-standard column width }
    property  Cols: TIndexedAVLTree read FCols;
    {@@ List of all comment records }
    property  Comments: TsComments read FComments;
    {@@ List of merged cells (contains TsCellRange records) }
    property  MergedCells: TsMergedCells read FMergedCells;
    {@@ List of hyperlink information records }
    property  Hyperlinks: TsHyperlinks read FHyperlinks;
    {@@ FormatSettings for localization of some formatting strings }
    property  FormatSettings: TFormatSettings read GetFormatSettings;
    {@@ Name of the sheet. In the popular spreadsheet applications this is
      displayed at the tab of the sheet. }
    property Name: string read FName write SetName;
    {@@ List of all row records of the worksheet having a non-standard row height }
    property  Rows: TIndexedAVLTree read FRows;
    {@@ Workbook to which the worksheet belongs }
    property  Workbook: TsWorkbook read FWorkbook;
    {@@ The default column width given in "character units" (width of the
      character "0" in the default font) }
    property DefaultColWidth: Single read FDefaultColWidth write FDefaultColWidth;
    {@@ The default row height is given in "line count" (height of the
      default font }
    property DefaultRowHeight: Single read FDefaultRowHeight write FDefaultRowHeight;

    // These are properties to interface to TsWorksheetGrid
    {@@ Parameters controlling visibility of grid lines and row/column headers,
        usage of frozen panes etc. }
    property  Options: TsSheetOptions read FOptions write FOptions;
    {@@ Column index of the selected cell of this worksheet }
    property  ActiveCellCol: Cardinal read FActiveCellCol;
    {@@ Row index of the selected cell of this worksheet }
    property  ActiveCellRow: Cardinal read FActiveCellRow;
    {@@ Number of frozen columns which do not scroll }
    property  LeftPaneWidth: Integer read FLeftPaneWidth write FLeftPaneWidth;
    {@@ Number of frozen rows which do not scroll }
    property  TopPaneHeight: Integer read FTopPaneHeight write FTopPaneHeight;
    {@@ Event fired when cell contents or formatting changes }
    property  OnChangeCell: TsCellEvent read FOnChangeCell write FOnChangeCell;
    {@@ Event fired when the font size in a cell changes }
    property  OnChangeFont: TsCellEvent read FOnChangeFont write FOnChangeFont;
    {@@ Event to override cell comparison for sorting }
    property  OnCompareCells: TsCellCompareEvent read FOnCompareCells write FOnCompareCells;
    {@@ Event fired when a cell is "selected". }
    property  OnSelectCell: TsCellEvent read FOnSelectCell write FOnSelectCell;

  end;

  {@@
    Option flags for the workbook

    @param  boVirtualMode      If in virtual mode date are not taken from cells
                               when a spreadsheet is written to file, but are
                               provided by means of the event OnWriteCellData.
                               Similarly, when data are read they are not added as
                               cells but passed the the event OnReadCellData;
    @param  boBufStream        When this option is set a buffered stream is used
                               for writing (a memory stream swapping to disk) or
                               reading (a file stream pre-reading chunks of data
                               to memory)
    @param  boAutoCalc         Automatically recalculate rpn formulas whenever
                               a cell value changes.
    @param  boCalcBeforeSaving Calculates formulas before saving the file.
                               Otherwise there are no results when the file is
                               loaded back by fpspreadsheet.
    @param  boReadFormulas     Allows to turn off reading of rpn formulas; this is
                               a precaution since formulas not correctly
                               implemented by fpspreadsheet could crash the
                               reading operation. }
  TsWorkbookOption = (boVirtualMode, boBufStream, boAutoCalc, boCalcBeforeSaving,
    boReadFormulas);

  {@@ Set of option flags for the workbook }
  TsWorkbookOptions = set of TsWorkbookOption;

  {@@
    Event fired when writing a file in virtual mode. The event handler has to
    pass data ("AValue") and formatting style to be copied from a template
    cell ("AStyleCell") to the writer }
  TsWorkbookWriteCellDataEvent = procedure(Sender: TObject; ARow, ACol: Cardinal;
    var AValue: variant; var AStyleCell: PCell) of object;

  {@@
    Event fired when reading a file in virtual mode. Read data are provided in
    the "ADataCell" (which is not added to the worksheet in virtual mode). }
  TsWorkbookReadCellDataEvent = procedure(Sender: TObject; ARow, ACol: Cardinal;
    const ADataCell: PCell) of object;

  {@@ Event procedure containing a specific worksheet }
  TsWorksheetEvent = procedure (Sender: TObject; ASheet: TsWorksheet) of object;

  {@@ Event procedure called when a worksheet is removed }
  TsRemoveWorksheetEvent = procedure (Sender: TObject; ASheetIndex: Integer) of object;

  {@@ The workbook contains the worksheets and provides methods for reading from
    and writing to file.
  }
  TsWorkbook = class
  private
    { Internal data }
    FWorksheets: TFPList;
    FFormat: TsSpreadsheetFormat;
    FBuiltinFontCount: Integer;
    //FPalette: array of TsColorValue;
    FVirtualColCount: Cardinal;
    FVirtualRowCount: Cardinal;
    FReadWriteFlag: TsReadWriteFlag;
    FCalculationLock: Integer;
    FOptions: TsWorkbookOptions;
    FActiveWorksheet: TsWorksheet;
    FOnOpenWorkbook: TNotifyEvent;
    FOnWriteCellData: TsWorkbookWriteCellDataEvent;
    FOnReadCellData: TsWorkbookReadCellDataEvent;
    FOnChangeWorksheet: TsWorksheetEvent;
    FOnRenameWorksheet: TsWorksheetEvent;
    FOnAddWorksheet: TsWorksheetEvent;
    FOnRemoveWorksheet: TsRemoveWorksheetEvent;
    FOnRemovingWorksheet: TsWorksheetEvent;
    FOnSelectWorksheet: TsWorksheetEvent;
//    FOnChangePalette: TNotifyEvent;
    FFileName: String;
    FLockCount: Integer;
    FLog: TStringList;

    { Setter/Getter }
    function GetErrorMsg: String;
    procedure SetVirtualColCount(AValue: Cardinal);
    procedure SetVirtualRowCount(AValue: Cardinal);

    { Callback procedures }
    procedure RemoveWorksheetsCallback(data, arg: pointer);

  protected
    FFontList: TFPList;
    FNumFormatList: TFPList;
    FCellFormatList: TsCellFormatList;

    { Internal methods }
    procedure GetLastRowColIndex(out ALastRow, ALastCol: Cardinal);
    procedure PrepareBeforeReading;
    procedure PrepareBeforeSaving;
    procedure ReCalc;

  public
    {@@ A copy of SysUtil's DefaultFormatSettings (converted to UTF8) to provide
      some kind of localization to some formatting strings.
      Can be modified before loading/writing files }
    FormatSettings: TFormatSettings;

    { Base methods }
    constructor Create;
    destructor Destroy; override;

    class function GetFormatFromFileHeader(const AFileName: TFileName;
      out SheetType: TsSpreadsheetFormat): Boolean;
    function CreateSpreadReader(AFormat: TsSpreadsheetFormat): TsBasicSpreadReader;
    function CreateSpreadWriter(AFormat: TsSpreadsheetFormat): TsBasicSpreadWriter;
    procedure ReadFromFile(AFileName: string; AFormat: TsSpreadsheetFormat); overload;
    procedure ReadFromFile(AFileName: string); overload;
    procedure ReadFromFileIgnoringExtension(AFileName: string);
    procedure ReadFromStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
    procedure WriteToFile(const AFileName: string;
      const AFormat: TsSpreadsheetFormat;
      const AOverwriteExisting: Boolean = False); overload;
    procedure WriteToFile(const AFileName: String; const AOverwriteExisting: Boolean = False); overload;
    procedure WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);

    { Worksheet list handling methods }
    function  AddWorksheet(AName: string;
      ReplaceDuplicateName: Boolean = false): TsWorksheet;
    function  CopyWorksheetFrom(AWorksheet: TsWorksheet;
      ReplaceDuplicateName: Boolean = false): TsWorksheet;
    function  GetFirstWorksheet: TsWorksheet;
    function  GetWorksheetByIndex(AIndex: Integer): TsWorksheet;
    function  GetWorksheetByName(AName: String): TsWorksheet;
    function  GetWorksheetCount: Integer;
    function  GetWorksheetIndex(AWorksheet: TsWorksheet): Integer;
    procedure RemoveAllWorksheets;
    procedure RemoveWorksheet(AWorksheet: TsWorksheet);
    procedure SelectWorksheet(AWorksheet: TsWorksheet);
    function  ValidWorksheetName(var AName: String;
      ReplaceDuplicateName: Boolean = false): Boolean;

    { String-to-cell/range conversion }
    function TryStrToCell(AText: String; out AWorksheet: TsWorksheet;
      out ARow,ACol: Cardinal; AListSeparator: Char = #0): Boolean;
    function TryStrToCellRange(AText: String; out AWorksheet: TsWorksheet;
      out ARange: TsCellRange; AListSeparator: Char = #0): Boolean;
    function TryStrToCellRanges(AText: String; out AWorksheet: TsWorksheet;
      out ARanges: TsCellRangeArray; AListSeparator: Char = #0): Boolean;

    { Cell format handling }
    function AddCellFormat(const AValue: TsCellFormat): Integer;
    function GetCellFormat(AIndex: Integer): TsCellFormat;
    function GetCellFormatAsString(AIndex: Integer): String;
    function GetNumCellFormats: Integer;
    function GetPointerToCellFormat(AIndex: Integer): PsCellFormat;

    { Font handling }
    function AddFont(const AFontName: String; ASize: Single;
      AStyle: TsFontStyles; AColor: TsColor): Integer; overload;
    function AddFont(const AFont: TsFont): Integer; overload;
    procedure DeleteFont(AFontIndex: Integer);
    function FindFont(const AFontName: String; ASize: Single;
      AStyle: TsFontStyles; AColor: TsColor): Integer;
    function GetBuiltinFontCount: Integer;
    function GetDefaultFont: TsFont;
    function GetDefaultFontSize: Single;
    function GetFont(AIndex: Integer): TsFont;
    function GetFontAsString(AIndex: Integer): String;
    function GetFontCount: Integer;
    function GetHyperlinkFont: TsFont;
    procedure InitFonts;
    procedure RemoveAllFonts;
    procedure ReplaceFont(AFontIndex: Integer; AFontName: String;
      ASize: Single; AStyle: TsFontStyles; AColor: TsColor);
    procedure SetDefaultFont(const AFontName: String; ASize: Single);

    { Number format handling }
    function AddNumberFormat(AFormatStr: String): Integer;
    function GetNumberFormat(AIndex: Integer): TsNumFormatParams;
    function GetNumberFormatCount: Integer;
                      (*
    { Color handling }
    function FPSColorToHexString(AColor: TsColor; ARGBColor: TFPColor): String;
    function GetColorName(AColorIndex: TsColor): string; overload;
    procedure GetColorName(AColorValue: TsColorValue; out AName: String); overload;
    function GetPaletteColor(AColorIndex: TsColor): TsColorValue;
    function GetPaletteColorAsHTMLStr(AColorIndex: TsColor): String;
    procedure SetPaletteColor(AColorIndex: TsColor; AColorValue: TsColorValue);
    function GetPaletteSize: Integer;
    procedure UseDefaultPalette;
    procedure UsePalette(APalette: PsPalette; APaletteCount: Word;
      ABigEndian: Boolean = false);
    function UsesColor(AColorIndex: TsColor): Boolean;
                        *)

    { Utilities }
    procedure UpdateCaches;

    { Error messages }
    procedure AddErrorMsg(const AMsg: String); overload;
    procedure AddErrorMsg(const AMsg: String; const Args: array of const); overload;
    procedure ClearErrorList;

    {@@ Identifies the "active" worksheet (only for visual controls)}
    property ActiveWorksheet: TsWorksheet read FActiveWorksheet;
    {@@ Retrieves error messages collected during reading/writing }
    property ErrorMsg: String read GetErrorMsg;
    {@@ Filename of the saved workbook }
    property FileName: String read FFileName;
    {@@ Identifies the file format which was detected when reading the file }
    property FileFormat: TsSpreadsheetFormat read FFormat;
    property VirtualColCount: cardinal read FVirtualColCount write SetVirtualColCount;
    property VirtualRowCount: cardinal read FVirtualRowCount write SetVirtualRowCount;
    property Options: TsWorkbookOptions read FOptions write FOptions;

    {@@ This event fires whenever a new worksheet is added }
    property OnAddWorksheet: TsWorksheetEvent read FOnAddWorksheet write FOnAddWorksheet;
    {@@ This event fires whenever the workbook palette changes. }
//    property OnChangePalette: TNotifyEvent read FOnChangePalette write FOnChangePalette;
    {@@ This event fires whenever a worksheet is changed }
    property OnChangeWorksheet: TsWorksheetEvent read FOnChangeWorksheet write FOnChangeWorksheet;
    {@@ This event fires whenever a workbook is loaded }
    property OnOpenWorkbook: TNotifyEvent read FOnOpenWorkbook write FOnOpenWorkbook;
    {@@ This event fires whenever a worksheet is renamed }
    property OnRenameWorksheet: TsWorksheetEvent read FOnRenameWorksheet write FOnRenameWorksheet;
    {@@ This event fires AFTER a worksheet has been deleted }
    property OnRemoveWorksheet: TsRemoveWorksheetEvent read FOnRemoveWorksheet write FOnRemoveWorksheet;
    {@@ This event fires BEFORE a worksheet is deleted }
    property OnRemovingWorksheet: TsWorksheetEvent read FOnRemovingWorksheet write FOnRemovingWorksheet;
    {@@ This event fires when a worksheet is made "active"}
    property OnSelectWorksheet: TsWorksheetEvent read FOnSelectWorksheet write FOnSelectWorksheet;
    {@@ This event allows to provide external cell data for writing to file,
      standard cells are ignored. Intended for converting large database files
      to a spreadsheet format. Requires Option boVirtualMode to be set. }
    property OnWriteCellData: TsWorkbookWriteCellDataEvent read FOnWriteCellData write FOnWriteCellData;
    {@@ This event accepts cell data while reading a spreadsheet file. Data are
      not encorporated in a spreadsheet, they are just passed through to the
      event handler for processing. Requires option boVirtualMode to be set. }
    property OnReadCellData: TsWorkbookReadCellDataEvent read FOnReadCellData write FOnReadCellData;
  end;

  { TsBasicSpreadReaderWriter }
  TsBasicSpreadReaderWriter = class
  protected
    {@@ Instance of the workbook which is currently being read or written. }
    FWorkbook: TsWorkbook;
    {@@ Instance of the worksheet which is currently being read or written. }
    FWorksheet: TsWorksheet;
    {@@ Limitations for the specific data file format }
    FLimitations: TsSpreadsheetFormatLimitations;
  public
    constructor Create(AWorkbook: TsWorkbook); virtual;  // to allow descendents to override it
    function Limitations: TsSpreadsheetFormatLimitations;
    {@@ Instance of the workbook which is currently being read/written. }
    property Workbook: TsWorkbook read FWorkbook;
  end;

  { TsBasicSpreadReader }
  TsBasicSpreadReader = class(TsBasicSpreadReaderWriter)
  public
    { General writing methods }
    procedure ReadFromFile(AFileName: string); virtual; abstract;
    procedure ReadFromStream(AStream: TStream); virtual; abstract;
    procedure ReadFromStrings(AStrings: TStrings); virtual; abstract;
  end;

  { TsBasicSpreadWriter }
  TsBasicSpreadWriter = class(TsBasicSpreadReaderWriter)
  public
    { Helpers }
    procedure CheckLimitations; virtual;
    { General writing methods }
    procedure WriteToFile(const AFileName: string;
      const AOverwriteExisting: Boolean = False); virtual; abstract;
    procedure WriteToStream(AStream: TStream); virtual; abstract;
    procedure WriteToStrings(AStrings: TStrings); virtual; abstract;
  end;

  {@@ TsSpreadReader class reference type }
  TsSpreadReaderClass = class of TsBasicSpreadReader;

  {@@ TsSpreadWriter class reference type }
  TsSpreadWriterClass = class of TsBasicSpreadWriter;


procedure CopyCellFormat(AFromCell, AToCell: PCell);
procedure CopyCellValue(AFromCell, AToCell: PCell);

//function SameCellBorders(ACell1, ACell2: PCell): Boolean; overload;
function SameCellBorders(AFormat1, AFormat2: PsCellFormat): Boolean; //overload;

function HasFormula(ACell: PCell): Boolean;

{ For debugging purposes }
procedure DumpFontsToFile(AWorkbook: TsWorkbook; AFileName: String);


implementation

uses
  Math, StrUtils, DateUtils, TypInfo, lazutf8, lazFileUtils, URIParser,
  fpsStrings, uvirtuallayer_ole,
  fpsUtils, fpsreaderwriter, fpsCurrency, fpsExprParser,
  fpsNumFormat, fpsNumFormatParser;

(*
const
  { These are reserved system colors by Microsoft
    0x0040 - Default foreground color - window text color in the sheet display.
    0x0041 - Default background color - window background color in the sheet
             display and is the default background color for a cell.
    0x004D - Default chart foreground color - window text color in the
             chart display.
    0x004E - Default chart background color - window background color in the
             chart display.
    0x004F - Chart neutral color which is black, an RGB value of (0,0,0).
    0x0051 - ToolTip text color - automatic font color for comments.
    0x7FFF - Font automatic color - window text color. }

  // Color indexes of reserved system colors
  DEF_FOREGROUND_COLOR = $0040;
  DEF_BACKGROUND_COLOR = $0041;
  DEF_CHART_FOREGROUND_COLOR = $004D;
  DEF_CHART_BACKGROUND_COLOR = $004E;
  DEF_CHART_NEUTRAL_COLOR = $004F;
  DEF_TOOLTIP_TEXT_COLOR = $0051;
  DEF_FONT_AUTOMATIC_COLOR = $7FFF;

  // Color rgb values of reserved system colors
  DEF_FOREGROUND_COLORVALUE = $000000;
  DEF_BACKGROUND_COLORVALUE = $FFFFFF;
  DEF_CHART_FOREGROUND_COLORVALUE = $000000;
  DEF_CHART_BACKGROUND_COLORVALUE = $FFFFFF;
  DEF_CHART_NEUTRAL_COLORVALUE = $FFFFFF;
  DEF_TOOLTIP_TEXT_COLORVALUE = $000000;
  DEF_FONT_AUTOMATIC_COLORVALUE = $000000;
       *)

{@@ ----------------------------------------------------------------------------
  Copies the format of a cell to another one.

  @param  AFromCell   Cell from which the format is to be copied
  @param  AToCell     Cell to which the format is to be copied
-------------------------------------------------------------------------------}
procedure CopyCellFormat(AFromCell, AToCell: PCell);
var
  sourceSheet, destSheet: TsWorksheet;
  fmt: TsCellFormat;
  numFmtParams: TsNumFormatParams;
  nfs: String;
  font: TsFont;
begin
  Assert(AFromCell <> nil);
  Assert(AToCell <> nil);
  sourceSheet := TsWorksheet(AFromCell.Worksheet);
  destSheet := TsWorksheet(AToCell.Worksheet);
  if (sourceSheet=nil) or (destSheet=nil) or (sourceSheet.Workbook = destSheet.Workbook) then
    AToCell^.FormatIndex := AFromCell^.FormatIndex
  else
  begin
    fmt := sourceSheet.ReadCellFormat(AFromCell);
    //destSheet.WriteCellFormat(AToCell, fmt);
    {
    if (uffBackground in fmt.UsedFormattingFields) then
    begin
      clr := sourceSheet.Workbook.GetPaletteColor(fmt.Background.BgColor);
      fmt.Background.BgColor := destSheet.Workbook.AddColorToPalette(clr);
      clr := sourceSheet.Workbook.GetPaletteColor(fmt.Background.FgColor);
      fmt.Background.FgColor := destSheet.Workbook.AddColorToPalette(clr);
    end;
    }
    if (uffFont in fmt.UsedFormattingFields) then
    begin
      font := sourceSheet.ReadCellFont(AFromCell);
      {
      clr := sourceSheet.Workbook.GetPaletteColor(font.Color);
      font.Color := destSheet.Workbook.AddColorToPalette(clr);
      }
      fmt.FontIndex := destSheet.Workbook.FindFont(font.FontName, font.Size, font.Style, font.Color);
      if fmt.FontIndex = -1 then
        fmt.FontIndex := destSheet.Workbook.AddFont(font.FontName, font.Size, font.Style, font.Color);
    end;
    {
    if (uffBorder in fmt.UsedFormattingFields) then
      for cb in fmt.Border do
      begin
        clr := sourceSheet.Workbook.GetPaletteColor(fmt.BorderStyles[cb].Color);
        fmt.BorderStyles[cb].Color := destSheet.Workbook.AddColorToPalette(clr);
      end;
      }
    if (uffNumberformat in fmt.UsedFormattingFields) then
    begin
      numFmtParams := sourceSheet.Workbook.GetNumberFormat(fmt.NumberFormatIndex);
      if numFmtParams <> nil then
      begin
        nfs := numFmtParams.NumFormatStr;
        fmt.NumberFormatIndex := destSheet.Workbook.AddNumberFormat(nfs);
      end;
    end;

    destSheet.WriteCellFormat(AToCell, fmt);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Copies the value of a cell to another one. Does not copy the formula, erases
  the formula of the destination cell if there is one!

  @param  AFromCell   Cell from which the value is to be copied
  @param  AToCell     Cell to which the value is to be copied
-------------------------------------------------------------------------------}
procedure CopyCellValue(AFromCell, AToCell: PCell);
begin
  Assert(AFromCell <> nil);
  Assert(AToCell <> nil);

  AToCell^.ContentType := AFromCell^.ContentType;
  AToCell^.NumberValue := AFromCell^.NumberValue;
  AToCell^.DateTimeValue := AFromCell^.DateTimeValue;
  AToCell^.BoolValue := AFromCell^.BoolValue;
  AToCell^.ErrorValue := AFromCell^.ErrorValue;
  AToCell^.UTF8StringValue := AFromCell^.UTF8StringValue;
  AToCell^.FormulaValue := '';    // This is confirmed with Excel
end;

{@@ ----------------------------------------------------------------------------
  Checks whether two format records have same border attributes

  @param  AFormat1  Pointer to the first one of the two format records to be compared
  @param  AFormat2  Pointer to the second one of the two format records to be compared
-------------------------------------------------------------------------------}
function SameCellBorders(AFormat1, AFormat2: PsCellFormat): Boolean;

  function NoBorder(AFormat: PsCellFormat): Boolean;
  begin
    Result := (AFormat = nil) or
      not (uffBorder in AFormat^.UsedFormattingFields) or
      (AFormat^.Border = []);
  end;

var
  nobrdr1, nobrdr2: Boolean;
  cb: TsCellBorder;
begin
  nobrdr1 := NoBorder(AFormat1);
  nobrdr2 := NoBorder(AFormat2);
  if (nobrdr1 and nobrdr2) then
    Result := true
  else
  if (nobrdr1 and (not nobrdr2) ) or ( (not nobrdr1) and nobrdr2) then
    Result := false
  else begin
    Result := false;
    if AFormat1^.Border <> AFormat2^.Border then
      exit;
    for cb in TsCellBorder do begin
      if AFormat1^.BorderStyles[cb].LineStyle <> AFormat2^.BorderStyles[cb].LineStyle then
        exit;
      if AFormat1^.BorderStyles[cb].Color <> AFormat2^.BorderStyles[cb].Color then
        exit;
    end;
    Result := true;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns TRUE if the cell contains a formula.

  @param   ACell   Pointer to the cell checked
-------------------------------------------------------------------------------}
function HasFormula(ACell: PCell): Boolean;
begin
  Result := Assigned(ACell) and (Length(ACell^.FormulaValue) > 0);
end;

function CompareCells(Item1, Item2: Pointer): Integer;
begin
  result := LongInt(PCell(Item1).Row) - PCell(Item2).Row;
  if Result = 0 then
    Result := LongInt(PCell(Item1).Col) - PCell(Item2).Col;
end;

function CompareRows(Item1, Item2: Pointer): Integer;
begin
  Result := LongInt(PRow(Item1).Row) - PRow(Item2).Row;
end;

function CompareCols(Item1, Item2: Pointer): Integer;
begin
  Result := LongInt(PCol(Item1).Col) - PCol(Item2).Col;
end;

function CompareMergedCells(Item1, Item2: Pointer): Integer;
begin
  Result := LongInt(PsCellRange(Item1)^.Row1) - PsCellRange(Item2)^.Row1;
  if Result = 0 then
    Result := LongInt(PsCellRange(Item1)^.Col1) - PsCellRange(Item2)^.Col1;
end;


{@@ ----------------------------------------------------------------------------
  Write the fonts stored for a given workbook to a file.
  FOR DEBUGGING ONLY.
-------------------------------------------------------------------------------}
procedure DumpFontsToFile(AWorkbook: TsWorkbook; AFileName: String);
var
  L: TStringList;
  i: Integer;
  fnt: TsFont;
begin
  L := TStringList.Create;
  try
    for i:=0 to AWorkbook.GetFontCount-1 do begin
      fnt := AWorkbook.GetFont(i);
      if fnt = nil then
        L.Add(Format('#%.3d: ---------------', [i]))
      else
        L.Add(Format('#%.3d: %-15s %4.1f %s%s%s%s %s', [
          i,
          fnt.FontName,
          fnt.Size,
          IfThen(fssBold in fnt.Style, 'b', '.'),
          IfThen(fssItalic in fnt.Style, 'i', '.'),
          IfThen(fssUnderline in fnt.Style, 'u', '.'),
          IfThen(fssStrikeOut in fnt.Style, 's', '.'),
          ColorToHTMLColorStr(fnt.Color)
          //AWorkbook.GetPaletteColorAsHTMLStr(fnt.Color)
        ]));
    end;
    L.SaveToFile(AFileName);
  finally
    L.Free;
  end;
end;


{*******************************************************************************
*                           TsWorksheet                                        *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Constructor of the TsWorksheet class.
-------------------------------------------------------------------------------}
constructor TsWorksheet.Create;
begin
  inherited Create;

  FCells := TsCells.Create(self);
  FRows := TIndexedAVLTree.Create(@CompareRows);
  FCols := TIndexedAVLTree.Create(@CompareCols);
  FComments := TsComments.Create;
  FMergedCells := TsMergedCells.Create;
  FHyperlinks := TsHyperlinks.Create;

  InitPageLayout(PageLayout);

  FDefaultColWidth := 12;
  FDefaultRowHeight := 1;

  FFirstRowIndex := $FFFFFFFF;
  FFirstColIndex := $FFFFFFFF;
  FLastRowIndex := 0;
  FLastColIndex := 0;

  FActiveCellRow := Cardinal(-1);
  FActiveCellCol := Cardinal(-1);

  FOptions := [soShowGridLines, soShowHeaders];
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsWorksheet class.
  Releases all memory, but does not delete from the workbook's worksheetList !!!
  NOTE: Don't call directly. Always use Workbook.RemoveWorksheet to remove a
  worksheet from a workbook.
-------------------------------------------------------------------------------}
destructor TsWorksheet.Destroy;
begin
  RemoveAllRows;
  RemoveAllCols;

  FCells.Free;
  FRows.Free;
  FCols.Free;
  FComments.Free;
  FMergedCells.Free;
  FHyperlinks.Free;

  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Helper function which constructs an rpn formula from the cell's string
  formula. This is needed, for example, when writing a formula to xls biff
  file format.
  The formula is stored in ACell.
  If ADestCell is not nil then the relative references are adjusted as seen
  from ADestCell. This means that this function returns the formula that
  would be created if ACell is copied to the location of ADestCell.
  Needed for copying formulas and for splitting shared formulas.
-------------------------------------------------------------------------------}
function TsWorksheet.BuildRPNFormula(ACell: PCell;
  ADestCell: PCell = nil): TsRPNFormula;
var
  parser: TsSpreadsheetParser;
begin
  if not HasFormula(ACell) then begin
    SetLength(Result, 0);
    exit;
  end;
  parser := TsSpreadsheetParser.Create(self);
  try
    if ADestCell <> nil then
      parser.PrepareCopyMode(ACell, ADestCell);
    parser.Expression := ACell^.FormulaValue;
    Result := parser.RPNFormula;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Calculates the formula in a cell
  Should not be called by itself because the result may depend on other cells
  which may have not yet been calculated. It is better to call CalcFormulas
  instead.

  @param  ACell  Cell containing the formula.
-------------------------------------------------------------------------------}
procedure TsWorksheet.CalcFormula(ACell: PCell);
var
  parser: TsSpreadsheetParser;
  res: TsExpressionResult;
  p: Integer;
  link, txt: String;
  cell: PCell;
begin
  ACell^.Flags := ACell^.Flags + [cfCalculating] - [cfCalculated];

  parser := TsSpreadsheetParser.Create(self);
  try
    try
      parser.Expression := ACell^.FormulaValue;
      res := parser.Evaluate;
    except
      on E:ECalcEngine do
      begin
        Workbook.AddErrorMsg(E.Message);
        Res := ErrorResult(errIllegalRef);
      end;
    end;

    case res.ResultType of
      rtEmpty     : WriteBlank(ACell);
      rtError     : WriteErrorValue(ACell, res.ResError);
      rtInteger   : WriteNumber(ACell, res.ResInteger);
      rtFloat     : WriteNumber(ACell, res.ResFloat);
      rtDateTime  : WriteDateTime(ACell, res.ResDateTime);
      rtString    : WriteUTF8Text(ACell, res.ResString);
      rtHyperlink : begin
                      link := ArgToString(res);
                      p := pos(HYPERLINK_SEPARATOR, link);
                      if p > 0 then
                      begin
                        txt := Copy(link, p+Length(HYPERLINK_SEPARATOR), Length(link));
                        link := Copy(link, 1, p-1);
                      end else
                        txt := link;
                      WriteHyperlink(ACell, link);
                      WriteUTF8Text(ACell, txt);
                    end;
      rtBoolean   : WriteBoolValue(ACell, res.ResBoolean);
      rtCell      : begin
                      cell := GetCell(res.ResRow, res.ResCol);
                      case cell^.ContentType of
                        cctNumber    : WriteNumber(ACell, cell^.NumberValue);
                        cctDateTime  : WriteDateTime(ACell, cell^.DateTimeValue);
                        cctUTF8String: WriteUTF8Text(ACell, cell^.UTF8StringValue);
                        cctBool      : WriteBoolValue(ACell, cell^.Boolvalue);
                        cctError     : WriteErrorValue(ACell, cell^.ErrorValue);
                        cctEmpty     : WriteBlank(ACell);
                      end;
                    end;
    end;
  finally
    parser.Free;
  end;

  ACell^.Flags := ACell^.Flags + [cfCalculated] - [cfCalculating];
end;

{@@ ----------------------------------------------------------------------------
  Calculates all formulas of the worksheet.

  Since formulas may reference not-yet-calculated cells, this occurs in
  two steps:
  1. All formula cells are marked as "not calculated".
  2. Cells are calculated. If referenced cells are found as being
     "not calculated" they are calculated and then tagged as "calculated".
  This results in an iterative calculation procedure. In the end, all cells
  are calculated.
-------------------------------------------------------------------------------}
procedure TsWorksheet.CalcFormulas;
var
  cell: PCell;
begin
  // prevent infinite loop due to triggering of formula calculation whenever
  // a cell changes during execution of CalcFormulas.
  inc(FWorkbook.FCalculationLock);
  try
    // Step 1 - mark all formula cells as "not calculated"
    for cell in FCells do
      if HasFormula(cell) then
        SetCalcState(cell, csNotCalculated);

    // Step 2 - calculate cells. If a not-yet-calculated cell is found it is
    // calculated and then marked as such.
    for cell in FCells do
      if HasFormula(cell) and (cell^.ContentType <> cctError) then
        CalcFormula(cell);

  finally
    dec(FWorkbook.FCalculationLock);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a cell given by its row and column indexes belongs to a
  specified rectangular cell range.
-------------------------------------------------------------------------------}
class function TsWorksheet.CellInRange(ARow, ACol: Cardinal;
  ARange: TsCellRange): Boolean;
begin
  Result := (ARow >= ARange.Row1) and (ARow <= ARange.Row2) and
            (ACol >= ARange.Col1) and (ACol <= ARange.Col2);
end;

{@@ ----------------------------------------------------------------------------
  Converts a FPSpreadsheet cell position, which is Row, Col in numbers
  and zero based - e.g. 0,0 - to a textual representation which is [Col][Row],
  where the Col is in letters and the row is in 1-based numbers - e.g. A1
-------------------------------------------------------------------------------}
class function TsWorksheet.CellPosToText(ARow, ACol: Cardinal): string;
begin
  Result := GetCellString(ARow, ACol, [rfRelCol, rfRelRow]);
end;

{@@ ----------------------------------------------------------------------------
  Checks entire workbook, whether this cell is used in any formula.

  @param   ARow  Row index of the cell considered
  @param   ACol  Column index of the cell considered
  @return  TRUE if the cell is used in a formula, FALSE if not
-------------------------------------------------------------------------------}
function TsWorksheet.CellUsedInFormula(ARow, ACol: Cardinal): Boolean;
var
  cell: PCell;
  fe: TsFormulaElement;
  i: Integer;
  rpnFormula: TsRPNFormula;
begin
  for cell in FCells do
  begin
    if HasFormula(cell) then begin
      if (cell^.Row = ARow) and (cell^.Col = ACol) then
      begin
        Result := true;
        exit;
      end;
      rpnFormula := BuildRPNFormula(cell);
      for i := 0 to Length(rpnFormula)-1 do
      begin
        fe := rpnFormula[i];
        case fe.ElementKind of
          fekCell, fekCellRef:
            if (fe.Row = ARow) and (fe.Col = ACol) then
            begin
              Result := true;
              exit;
            end;
          fekCellRange:
            if (fe.Row <= ARow) and (ARow <= fe.Row2) and
               (fe.Col <= ACol) and (ACol <= fe.Col2) then
            begin
              Result := true;
              exit;
            end;
        end;
      end;
    end;
  end;
  SetLength(rpnFormula, 0);
  Result := false;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a cell contains a comment and returns a pointer to the
  comment data.

  @param  ACell  Pointer to the cell
  @return Pointer to the TsComment record (nil, if the cell does not have a
          comment)
-------------------------------------------------------------------------------}
function TsWorksheet.FindComment(ACell: PCell): PsComment;
begin
  if HasComment(ACell) then
    Result := PsComment(FComments.FindByRowCol(ACell^.Row, ACell^.Col))
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a specific cell contains a comment
-------------------------------------------------------------------------------}
function TsWorksheet.HasComment(ACell: PCell): Boolean;
begin
  Result := (ACell <> nil) and (cfHasComment in ACell^.Flags);
end;

{@@ ----------------------------------------------------------------------------
  Returns the comment text attached to a specific cell

  @param  ARow   (0-based) index to the row
  @param  ACol   (0-based) index to the column
  @return Text assigned to the cell as a comment
-------------------------------------------------------------------------------}
function TsWorksheet.ReadComment(ARow, ACol: Cardinal): String;
var
  comment: PsComment;
begin
  Result := '';
  comment := PsComment(FComments.FindByRowCol(ARow, ACol));
  if comment <> nil then
    Result := comment^.Text;
end;

{@@ ----------------------------------------------------------------------------
  Returns the comment text attached to a specific cell

  @param  ACell  Pointer to the cell
  @return Text assigned to the cell as a comment
-------------------------------------------------------------------------------}
function TsWorksheet.ReadComment(ACell: PCell): String;
var
  comment: PsComment;
begin
  Result := '';
  comment := FindComment(ACell);
  if comment <> nil then
    Result := comment^.Text;
end;

{@@ ----------------------------------------------------------------------------
  Adds a comment to a specific cell

  @param  ARow   (0-based) row index of the cell
  @param  ACol   (0-based) column index of the cell
  @param  AText  Comment text
  @return Pointer to the cell containing the comment
-------------------------------------------------------------------------------}
function TsWorksheet.WriteComment(ARow, ACol: Cardinal; AText: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteComment(Result, AText);
end;

{@@ ----------------------------------------------------------------------------
  Adds a comment to a specific cell

  @param  ACell  Pointer to the cell
  @param  AText  Comment text
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteComment(ACell: PCell; AText: String);
begin
  if ACell = nil then
    exit;

  // Remove the comment if an empty string is passed
  if AText = '' then
  begin
    RemoveComment(ACell);
    exit;
  end;

  // Add new comment record
  FComments.AddComment(ACell^.Row, ACell^.Col, AText);
  Include(ACell^.Flags, cfHasComment);

  ChangedCell(ACell^.Row, ACell^.Col);

end;


{ Hyperlinks }

{@@ ----------------------------------------------------------------------------
  Checks whether the specified cell contains a hyperlink and returns a pointer
  to the hyperlink data.

  @param  ACell  Pointer to the cell
  @return Pointer to the TsHyperlink record, or NIL if the cell does not contain
          a hyperlink.
-------------------------------------------------------------------------------}
function TsWorksheet.FindHyperlink(ACell: PCell): PsHyperlink;
begin
  if HasHyperlink(ACell) then
    Result := PsHyperlink(FHyperlinks.FindByRowCol(ACell^.Row, ACell^.Col))
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified cell contains a hyperlink
-------------------------------------------------------------------------------}
function TsWorksheet.HasHyperlink(ACell: PCell): Boolean;
begin
  Result := (ACell <> nil) and (cfHyperlink in ACell^.Flags);
end;

{@@ ----------------------------------------------------------------------------
  Reads the hyperlink information of a specified cell.

  @param   ACell         Pointer to the cell considered
  @returns Record with the hyperlink data assigned to the cell.
           If the cell is not a hyperlink the result field Kind is hkNone.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadHyperlink(ACell: PCell): TsHyperlink;
var
  hyperlink: PsHyperlink;
begin
  hyperlink := FindHyperlink(ACell);
  if hyperlink <> nil then
    Result := hyperlink^
  else
  begin
    Result.Row := ACell^.Row;
    Result.Col := ACell^.Col;
    Result.Target := '';
    Result.Tooltip := '';
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes a hyperlink from the specified cell. Releaes memory occupied by
  the associated TsHyperlink record. Cell content type is converted to
  cctUTF8String.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveHyperlink(ACell: PCell);
begin
  if HasHyperlink(ACell) then
  begin
    FHyperlinks.DeleteHyperlink(ACell^.Row, ACell^.Col);
    Exclude(ACell^.Flags, cfHyperlink);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the passed string represents a valid hyperlink target

  @param   AValue  String to be checked. Must be either a fully qualified URI,
                   a local relative (!) file name, or a # followed by a cell
                   address in the current workbook
  @param   AErrMsg Error message in case that the string is not correct.
  @returns TRUE if the string is correct, FALSE otherwise
-------------------------------------------------------------------------------}
function TsWorksheet.ValidHyperlink(AValue: String; out AErrMsg: String): Boolean;
var
  u: TUri;
  sheet: TsWorksheet;
  r, c: Cardinal;
begin
  Result := false;
  AErrMsg := '';
  if AValue = '' then
  begin
    AErrMsg := rsEmptyHyperlink;
    exit;
  end else
  if (AValue[1] = '#') then
  begin
    Delete(AValue, 1, 1);
    if not FWorkbook.TryStrToCell(AValue, sheet, r, c) then
    begin
      AErrMsg := Format(rsNoValidHyperlinkInternal, ['#'+AValue]);
      exit;
    end;
  end else
  begin
    u := ParseURI(AValue);
    if SameText(u.Protocol, 'mailto') then
    begin
      Result := true;   // To do: Check email address here...
      exit;
    end else
    if SameText(u.Protocol, 'file') then
    begin
      if FilenameIsAbsolute(u.Path + u.Document) then
      begin
        Result := true;
        exit;
      end else
      begin
        AErrMsg := Format(rsLocalfileHyperlinkAbs, [AValue]);
        exit;
      end;
    end else
    begin
      Result := true;
      exit;
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Assigns a hyperlink to the cell at the specified row and column
  Cell content is not affected by the presence of a hyperlink.

  @param  ARow          Row index of the cell considered
  @param  ACol          Column index of the cell considered
  @param  ATarget       Hyperlink address given as a fully qualitifed URI for
                        external links, or as a # followed by a cell address
                        for internal links.
  @param  ATooltip      Text for popup tooltip hint used by Excel
  @returns Pointer to the cell with the hyperlink
-------------------------------------------------------------------------------}
function TsWorksheet.WriteHyperlink(ARow, ACol: Cardinal; ATarget: String;
  ATooltip: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteHyperlink(Result, ATarget, ATooltip);
end;

{@@ ----------------------------------------------------------------------------
  Assigns a hyperlink to the specified cell.

  @param  ACell         Pointer to the cell considered
  @param  ATarget       Hyperlink address given as a fully qualitifed URI for
                        external links, or as a # followed by a cell address
                        for internal links. Local files can be specified also
                        by their name relative to the workbook.
                        An existing hyperlink is removed if ATarget is empty.
  @param  ATooltip      Text for popup tooltip hint used by Excel
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteHyperlink(ACell: PCell; ATarget: String;
  ATooltip: String = '');

  function GetDisplayText(ATarget: String): String;
  var
    target, bm: String;
  begin
    SplitHyperlink(ATarget, target, bm);
    if pos('file:', lowercase(ATarget))=1 then
    begin
      URIToFilename(target, Result);
      ForcePathDelims(Result);
      if bm <> '' then Result := Result + '#' + bm;
    end else
    if target = '' then
      Result := bm
    else
      Result := ATarget;
  end;

var
  fmt: TsCellFormat;
  noCellText: Boolean = false;
begin
  if ACell = nil then
    exit;

  fmt := ReadCellFormat(ACell);

  // Empty target string removes the hyperlink. Resets the font from hyperlink
  // to default font.
  if ATarget = '' then begin
    RemoveHyperlink(ACell);
    if fmt.FontIndex = HYPERLINK_FONTINDEX then
      WriteFont(ACell, DEFAULT_FONTINDEX);
    exit;
  end;

  // Detect whether the cell already has a hyperlink, but has no other content.
  if HasHyperlink(ACell) then
    noCellText := (ACell^.ContentType = cctUTF8String) and
      (GetDisplayText(ReadHyperlink(ACell).Target) = ReadAsUTF8Text(ACell));

  // Attach the hyperlink to the cell
  FHyperlinks.AddHyperlink(ACell^.Row, ACell^.Col, ATarget, ATooltip);
  Include(ACell^.Flags, cfHyperlink);

  // If there is no other cell content use the target as cell label string.
  if (ACell^.ContentType = cctEmpty) or noCellText then
  begin
    ACell^.ContentType := cctUTF8String;
    ACell^.UTF8StringValue := GetDisplayText(ATarget);
  end;

  // Select the hyperlink font.
  if fmt.FontIndex = DEFAULT_FONTINDEX then
  begin
    fmt.FontIndex := HYPERLINK_FONTINDEX;
    Include(fmt.UsedFormattingFields, uffFont);
    ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt);
  end;

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Is called whenever a cell value or formatting has changed. Fires an event
  "OnChangeCell". This is handled by TsWorksheetGrid to update the grid cell.

  @param  ARow   Row index of the cell which has been changed
  @param  ACol   Column index of the cell which has been changed
-------------------------------------------------------------------------------}
procedure TsWorksheet.ChangedCell(ARow, ACol: Cardinal);
begin
  if FWorkbook.FReadWriteFlag = rwfRead then
    exit;

  if (FWorkbook.FCalculationLock = 0) and (boAutoCalc in FWorkbook.Options) then
  begin
    if CellUsedInFormula(ARow, ACol) then
      CalcFormulas;
  end;
  if Assigned(FOnChangeCell) then
    FOnChangeCell(Self, ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Is called whenever a font height changes. Fires an even "OnChangeFont"
  which is handled by TsWorksheetGrid to update the row heights.

  @param  ARow  Row index of the cell for which the font height has changed
  @param  ACol  Column index of the cell for which the font height has changed.
-------------------------------------------------------------------------------}
procedure TsWorksheet.ChangedFont(ARow, ACol: Cardinal);
begin
  if FWorkbook.FReadWriteFlag = rwfRead then
    exit;
  if Assigned(FOnChangeFont) then
    FOnChangeFont(Self, ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Copies a cell to a cell at another location. The new cell has the same values
  and the same formatting. It differs in formula (adapted relative references)
  and col/row indexes.

  @param   FromCell   Pointer to the source cell which will be copied
  @param   ToCell     Pointer to the destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyCell(AFromCell, AToCell: PCell);
var
  toRow, toCol: LongInt;
  row1, col1, row2, col2: Cardinal;
  hyperlink: PsHyperlink;
begin
  if (AFromCell = nil) or (AToCell = nil) then
    exit;

  // Remember the row and column indexes of the destination cell.
  toRow := AToCell^.Row;
  toCol := AToCell^.Col;

  // Copy cell values and formats
  AToCell^ := AFromCell^;

  // Fix row and column indexes overwritten
  AToCell^.Row := toRow;
  AToCell^.Col := toCol;
  AToCell^.Worksheet := self;

  // Fix relative references in formulas
  // This also fires the OnChange event.
  CopyFormula(AFromCell, AToCell);

  // Copy cell format
  CopyFormat(AFromCell, AToCell);

  // Merged?
  if IsMergeBase(AFromCell) then
  begin
    FindMergedRange(AFromCell, row1, col1, row2, col2);
    MergeCells(toRow, toCol, toRow + LongInt(row2) - LongInt(row1), toCol + LongInt(col2) - LongInt(col1));
  end;

  // Copy comment
  if HasComment(AFromCell) then
    WriteComment(AToCell, ReadComment(AFromCell));

  // Copy hyperlink
  hyperlink := FindHyperlink(AFromCell);
  if hyperlink <> nil then
    WriteHyperlink(AToCell, hyperlink^.Target, hyperlink^.Tooltip);

  // Notify visual controls of possibly changed row heights.
  ChangedFont(AToCell^.Row, AToCell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Copies a cell. The source cell can be located in a different worksheet, while
  the destination cell must be in the same worksheet which calls the methode.

  @param AFromRow  Row index of the source cell
  @param AFromCol  Column index of the source cell
  @param AToRow    Row index of the destination cell
  @param AToCol    Column index of the destination cell
  @param AFromWorksheet  Worksheet containing the source cell. Self, if omitted.
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyCell(AFromRow, AFromCol, AToRow, AToCol: Cardinal;
  AFromWorksheet: TsWorksheet = nil);
var
  srcCell, destCell: PCell;
begin
  if AFromWorksheet = nil then
    AFromWorksheet := self;

  srcCell := AFromWorksheet.FindCell(AFromRow, AFromCol);
  destCell := GetCell(AToRow, AToCol);

  CopyCell(srcCell, destCell);

  ChangedCell(AToRow, AToCol);
  ChangedFont(AToRow, AToCol);
end;

{@@ ----------------------------------------------------------------------------
  Copies all format parameters from the format cell to another cell.

  @param AFromCell  Pointer to source cell
  @param AToCell    Pointer to destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyFormat(AFromCell, AToCell: PCell);
begin
  if (AFromCell = nil) or (AToCell = nil) then
    exit;

  CopyCellFormat(AFromCell, AToCell);

  ChangedCell(AToCell^.Row, AToCell^.Col);
  ChangedFont(AToCell^.Row, AToCell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Copies all format parameters from a given cell to another cell identified
  by its row/column indexes.

  @param  AFormatCell Pointer to the source cell from which the format is copied.
  @param  AToRow      Row index of the destination cell
  @param  AToCol      Column index of the destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyFormat(AFormatCell: PCell; AToRow, AToCol: Cardinal);
begin
  CopyFormat(AFormatCell, GetCell(AToRow, AToCol));
end;

{@@ ----------------------------------------------------------------------------
  Copies the formula of a specified cell to another cell. Adapts relative
  cell references to the new cell.

  @param  AFromCell  Pointer to the source cell from which the formula is to be
                     copied
  @param  AToCell    Pointer to the destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyFormula(AFromCell, AToCell: PCell);
var
  rpnFormula: TsRPNFormula;
begin
  if (AFromCell = nil) or (AToCell = nil) then
    exit;

  if AFromCell^.FormulaValue = '' then
    AToCell^.FormulaValue := ''
  else
  begin
    // Here we convert the formula to an rpn formula as seen from source...
    rpnFormula := BuildRPNFormula(AFromCell, AToCell);
    // ... and here we reconstruct the string formula as seen from destination cell.
    AToCell^.FormulaValue := ConvertRPNFormulaToStringFormula(rpnFormula);
  end;

  ChangedCell(AToCell^.Row, AToCell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Copies the formula of a specified cell to another cell given by its row and
  column index. Relative cell references are adapted to the new cell.

  @param  AFormatCell Pointer to the source cell containing the formula to be
                      copied
  @param  AToRow      Row index of the destination cell
  @param  AToCol      Column index of the destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyFormula(AFormulaCell: PCell; AToRow, AToCol: Cardinal);
begin
  CopyFormula(AFormulaCell, GetCell(AToRow, AToCol));
end;

{@@ ----------------------------------------------------------------------------
  Copies the value of a specified cell to another cell (without copying
  formulas or formats)

  @param  AFromCell  Pointer to the source cell providing the value to be copied
  @param  AToCell    Pointer to the destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyValue(AFromCell, AToCell: PCell);
begin
  if (AFromCell = nil) or (AToCell = nil) then
    exit;

  CopyCellValue(AFromCell, AToCell);

  ChangedCell(AToCell^.Row, AToCell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Copies the value of a specified cell to another cell given by its row and
  column index

  @param  AValueCell  Pointer to the cell containing the value to be copied
  @param  AToRow      Row index of the destination cell
  @param  AToCol      Column index of the destination cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.CopyValue(AValueCell: PCell; AToRow, AToCol: Cardinal);
begin
  CopyValue(AValueCell, GetCell(AToRow, AToCol));
end;

{@@ ----------------------------------------------------------------------------
  Deletes a specified cell. If the cell belongs to a merged block its content
  and formatting is erased. Otherwise the cell is destroyed and its memory is
  released.
-------------------------------------------------------------------------------}
procedure TsWorksheet.DeleteCell(ACell: PCell);
{$warning TODO: Shift cells to the right/below !!! ??? }
begin
  if ACell = nil then
    exit;

  // Does cell have a comment? -->  remove it
  if HasComment(ACell) then
    WriteComment(ACell, '');

  // Cell is part of a merged block? --> Erase content, formatting etc.
  if IsMerged(ACell)  then
  begin
    EraseCell(ACell);
    exit;
  end;

  // Destroy the cell, and remove it from the tree
  RemoveAndFreeCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Internal call-back procedure for looping through all cells when deleting
  a specified column. Deletion happens in DeleteCol BEFORE this callback!
-------------------------------------------------------------------------------}
procedure TsWorksheet.DeleteColCallback(data, arg: Pointer);
var
  cell: PCell;
  col: Cardinal;
  formula: TsRPNFormula;
  i: Integer;
begin
  col := LongInt({%H-}PtrInt(arg));
  cell := PCell(data);
  if cell = nil then   // This should not happen. Just to make sure...
    exit;

  // Update column index of moved cell
  if (cell^.Col > col) then
    dec(cell^.Col);

  // Update formulas
  if HasFormula(cell) then
  begin
    // (1) create an rpn formula
    formula := BuildRPNFormula(cell);
    // (2) update cell addresses affected by the deletion of the column
    for i:=0 to High(formula) do
    begin
      if (formula[i].ElementKind in [fekCell, fekCellRef, fekCellRange]) then
      begin
        if formula[i].Col = col then
        begin
          formula[i].ElementKind := fekErr;
          formula[i].IntValue := ord(errIllegalRef);
        end else
        if formula[i].Col > col then
          dec(formula[i].Col);
        if (formula[i].ElementKind = fekCellRange) then
        begin
          if (formula[i].Col2 = col) then
          begin
            formula[i].ElementKind := fekErr;
            formula[i].IntValue := ord(errIllegalRef);
          end
          else
          if (formula[i].Col2 > col) then
            dec(formula[i].Col2);
        end;
      end;
    end;
    // (3) convert rpn formula back to string formula
    cell^.FormulaValue := ConvertRPNFormulaToStringFormula(formula);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Internal call-back procedure for looping through all cells when deleting
  a specified row. Deletion happens in DeleteRow BEFORE this callback!
-------------------------------------------------------------------------------}
procedure TsWorksheet.DeleteRowCallback(data, arg: Pointer);
var
  cell: PCell;
  row: Cardinal;
  formula: TsRPNFormula;
  i: Integer;
begin
  row := LongInt({%H-}PtrInt(arg));
  cell := PCell(data);
  if cell = nil then   // This should not happen. Just to make sure...
    exit;

  // Update row index of moved cell
  if (cell^.Row > row) then
    dec(cell^.Row);

  // Update formulas
  if HasFormula(cell) then
  begin
    // (1) create an rpn formula
    formula := BuildRPNFormula(cell);
    // (2) update cell addresses affected by the deletion of the column
    for i:=0 to High(formula) do
    begin
      if (formula[i].ElementKind in [fekCell, fekCellRef, fekCellRange]) then
      begin
        if formula[i].Row = row then
        begin
          formula[i].ElementKind := fekErr;
          formula[i].IntValue := ord(errIllegalRef);
        end else
        if formula[i].Row > row then
          dec(formula[i].Row);
        if (formula[i].ElementKind = fekCellRange) then
        begin
          if (formula[i].Row2 = row) then
          begin
            formula[i].ElementKind := fekErr;
            formula[i].IntValue := ord(errIllegalRef);
          end
          else
          if (formula[i].Row2 > row) then
            dec(formula[i].Row2);
        end;
      end;
    end;
    // (3) convert rpn formula back to string formula
    cell^.FormulaValue := ConvertRPNFormulaToStringFormula(formula);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Erases content and formatting of a cell. The cell still occupies memory.

  @param  ACell  Pointer to cell to be erased.
-------------------------------------------------------------------------------}
procedure TsWorksheet.EraseCell(ACell: PCell);
var
  r, c: Cardinal;
begin
  if ACell <> nil then begin
    r := ACell^.Row;
    c := ACell^.Col;

    // Unmerge range if the cell is the base of a merged block
    if IsMergeBase(ACell) then
      UnmergeCells(r, c);

    // Remove the comment if the cell has one
    if HasComment(ACell) then
      WriteComment(r, c, '');

    // Erase all cell content
    InitCell(r, c, ACell^);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Exchanges two cells

  @param  ARow1   Row index of the first cell
  @param  ACol1   Column index of the first cell
  @param  ARow2   Row index of the second cell
  @param  ACol2   Column index of the second cell

  @note          This method does not take care of merged cells and does not
                 check for this situation. Therefore, the method is not public!
-------------------------------------------------------------------------------}
procedure TsWorksheet.ExchangeCells(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  FCells.Exchange(ARow1, ACol1, ARow2, ACol2);
  FComments.Exchange(ARow1, ACol1, ARow2, ACol2);
  FHyperlinks.Exchange(ARow1, ACol1, ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Adds a new cell at a specified row and column index to the Cells list.

  NOTE: It is not checked if there exists already another cell at this location.
  This case must be avoided. USE CAREFULLY WITHOUT FindCell
  (e.g., during reading into empty worksheets).
-------------------------------------------------------------------------------}
function TsWorksheet.AddCell(ARow, ACol: Cardinal): PCell;
begin
  Result := Cells.AddCell(ARow, ACol);
  if FFirstColIndex = $FFFFFFFF then FFirstColIndex := GetFirstColIndex(true)
    else FFirstColIndex := Min(FFirstColIndex, ACol);
  if FFirstRowIndex = $FFFFFFFF then FFirstRowIndex := GetFirstRowIndex(true)
    else FFirstRowIndex := Min(FFirstRowIndex, ARow);
  if FLastColIndex = 0 then FLastColIndex := GetLastColIndex(true)
    else FLastColIndex := Max(FLastColIndex, ACol);
  if FLastRowIndex = 0 then FLastRowIndex := GetLastRowIndex(true)
    else FLastRowIndex := Max(FLastRowIndex, ARow);
end;

{@@ ----------------------------------------------------------------------------
  Tries to locate a Cell in the list of already written Cells

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @return Pointer to the cell if found, or nil if not found
  @see    TCell
-------------------------------------------------------------------------------}
function TsWorksheet.FindCell(ARow, ACol: Cardinal): PCell;
begin
  Result := PCell(FCells.FindByRowCol(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Tries to locate a cell in the list of already written cells

  @param  AddressStr  Address of the cell in Excel A1 notation
  @return Pointer to the cell if found, or nil if not found
  @see    TCell
-------------------------------------------------------------------------------}
function TsWorksheet.FindCell(AddressStr: String): PCell;
var
  r, c: Cardinal;
begin
  if ParseCellString(AddressStr, r, c) then
    Result := FindCell(r, c)
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Obtains an allocated cell at the desired location.

  If the cell already exists, a pointer to it will be returned.

  If not, then new memory for the cell will be allocated, a pointer to it
  will be returned and it will be added to the list of cells.

  @param  ARow      Row index of the cell
  @param  ACol      Column index of the cell

  @return A pointer to the cell at the desired location.

  @see    TCell
-------------------------------------------------------------------------------}
function TsWorksheet.GetCell(ARow, ACol: Cardinal): PCell;
begin
  Result := Cells.FindCell(ARow, ACol);
  if Result = nil then
    Result := AddCell(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Obtains an allocated cell at the desired location.

  If the Cell already exists, a pointer to it will be returned.

  If not, then new memory for the cell will be allocated, a pointer to it
  will be returned and it will be added to the list of cells.

  @param  AddressStr  Address of the cell in Excel A1 notation (an exception is
                      raised in case on an invalid cell address).
  @return A pointer to the cell at the desired location.

  @see    TCell
-------------------------------------------------------------------------------}
function TsWorksheet.GetCell(AddressStr: String): PCell;
var
  r, c: Cardinal;
begin
  if ParseCellString(AddressStr, r, c) then
    Result := GetCell(r, c)
  else
    raise Exception.CreateFmt(rsNoValidCellAddress, [AddressStr]);
end;

{@@ ----------------------------------------------------------------------------
  Returns the number of cells in the worksheet with contents.

  @return The number of cells with contents in the worksheet
-------------------------------------------------------------------------------}
function TsWorksheet.GetCellCount: Cardinal;
begin
  Result := FCells.Count;
end;

{@@ ----------------------------------------------------------------------------
  Determines the number of decimals displayed for the number in the cell

  @param  ACell            Pointer to the cell under investigation
  @return Number of decimals places used in the string display of the cell.
-------------------------------------------------------------------------------}
function TsWorksheet.GetDisplayedDecimals(ACell: PCell): Byte;
var
  i, p: Integer;
  s: String;
begin
  Result := 0;
  if (ACell <> nil) and (ACell^.ContentType = cctNumber) then
  begin
    s := ReadAsUTF8Text(ACell);
    p := pos(Workbook.FormatSettings.DecimalSeparator, s);
    if p > 0 then
    begin
      i := p+1;
      while (i <= Length(s)) and (s[i] in ['0'..'9']) do inc(i);
      Result := i - (p+1);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines some number format attributes (decimal places, currency symbol) of
  a cell

  @param  ACell            Pointer to the cell under investigation
  @param  ADecimals        Number of decimal places that can be extracted from
                           the formatting string, e.g. in case of '0.000' this
                           would be 3.
  @param  ACurrencySymbol  String representing the currency symbol extracted from
                           the formatting string.

  @return true if the the format string could be analyzed successfully, false if not
-------------------------------------------------------------------------------}
function TsWorksheet.GetNumberFormatAttributes(ACell: PCell; out ADecimals: byte;
  out ACurrencySymbol: String): Boolean;
var
  parser: TsNumFormatParser;
  nf: TsNumberFormat;
  nfs: String;
begin
  Result := false;
  if ACell <> nil then
  begin
    ReadNumFormat(ACell, nf, nfs);
    parser := TsNumFormatParser.Create(nfs, FWorkbook.FormatSettings);
    try
      if parser.Status = psOK then
      begin
        nf := parser.NumFormat;
        if (nf = nfGeneral) and (ACell^.ContentType = cctNumber) then
        begin
          ADecimals := GetDisplayedDecimals(ACell);
          ACurrencySymbol := '';
        end else
        if IsDateTimeFormat(nf) then
        begin
          ADecimals := 2;
          ACurrencySymbol := '?';
        end
        else
        begin
          ADecimals := parser.Decimals;
          ACurrencySymbol := parser.CurrencySymbol;
        end;
        Result := true;
      end;
    finally
      parser.Free;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the 0-based index of the first column with a cell with contents.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the first column is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the first column.
  @see GetCellCount
-------------------------------------------------------------------------------}
function TsWorksheet.GetFirstColIndex(AForceCalculation: Boolean = false): Cardinal;
var
  cell: PCell;
  i: Integer;
begin
  if AForceCalculation then
  begin
    Result := Cardinal(-1);
    for cell in FCells do
      Result := Math.Min(Result, cell^.Col);
    // In addition, there may be column records defining the column width even
    // without content
    for i:=0 to FCols.Count-1 do
      if FCols[i] <> nil then
        Result := Math.Min(Result, PCol(FCols[i])^.Col);
    // Store the result
    FFirstColIndex := Result;
  end
  else
  begin
    Result := FFirstColIndex;
    if Result = cardinal(-1) then
      Result := GetFirstColIndex(true);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the 0-based index of the last column containing a cell with a
  column record (due to content or formatting), or containing a Col record.

  If no cells have contents or there are no column records, zero will be
  returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the last column is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the last column.
  @see GetCellCount
  @see GetLastOccupiedColIndex
-------------------------------------------------------------------------------}
function TsWorksheet.GetLastColIndex(AForceCalculation: Boolean = false): Cardinal;
var
  i: Integer;
begin
  if AForceCalculation then
  begin
    // Traverse the tree from lowest to highest.
    // Since tree primary sort order is on row highest col could exist anywhere.
    Result := GetLastOccupiedColIndex;
    // In addition, there may be column records defining the column width even
    // without cells
    for i:=0 to FCols.Count-1 do
      if FCols[i] <> nil then
        Result := Math.Max(Result, PCol(FCols[i])^.Col);
    // Store the result
    FLastColIndex := Result;
  end
  else
    Result := FLastColIndex;
end;

{@@ ----------------------------------------------------------------------------
  Deprecated, use GetLastColIndex instead

  @see GetLastColIndex
-------------------------------------------------------------------------------}
function TsWorksheet.GetLastColNumber: Cardinal;
begin
  Result := GetLastColIndex;
end;

{@@ ----------------------------------------------------------------------------
  Returns the 0-based index of the last column with a cell with contents.
  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @see GetCellCount
  @see GetLastColIndex
-------------------------------------------------------------------------------}
function TsWorksheet.GetLastOccupiedColIndex: Cardinal;
var
  cell: PCell;
begin
  Result := 0;
  // Traverse the tree from lowest to highest.
  // Since tree's primary sort order is on row, highest col could exist anywhere.
  for cell in FCells do
    Result := Math.Max(Result, cell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Returns the 0-based index of the first row with a cell with data or formatting.
  If no cells have contents, -1 will be returned.

  @param  AForceCalculation  The index of the first row is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the first row.
  @see GetCellCount
-------------------------------------------------------------------------------}
function TsWorksheet.GetFirstRowIndex(AForceCalculation: Boolean = false): Cardinal;
var
  cell: PCell;
  i: Integer;
begin
  if AForceCalculation then
  begin
    Result := $FFFFFFFF;
    cell := FCells.GetFirstCell;
    if cell <> nil then Result := cell^.Row;
    // In addition, there may be row records even for rows without cells.
    for i:=0 to FRows.Count-1 do
      if FRows[i] <> nil then
        Result := Math.Min(Result, PRow(FRows[i])^.Row);
    // Store result
    FFirstRowIndex := Result;
  end
  else
  begin
    Result := FFirstRowIndex;
    if Result = Cardinal(-1) then
      Result := GetFirstRowIndex(true);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the 0-based index of the last row with a cell with contents or with
  a ROW record.

  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @param  AForceCalculation  The index of the last row is continuously updated
                             whenever a new cell is created. If AForceCalculation
                             is true all cells are scanned to determine the index
                             of the last row.
  @see GetCellCount
  @see GetLastOccupiedRowIndex
-------------------------------------------------------------------------------}
function TsWorksheet.GetLastRowIndex(AForceCalculation: Boolean = false): Cardinal;
var
  i: Integer;
begin
  if AForceCalculation then
  begin
    // Index of highest row with at least one existing cell
    Result := GetLastOccupiedRowIndex;
    // In addition, there may be row records even for empty rows.
    for i:=0 to FRows.Count-1 do
      if FRows[i] <> nil then
        Result := Math.Max(Result, PRow(FRows[i])^.Row);
    // Store result
    FLastRowIndex := Result;
  end
  else
    Result := FLastRowIndex
end;

{@@ ----------------------------------------------------------------------------
  Returns the 0-based index of the last row with a cell with contents.
  If no cells have contents, zero will be returned, which is also a valid value.

  Use GetCellCount to verify if there is at least one cell with contents in the
  worksheet.

  @see GetCellCount
  @see GetLastRowIndex
-------------------------------------------------------------------------------}
function TsWorksheet.GetLastOccupiedRowIndex: Cardinal;
var
  cell: PCell;
begin
  Result := 0;
  cell := FCells.GetLastCell;
  if Assigned(cell) then
    Result := cell^.Row;
end;

{@@ ----------------------------------------------------------------------------
  Deprecated, use GetLastColIndex instead

  @see GetLastColIndex
-------------------------------------------------------------------------------}
function TsWorksheet.GetLastRowNumber: Cardinal;
begin
  Result := GetLastRowIndex;
end;

{@@ ----------------------------------------------------------------------------
  Reads the contents of a cell and returns an user readable text
  representing the contents of the cell.

  The resulting string is UTF-8 encoded.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @return The text representation of the cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsUTF8Text(ARow, ACol: Cardinal): string; //ansistring;
begin
  Result := ReadAsUTF8Text(GetCell(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Reads the contents of a cell and returns an user readable text
  representing the contents of the cell.

  The resulting string is UTF-8 encoded.

  @param  ACell     Pointer to the cell
  @return The text representation of the cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsUTF8Text(ACell: PCell): string; //ansistring;
begin
  Result := ReadAsUTF8Text(ACell, FWorkbook.FormatSettings);
end;

{@@ ----------------------------------------------------------------------------
  Reads the contents of a cell and returns an user readable text
  representing the contents of the cell.

  The resulting string is UTF-8 encoded.

  @param  ACell            Pointer to the cell
  @param  AFormatSettings  Format settings to be used for string conversion
                           of numbers and date/times.
  @return The text representation of the cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsUTF8Text(ACell: PCell;
  AFormatSettings: TFormatSettings): string;
var
  fmt: PsCellFormat;
  hyperlink: PsHyperlink;
  numFmt: TsNumFormatParams;
  nf: TsNumberFormat;
  nfs: String;

begin
  Result := '';
  if ACell = nil then
    Exit;

  fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
  numFmt := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);

  with ACell^ do
    case ContentType of
      cctUTF8String:
        Result := UTF8StringValue;

      cctNumber:
        Result := ConvertFloatToStr(NumberValue, numFmt, AFormatSettings);

      cctDateTime:
        if Assigned(numFmt) then
          Result := ConvertFloatToStr(DateTimeValue, numFmt, AFormatSettings)
        else
        if not IsNaN(DateTimeValue) then
        begin
          if frac(DateTimeValue) = 0 then  // date only
            nf := nfShortDate
          else
          if trunc(DateTimeValue) = 0 then  // time only
            nf := nfLongTime
          else
            nf := nfShortDateTime;
          nfs := BuildDateTimeFormatString(nf, AFormatSettings);
          Result := FormatDateTime(nfs, DateTimeValue, AFormatSettings);
        end;

      cctBool:
        Result := StrUtils.IfThen(BoolValue, rsTRUE, rsFALSE);

      cctError:
        case TsErrorValue(ErrorValue) of
          errEmptyIntersection  : Result := rsErrEmptyIntersection;
          errDivideByZero       : Result := rsErrDivideByZero;
          errWrongType          : Result := rsErrWrongType;
          errIllegalRef         : Result := rsErrIllegalRef;
          errWrongName          : Result := rsErrWrongName;
          errOverflow           : Result := rsErrOverflow;
          errArgError           : Result := rsErrArgError;
          errFormulaNotSupported: Result := rsErrFormulaNotSupported;
        end;

      else   // blank --> display hyperlink target if available
        Result := '';
        if HasHyperlink(ACell) then
        begin
          hyperlink := FindHyperlink(ACell);
          if hyperlink <> nil then Result := hyperlink^.Target;
        end;
    end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the value of a cell as a number.

  If the cell contains a date/time value its serial value is returned
  (as FPC TDateTime).

  If the cell contains a text value it is attempted to convert it to a number.

  If the cell is empty or its contents cannot be represented as a number the
  value 0.0 is returned.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @return Floating-point value representing the cell contents, or 0.0 if cell
          does not exist or its contents cannot be converted to a number.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsNumber(ARow, ACol: Cardinal): Double;
begin
  Result := ReadAsNumber(FindCell(ARow, ACol));
end;

{@@ ----------------------------------------------------------------------------
  Returns the value of a cell as a number.

  If the cell contains a date/time value its serial value is returned
  (as FPC TDateTime).

  If the cell contains a text value it is attempted to convert it to a number.

  If the cell is empty or its contents cannot be represented as a number the
  value 0.0 is returned.

  @param  ACell     Pointer to the cell
  @return Floating-point value representing the cell contents, or 0.0 if cell
          does not exist or its contents cannot be converted to a number.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsNumber(ACell: PCell): Double;
begin
  Result := 0.0;
  if ACell = nil then
    exit;

  case ACell^.ContentType of
    cctDateTime:
      Result := ACell^.DateTimeValue; //this is in FPC TDateTime format, not Excel
    cctNumber:
      Result := ACell^.NumberValue;
    cctUTF8String:
      if not TryStrToFloat(ACell^.UTF8StringValue, Result, FWorkbook.FormatSettings)
        then Result := 0.0;
    cctBool:
      if ACell^.BoolValue then Result := 1.0 else Result := 0.0;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the contents of a cell and returns the date/time value of the cell.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AResult   Date/time value of the cell (or 0.0, if no date/time cell)
  @return True if the cell is a datetime value, false otherwise
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsDateTime(ARow, ACol: Cardinal;
  out AResult: TDateTime): Boolean;
begin
  Result := ReadAsDateTime(FindCell(ARow, ACol), AResult);
end;

{@@ ----------------------------------------------------------------------------
  Reads the contents of a cell and returns the date/time value of the cell.

  @param  ACell     Pointer to the cell
  @param  AResult   Date/time value of the cell (or 0.0, if no date/time cell)
  @return True if the cell is a datetime value, false otherwise
-------------------------------------------------------------------------------}
function TsWorksheet.ReadAsDateTime(ACell: PCell;
  out AResult: TDateTime): Boolean;
begin
  if (ACell = nil) or (ACell^.ContentType <> cctDateTime) then
  begin
    AResult := 0;
    Result := False;
    Exit;
  end;

  AResult := ACell^.DateTimeValue;
  Result := True;
end;

{@@ ----------------------------------------------------------------------------
  If a cell contains a formula (string formula or RPN formula) the formula
  is returned as a string in Excel syntax.

  @param   ACell      Pointer to the cell considered
  @param   ALocalized If true, the formula is returned with decimal and list
                      separators accoding to the workbook's FormatSettings.
                      Otherwise it uses dot and comma, respectively.
  @return  Formula string in Excel syntax (does not contain a leading "=")
-------------------------------------------------------------------------------}
function TsWorksheet.ReadFormulaAsString(ACell: PCell;
  ALocalized: Boolean = false): String;
var
  parser: TsSpreadsheetParser;
begin
  Result := '';
  if ACell = nil then
    exit;
  if HasFormula(ACell) then begin
    if ALocalized then
    begin
      // case (1): Formula is localized and has to be converted to default syntax   // !!!! Is this comment correct?
      parser := TsSpreadsheetParser.Create(self);
      try
        parser.Expression := ACell^.FormulaValue;
        Result := parser.LocalizedExpression[Workbook.FormatSettings];
      finally
        parser.Free;
      end;
    end
    else
      // case (2): Formula is already in default syntax
      Result := ACell^.FormulaValue;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns to numeric equivalent of the cell contents. This is the NumberValue
  of a number cell, the DateTimeValue of a date/time cell, the ordinal BoolValue
  of a boolean cell, or the string converted to a number of a string cell.
  All other cases return NaN.

  @param   ACell   Cell to be considered
  @param   AValue  (output) extracted numeric value
  @return  True if conversion to number is successful, otherwise false
-------------------------------------------------------------------------------}
function TsWorksheet.ReadNumericValue(ACell: PCell; out AValue: Double): Boolean;
begin
  AValue := NaN;
  if ACell <> nil then begin
    Result := True;
    case ACell^.ContentType of
      cctNumber:
        AValue := ACell^.NumberValue;
      cctDateTime:
        AValue := ACell^.DateTimeValue;
      cctBool:
        AValue := ord(ACell^.BoolValue);
      else
        if (ACell^.ContentType <> cctUTF8String) or
           not TryStrToFloat(ACell^.UTF8StringValue, AValue) or
           not TryStrToDateTime(ACell^.UTF8StringValue, AValue)
        then
          Result := False;
      end;
  end else
    Result := False;
end;

{@@ ----------------------------------------------------------------------------
  Converts an RPN formula (as read from an xls biff file, for example) to a
  string formula.

  @param    AFormula  Array of rpn formula tokens
  @return   Formula string in Excel syntax (without leading "=")
-------------------------------------------------------------------------------}
function TsWorksheet.ConvertRPNFormulaToStringFormula(const AFormula: TsRPNFormula): String;
var
  parser: TsSpreadsheetParser;
begin
  Result := '';

  parser := TsSpreadsheetParser.Create(self);
  try
    parser.RPNFormula := AFormula;
    Result := parser.Expression;
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the CalcState flag of the specified cell. This flag tells whether a
  formula in the cell has not yet been calculated (csNotCalculated), is
  currently being calculated (csCalculating), or has already been calculated
  (csCalculated.

  @param   ACell   Pointer to cell considered
  @return  Enumerated value of the cell's calculation state
           (csNotCalculated, csCalculating, csCalculated)
-------------------------------------------------------------------------------}
function TsWorksheet.GetCalcState(ACell: PCell): TsCalcState;
var
  calcState: TsCellFlags;
begin
  Result := csNotCalculated;
  if (ACell = nil) then
    exit;
  calcState := ACell^.Flags * [cfCalculating, cfCalculated];
  if calcState = [] then
    Result := csNotCalculated
  else
  if calcState = [cfCalculating] then
    Result := csCalculating
  else
  if calcState = [cfCalculated] then
    Result := csCalculated
  else
    raise Exception.Create('[TsWorksheet.GetCalcState] Illegal cell flags.');
end;

{@@ ----------------------------------------------------------------------------
  Set the CalcState flag of the specified cell. This flag tells whether a
  formula in the cell has not yet been calculated (csNotCalculated), is
  currently being calculated (csCalculating), or has already been calculated
  (csCalculated).

  For internal use only!

  @param  ACell   Pointer to cell considered
  @param  AValue  New value for the calculation state of the cell
                  (csNotCalculated, csCalculating, csCalculated)
-------------------------------------------------------------------------------}
procedure TsWorksheet.SetCalcState(ACell: PCell; AValue: TsCalcState);
begin
  case AValue of
    csNotCalculated:
      ACell^.Flags := ACell^.Flags - [cfCalculated, cfCalculating];
    csCalculating:
      ACell^.Flags := ACell^.Flags + [cfCalculating] - [cfCalculated];
    csCalculated:
      ACell^.Flags := ACell^.Flags + [cfCalculated] - [cfCalculating];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the set of used formatting fields of a cell.

  Each cell contains a set of "used formatting fields". Formatting is applied
  only if the corresponding element is contained in the set.

  @param  ACell   Pointer to the cell
  @return Set of elements used in formatting the cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadUsedFormatting(ACell: PCell): TsUsedFormattingFields;
var
  fmt: PsCellFormat;
begin
  if ACell = nil then
  begin
    Result := [];
    Exit;
  end;
  fmt := FWorkbook.GetPointerToCellFormat(ACell^.FormatIndex);
  Result := fmt^.UsedFormattingFields;
end;

{@@ ----------------------------------------------------------------------------
  Returns the background fill pattern and colors of a cell.

  @param  ACell  Pointer to the cell
  @return TsFillPattern record (or EMPTY_FILL, if the cell does not have a
          filled background
-------------------------------------------------------------------------------}
function TsWorksheet.ReadBackground(ACell: PCell): TsFillPattern;
var
  fmt : PsCellFormat;
begin
  Result := EMPTY_FILL;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffBackground in fmt^.UsedFormattingFields) then
      Result := fmt^.Background;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the background color of a cell as rbg value

  @param    ACell  Pointer to the cell
  @return   Value containing the rgb bytes in little-endian order
-------------------------------------------------------------------------------}
function TsWorksheet.ReadBackgroundColor(ACell: PCell): TsColor;
var
  fmt: PsCellFormat;
begin
  Result := scTransparent;
  if ACell <> nil then
  begin
    fmt :=  Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffBackground in fmt^.UsedFormattingFields) then
    begin
      if (fmt^.Background.Style = fsSolidFill) then
        Result := fmt^.Background.FgColor
      else
        Result := fmt^.Background.BgColor;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines which borders are drawn around a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellBorders(ACell: PCell): TsCellBorders;
var
  fmt: PsCellFormat;
begin
  Result := [];
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffBorder in fmt^.UsedFormattingFields) then
      Result := fmt^.Border;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines which the style of a particular cell border
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellBorderStyle(ACell: PCell;
  ABorder: TsCelLBorder): TsCellBorderStyle;
var
  fmt: PsCellFormat;
begin
  Result := DEFAULT_BORDERSTYLES[ABorder];
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := fmt^.BorderStyles[ABorder];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines which all border styles of a given cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellBorderStyles(ACell: PCell): TsCellBorderStyles;
var
  fmt: PsCellFormat;
  b: TsCellBorder;
begin
  Result := DEFAULT_BORDERSTYLES;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    for b in fmt.Border do
      Result[b] := fmt^.BorderStyles[b];
  end;
end;

{@@ ----------------------------------------------------------------------------
  Determines the font used by a specified cell. Returns the workbook's default
  font if the cell does not exist. Considers the uffBold and uffFont formatting
  fields of the cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellFont(ACell: PCell): TsFont;
var
  fmt: PsCellFormat;
begin
  Result := nil;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := Workbook.GetFont(fmt^.FontIndex);
  end;
  if Result = nil then
    Result := Workbook.GetDefaultFont;
end;

{@@ ----------------------------------------------------------------------------
  Returns the format record that is assigned to a specified cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadCellFormat(ACell: PCell): TsCellFormat;
begin
  Result := Workbook.GetCellFormat(ACell^.FormatIndex);
end;

{@@ ----------------------------------------------------------------------------
  Returns the horizontal alignment of a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadHorAlignment(ACell: PCell): TsHorAlignment;
var
  fmt: PsCellFormat;
begin
  Result := haDefault;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffHorAlign in fmt^.UsedFormattingFields) then
      Result := fmt^.HorAlignment;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the number format type and format string used in a specific cell
-------------------------------------------------------------------------------}
procedure TsWorksheet.ReadNumFormat(ACell: PCell; out ANumFormat: TsNumberFormat;
  out ANumFormatStr: String);
var
  fmt: PsCellFormat;
  numFmt: TsNumFormatParams;
begin
  ANumFormat := nfGeneral;
  ANumFormatStr := '';
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffNumberFormat in fmt^.UsedFormattingFields) then
    begin
      numFmt := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
      if numFmt <> nil then
      begin
        ANumFormat := numFmt.NumFormat;
        ANumFormatStr := numFmt.NumFormatStr;
      end else
      begin
        ANumFormat := nfGeneral;
        ANumFormatStr := '';
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the text orientation of a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadTextRotation(ACell: PCell): TsTextRotation;
var
  fmt: PsCellFormat;
begin
  Result := trHorizontal;
  if ACell <> nil then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffTextRotation in fmt^.UsedFormattingFields) then
      Result := fmt^.TextRotation;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the vertical alignment of a specific cell
-------------------------------------------------------------------------------}
function TsWorksheet.ReadVertAlignment(ACell: PCell): TsVertAlignment;
var
  fmt: PsCellFormat;
begin
  Result := vaDefault;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    if (uffVertAlign in fmt^.UsedFormattingFields) then
      Result := fmt^.VertAlignment;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns whether a specific cell support word-wrapping.
-------------------------------------------------------------------------------}
function TsWorksheet.ReadWordwrap(ACell: PCell): boolean;
var
  fmt: PsCellFormat;
begin
  Result := false;
  if (ACell <> nil) then
  begin
    fmt := Workbook.GetPointerToCellFormat(ACell^.FormatIndex);
    Result := uffWordwrap in fmt^.UsedFormattingFields;
  end;
end;


{ Merged cells }

{@@ ----------------------------------------------------------------------------
  Finds the upper left cell of a merged block to which a specified cell belongs.
  This is the "merge base". Returns nil if the cell is not merged.

  @param  ACell  Cell under investigation
  @return A pointer to the cell in the upper left corner of the merged block
          to which ACell belongs.
          If ACell is isolated then the function returns nil.
-------------------------------------------------------------------------------}
function TsWorksheet.FindMergeBase(ACell: PCell): PCell;
var
  rng: PsCellRange;
begin
  Result := nil;
  if IsMerged(ACell) then
  begin
    rng := FMergedCells.FindRangeWithCell(ACell^.Row, ACell^.Col);
    if rng <> nil then
      Result := FindCell(rng^.Row1, rng^.Col1);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Merges adjacent individual cells to a larger single cell

  @param  ARow1   Row index of the upper left corner of the cell range
  @param  ACol1   Column index of the upper left corner of the cell range
  @param  ARow2   Row index of the lower right corner of the cell range
  @param  ACol2   Column index of the lower right corner of the cell range
-------------------------------------------------------------------------------}
procedure TsWorksheet.MergeCells(ARow1, ACol1, ARow2, ACol2: Cardinal);
var
  rng: PsCellRange;
  cell: PCell;
  r, c: Cardinal;
begin
  // A single cell cannot be merged
  if (ARow1 = ARow2) and (ACol1 = ACol2) then
    exit;

  // Is cell ARow1/ACol1 already the base of a merged range? ...
  rng := PsCellRange(FMergedCells.FindByRowCol(ARow1, ACol1));
  // ... no: --> Add a new merged range
  if rng = nil then
    FMergedCells.AddRange(ARow1, ACol1, ARow2, ACol2)
  else
  // ... yes: --> modify the merged range accordingly
  begin
    // unmark previously merged range
    for cell in Cells.GetRangeEnumerator(rng^.Row1, rng^.Col1, rng^.Row2, rng^.Col2) do
      Exclude(cell^.Flags, cfMerged);
    // Define new limits of merged range
    rng^.Row2 := ARow2;
    rng^.Col2 := ACol2;
  end;

  // Mark all cells in the range as "merged"
  for r := ARow1 to ARow2 do
    for c := ACol1 to ACol2 do
    begin
      cell := GetCell(r, c);   // if not existent create new cell
      Include(cell^.Flags, cfMerged);
    end;

  ChangedCell(ARow1, ACol1);
end;

{@@ ----------------------------------------------------------------------------
  Merges adjacent individual cells to a larger single cell

  @param  ARange  Cell range string given in Excel notation (e.g: A1:D5).
                  A non-range string (e.g. A1) is not allowed.
-------------------------------------------------------------------------------}
procedure TsWorksheet.MergeCells(ARange: String);
var
  r1, r2, c1, c2: Cardinal;
begin
  if ParseCellRangeString(ARange, r1, c1, r2, c2) then
    MergeCells(r1, c1, r2, c2);
end;

{@@ ----------------------------------------------------------------------------
  Disconnects merged cells to make them individual cells again.

  Input parameter is a cell which belongs to the range to be unmerged.

  @param  ARow   Row index of a cell considered to belong to the cell block
  @param  ACol   Column index of a cell considered to belong to the cell block
-------------------------------------------------------------------------------}
procedure TsWorksheet.UnmergeCells(ARow, ACol: Cardinal);
var
  rng: PsCellRange;
  cell: PCell;
begin
  rng := FMergedCells.FindRangeWithCell(ARow, ACol);
  if rng <> nil then
  begin
    // Remove the "merged" flag from the cells in the merged range to make them
    // isolated again...
    for cell in Cells.GetRangeEnumerator(rng^.Row1, rng^.Col1, rng^.Row2, rng^.Col2) do
      Exclude(cell^.Flags, cfMerged);
    // ... and delete the range
    FMergedCells.DeleteRange(rng^.Row1, rng^.Col1);
  end;

  ChangedCell(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Disconnects merged cells to make them individual cells again.

  @param  ARange  Cell (range) string given in Excel notation (e.g: A1, or A1:D5)
                  In case of a range string, only the upper left corner cell is
                  considered. It must belong to the merged range of cells to be
                  unmerged.
-------------------------------------------------------------------------------}
procedure TsWorksheet.UnmergeCells(ARange: String);
var
  sheet: TsWorksheet;
  rng: TsCellRange;
begin
  if Workbook.TryStrToCellRange(ARange, sheet, rng) then
    UnmergeCells(rng.Row1, rng.Col1);
end;

{@@ ----------------------------------------------------------------------------
  Determines the merged cell block to which a particular cell belongs

  @param   ACell  Pointer to the cell being investigated
  @param   ARow1  (output) Top row index of the merged block
  @param   ACol1  (outout) Left column index of the merged block
  @param   ARow2  (output) Bottom row index of the merged block
  @param   ACol2  (output) Right column index of the merged block

  @return  True if the cell belongs to a merged block, False if not, or if the
           cell does not exist at all.
-------------------------------------------------------------------------------}
function TsWorksheet.FindMergedRange(ACell: PCell;
  out ARow1, ACol1, ARow2, ACol2: Cardinal): Boolean;
var
  rng: PsCellRange;
begin
  if IsMerged(ACell) then
  begin
    rng := FMergedCells.FindRangeWithCell(ACell^.Row, ACell^.Col);
    if rng <> nil then
    begin
      ARow1 := rng^.Row1;
      ACol1 := rng^.Col1;
      ARow2 := rng^.Row2;
      ACol2 := rng^.Col2;
      Result := true;
      exit;
    end;
  end;
  Result := false;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the two specified cells belong to the same merged cell block.

  @param   ACell1  Pointer to the first cell
  @param   ACell2  Pointer to the second cell
  @reult   TRUE if both cells belong to the same merged cell block
           FALSE if the cells are not merged or are in different blocks
-------------------------------------------------------------------------------}
function TsWorksheet.InSameMergedRange(ACell1, ACell2: PCell): Boolean;
begin
  Result := IsMerged(ACell1) and IsMerged(ACell2) and
            (FindMergeBase(ACell1) = FindMergeBase(ACell2));
end;

{@@ ----------------------------------------------------------------------------
  Returns true if the specified cell is the base of a merged cell range, i.e.
  the upper left corner of that range.

  @param   ACell  Pointer to the cell being considered
  @return  True if the cell is the upper left corner of a merged range
           False if not
-------------------------------------------------------------------------------}
function TsWorksheet.IsMergeBase(ACell: PCell): Boolean;
begin
  Result := (ACell <> nil) and (ACell = FindMergeBase(ACell));
end;

{@@ ----------------------------------------------------------------------------
  Returns TRUE if the specified cell belongs to a merged block

  @param   ACell  Pointer to the cell of interest
  @return  TRUE if the cell belongs to a merged block, FALSE if not.
-------------------------------------------------------------------------------}
function TsWorksheet.IsMerged(ACell: PCell): Boolean;
begin
  Result := (ACell <> nil) and (cfMerged in ACell^.Flags);
end;

{@@ ----------------------------------------------------------------------------
  Removes the comment from a cell and releases the memory occupied by the node.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveComment(ACell: PCell);
begin
  if HasComment(ACell) then
  begin
    FComments.DeleteComment(ACell^.Row, ACell^.Col);
    Exclude(ACell^.Flags, cfHasComment);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes a cell from its tree container. DOES NOT RELEASE ITS MEMORY!

  @param  ARow   Row index of the cell to be removed
  @param  ACol   Column index of the cell to be removed
  @return  Pointer to the cell removed
-------------------------------------------------------------------------------}
function TsWorksheet.RemoveCell(ARow, ACol: Cardinal): PCell;
begin
  Result := PCell(FCells.FindByRowCol(ARow, ACol));
  if Result <> nil then FCells.Remove(Result);
end;

{@@ ----------------------------------------------------------------------------
  Removes a cell and releases its memory. If a comment is attached to the
  cell then it is removed and releaded as well.

  Just for internal usage since it does not modify the other cells affected.
  And it does not change other records depending on the cell (comments,
  merged ranges etc).

  @param  ARow   Row index of the cell to be removed
  @param  ACol   Column index of the cell to be removed
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveAndFreeCell(ARow, ACol: Cardinal);
begin
  FCells.DeleteCell(ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Setter for the worksheet name property. Checks if the name is valid, and
  exits without any change if not. Creates an event OnChangeWorksheet.
-------------------------------------------------------------------------------}
procedure TsWorksheet.SetName(const AName: String);
begin
  if AName = FName then
    exit;
  if (FWorkbook <> nil) then //and FWorkbook.ValidWorksheetName(AName) then
  begin
    FName := AName;
    if (FWorkbook.FLockCount = 0) and Assigned(FWorkbook.FOnChangeWorksheet) then
      FWorkbook.FOnRenameWorksheet(FWorkbook, self);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Compare function for sorting of rows and columns called directly by Sort()
  The compare algorithm starts with the first key parameters. If cells are
  found to be "equal" the next parameter is set is used until a difference is
  found, or all parameters are used.

  @param   ARow1         Row index of the first cell to be compared
  @param   ACol1         Column index of the first cell to be compared
  @param   ARow2         Row index of the second cell to be compared
  @parem   ACol2         Column index of the second cell to be compared
  @param   ASortOptions  Sorting options: case-insensitive and/or descending
  @return  -1 if the first cell is "smaller", i.e. is sorted in front of the
              second one
           +1 if the first cell is "larger", i.e. is behind the second one
           0  if both cells are equal
------------------------------------------------------------------------------- }
function TsWorksheet.DoCompareCells(ARow1, ACol1, ARow2, ACol2: Cardinal;
  ASortOptions: TsSortOptions): Integer;
var
  cell1, cell2: PCell;  // Pointers to the cells to be compared
  key: Integer;
begin
  cell1 := FindCell(ARow1, ACol1);
  cell2 := FindCell(ARow2, ACol2);
  Result := DoInternalCompareCells(cell1, cell2, ASortOptions);
  if Result = 0 then begin
    key := 1;
    while (Result = 0) and (key <= High(FSortParams.Keys)) do
    begin
      if FSortParams.SortByCols then
      begin
        cell1 := FindCell(ARow1, FSortParams.Keys[key].ColRowIndex);
        cell2 := FindCell(ARow2, FSortParams.Keys[key].ColRowIndex);
      end else
      begin
        cell1 := FindCell(FSortParams.Keys[key].ColRowIndex, ACol1);
        cell2 := FindCell(FSortParams.Keys[key].ColRowIndex, ACol2);
      end;
      Result := DoInternalCompareCells(cell1, cell2, ASortOptions);
      inc(key);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Compare function for sorting of rows and columns. Called by DoCompareCells.

  @param    ACell1        Pointer to the first cell of the comparison
  @param    ACell2        Pointer to the second cell of the comparison
  @param    ASortOptions  Options for sorting: descending and/or case-insensitive
  @return   -1 if the first cell is "smaller"
            +1 if the first cell is "larger",
            0 if both cells are "equal"

            Date/time and boolean cells are sorted like number cells according
            to their number value
            Label cells are sorted as UTF8 strings.

            In case of mixed cell content types the order is determined by
            the parameter Priority of the SortParams.
            Empty cells are always at the end (in both ascending and descending
            order)
-------------------------------------------------------------------------------}
function TsWorksheet.DoInternalCompareCells(ACell1, ACell2: PCell;
  ASortOptions: TsSortOptions): Integer;
// Sort priority in Excel:
// numbers < alpha < blank (ascending)
// alpha < numbers < blank (descending)
var
  number1, number2: Double;
begin
  result := 0;
  if Assigned(OnCompareCells) then
    OnCompareCells(Self, ACell1, ACell2, Result)
  else
  begin
    if (ACell1 = nil) and (ACell2 = nil) then
      Result := 0
    else
    if (ACell1 = nil) or (ACell1^.ContentType = cctEmpty) then
    begin
      Result := +1;   // Empty cells go to the end
      exit;           // Avoid SortOrder to bring the empty cell to the top!
    end else
    if (ACell2 = nil) or (ACell2^.ContentType = cctEmpty) then
    begin
      Result := -1;   // Empty cells go to the end
      exit;           // Avoid SortOrder to bring the empty cell to the top!
    end else
    if (ACell1^.ContentType = cctEmpty) and (ACell2^.ContentType = cctEmpty) then
      Result := 0
    else
    if (ACell1^.ContentType = cctUTF8String) and (ACell2^.ContentType = cctUTF8String) then
    begin
      if ssoCaseInsensitive in ASortOptions then
        Result := UTF8CompareText(ACell1^.UTF8StringValue, ACell2^.UTF8StringValue)
      else
        Result := UTF8CompareStr(ACell1^.UTF8StringValue, ACell2^.UTF8StringValue);
    end else
    if (ACell1^.ContentType = cctUTF8String) and (ACell2^.ContentType <> cctUTF8String) then
      case FSortParams.Priority of
        spNumAlpha: Result := +1;  // numbers before text
        spAlphaNum: Result := -1;  // text before numbers
      end
    else
    if (ACell1^.ContentType <> cctUTF8String) and (ACell2^.ContentType = cctUTF8String) then
      case FSortParams.Priority of
        spNumAlpha: Result := -1;
        spAlphaNum: Result := +1;
      end
    else
    begin
      ReadNumericValue(ACell1, number1);
      ReadNumericValue(ACell2, number2);
      Result := CompareValue(number1, number2);
    end;
  end;
  if ssoDescending in ASortOptions then
    Result := -Result;
end;

{@@ ----------------------------------------------------------------------------
  Exchanges columns or rows, depending on value of "AIsColumn"

  @param  AIsColumn   if true the exchange is done for columns, otherwise for rows
  @param  AIndex      Index of the column (if AIsColumn is true) or the row
                      (if AIsColumn is false) which is to be exchanged with the
                      one having index "WidthIndex"
  @param  WithIndex   Index of the column (if AIsColumn is true) or the row
                      (if AIsColumn is false) with which "AIndex" is to be
                      replaced.
  @param  AFromIndex  First row (if AIsColumn is true) or column (if AIsColumn
                      is false) which is affected by the exchange
  @param  AToIndex    Last row (if AIsColumn is true) or column (if AsColumn is
                      false) which is affected by the exchange
-------------------------------------------------------------------------------}
procedure TsWorksheet.DoExchangeColRow(AIsColumn: Boolean;
  AIndex, WithIndex: Cardinal; AFromIndex, AToIndex: Cardinal);
var
  r, c: Cardinal;
begin
  if AIsColumn then
    for r := AFromIndex to AToIndex do
      ExchangeCells(r, AIndex, r, WithIndex)
  else
    for c := AFromIndex to AToIndex do
      ExchangeCells(AIndex, c, WithIndex, c);
end;

{@@ ----------------------------------------------------------------------------
  Sorts a range of cells defined by the cell rectangle from ARowFrom/AColFrom
  to ARowTo/AColTo according to the parameters specified in ASortParams

  @param  ASortParams   Set of parameters to define sorting along rows or colums,
                        the sorting key column or row indexes, and the sorting
                        directions
  @param  ARange        Cell range to be sorted, in Excel notation, such as 'A1:C8'
-------------------------------------------------------------------------------}
procedure TsWorksheet.Sort(ASortParams: TsSortParams; ARange: String);
var
  r1,c1, r2,c2: Cardinal;
begin
  if ParseCellRangeString(ARange, r1, c1, r2, c2) then
    Sort(ASortParams, r1, c1, r2, c2)
  else
    raise Exception.CreateFmt(rsNoValidCellRangeAddress, [ARange]);
end;

{@@ ----------------------------------------------------------------------------
  Sorts a range of cells defined by the cell rectangle from ARowFrom/AColFrom
  to ARowTo/AColTo according to the parameters specified in ASortParams

  @param  ASortParams   Set of parameters to define sorting along rows or colums,
                        the sorting key column or row indexes, and the sorting
                        directions
  @param  ARowFrom      Top row of the range to be sorted
  @param  AColFrom      Left column of the range to be sorted
  @param  ARowTo        Last row of the range to be sorted
  @param  AColTo        Right column of the range to be sorted
-------------------------------------------------------------------------------}
procedure TsWorksheet.Sort(const ASortParams: TsSortParams;
  ARowFrom, AColFrom, ARowTo, AColTo: Cardinal);
// code "borrowed" from grids.pas and adapted to multi-key sorting

  procedure QuickSort(L,R: Integer);
  var
    I,J: Integer;
    P: Integer;
    index: Integer;
    options: TsSortOptions;
  begin
    index := ASortParams.Keys[0].ColRowIndex;   // less typing...
    options := ASortParams.Keys[0].Options;
    repeat
      I := L;
      J := R;
      P := (L + R) div 2;
      repeat
        if ASortParams.SortByCols then
        begin
          while DoCompareCells(P, index, I, index, options) > 0 do inc(I);
          while DoCompareCells(P, index, J, index, options) < 0 do dec(J);
        end else
        begin
          while DoCompareCells(index, P, index, I, options) > 0 do inc(I);
          while DoCompareCells(index, P, index, J, options) < 0 do dec(J);
        end;

        if I <= J then
        begin
          if I <> J then
          begin
            if ASortParams.SortByCols then
            begin
              if DoCompareCells(I, index, J, index, options) <> 0 then
                DoExchangeColRow(not ASortParams.SortByCols, J,I, AColFrom, AColTo);
            end else
            begin
              if DoCompareCells(index, I, index, J, options) <> 0 then
                DoExchangeColRow(not ASortParams.SortByCols, J,I, ARowFrom, ARowTo);
            end;
          end;

          if P = I then
            P := J
          else
          if P = J then
            P := I;

          inc(I);
          dec(J);
        end;
      until I > J;

      if L < J then
        QuickSort(L, J);

      L := I;
    until I >= R;
  end;

  function ContainsMergedCells: boolean;
  var
    cell: PCell;
  begin
    result := false;
    for cell in Cells.GetRangeEnumerator(ARowFrom, AColFrom, ARowTo, AColTo) do
      if IsMerged(cell) then
        exit(true);
  end;

begin
  if ContainsMergedCells then
    raise Exception.Create(rsCannotSortMerged);

  FSortParams := ASortParams;
  if ASortParams.SortByCols then
    QuickSort(ARowFrom, ARowTo)
  else
    QuickSort(AColFrom, AColTo);
  ChangedCell(ARowFrom, AColFrom);
end;


{@@ ----------------------------------------------------------------------------
  Marks a specified cell as "selected". Only needed by the visual controls.
-------------------------------------------------------------------------------}
procedure TsWorksheet.SelectCell(ARow, ACol: Cardinal);
begin
  FActiveCellRow := ARow;
  FActiveCellCol := ACol;
  if Assigned(FOnSelectCell) then
    FOnSelectCell(Self, ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Clears the list of seleccted cell ranges
  Only needed by the visual controls.
-------------------------------------------------------------------------------}
procedure TsWorksheet.ClearSelection;
begin
  SetLength(FSelection, 0);
end;

{@@ ----------------------------------------------------------------------------
  Returns the list of selected cell ranges
-------------------------------------------------------------------------------}
function TsWorksheet.GetSelection: TsCellRangeArray;
var
  i: Integer;
begin
  SetLength(Result, Length(FSelection));
  for i:=0 to High(FSelection) do
    Result[i] := FSelection[i];
end;

{@@ ----------------------------------------------------------------------------
  Returns all selection ranges as an Excel string
-------------------------------------------------------------------------------}
function TsWorksheet.GetSelectionAsString: String;
const
  RELATIVE = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];
var
  i: Integer;
  L: TStringList;
begin
  L := TStringList.Create;
  try
    for i:=0 to Length(FSelection)-1 do
      with FSelection[i] do
        L.Add(GetCellRangeString(Row1, Col1, Row2, Col2, RELATIVE, true));
    L.Delimiter := DefaultFormatSettings.ListSeparator;
    L.StrictDelimiter := true;
    Result := L.DelimitedText;
  finally
    L.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the number of selected cell ranges
-------------------------------------------------------------------------------}
function TsWorksheet.GetSelectionCount: Integer;
begin
  Result := Length(FSelection);
end;

{@@ ----------------------------------------------------------------------------
  Marks an array of cell ranges as "selected". Only needed for visual controls
-------------------------------------------------------------------------------}
procedure TsWorksheet.SetSelection(const ASelection: TsCellRangeArray);
var
  i: Integer;
begin
  SetLength(FSelection, Length(ASelection));
  for i:=0 to High(FSelection) do
    FSelection[i] := ASelection[i];
end;

{@@ ----------------------------------------------------------------------------
  Helper method to update internal caching variables
-------------------------------------------------------------------------------}
procedure TsWorksheet.UpdateCaches;
begin
  FFirstColIndex := GetFirstColIndex(true);
  FFirstRowIndex := GetFirstRowIndex(true);
  FLastColIndex := GetLastColIndex(true);
  FLastRowIndex := GetLastRowIndex(true);
end;

{@@ ----------------------------------------------------------------------------
  Writes UTF-8 encoded text to a cell.

  On formats that don't support unicode, the text will be converted
  to ISO Latin 1.

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AText     The text to be written encoded in utf-8
  @return Pointer to cell created or used
-------------------------------------------------------------------------------}
function TsWorksheet.WriteUTF8Text(ARow, ACol: Cardinal; AText: ansistring): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteUTF8Text(Result, AText);
end;

{@@ ----------------------------------------------------------------------------
  Writes UTF-8 encoded text to a cell.

  On formats that don't support unicode, the text will be converted
  to ISO Latin 1.

  @param  ACell     Pointer to the cell
  @param  AText     The text to be written encoded in utf-8
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteUTF8Text(ACell: PCell; AText: ansistring);
var
  r, c: Cardinal;
  hyperlink: TsHyperlink;
begin
  if ACell = nil then
    exit;

  if (AText = '') and HasHyperlink(ACell) then
  begin
    hyperlink := ReadHyperlink(ACell);
    AText := hyperlink.Target;
    if pos('file:', hyperlink.Target)=1 then
    begin
      URIToFileName(AText, AText);
//      Delete(AText, 1, Length('file:///'));
      ForcePathDelims(AText);
    end;
  end;

  if (AText = '') then
  begin
    if (Workbook.GetCellFormat(ACell^.FormatIndex).UsedFormattingFields = []) and
       (ACell^.Flags * [cfHyperlink, cfHasComment, cfMerged] = []) and
       (ACell^.FormulaValue = '')
    then
    begin
      r := ACell^.Row;
      c := ACell^.Col;
      RemoveCell(r, c);
      ChangedCell(r, c);
    end else
      WriteBlank(ACell);
    exit;
  end;

  ACell^.ContentType := cctUTF8String;
  ACell^.UTF8StringValue := AText;

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Writes a floating-point number to a cell, does not change the number format

  @param  ARow         Cell row index
  @param  ACol         Cell column index
  @param  ANumber      Number to be written
  @return Pointer to cell created or used
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumber(Result, ANumber);
end;

{@@ ----------------------------------------------------------------------------
  Writes a floating-point number to a cell, does not change the number format

  @param  ACell        Pointer to the cell
  @param  ANumber      Number to be written
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: Double);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a floating-point number to a cell

  @param  ARow         Cell row index
  @param  ACol         Cell column index
  @param  ANumber      Number to be written
  @param  ANumFormat   Identifier for a built-in number format, e.g. nfFixed (optional)
  @param  ADecimals    Number of decimal places used for formatting (optional)
  @return Pointer to cell created or used
  @see    TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: double;
  ANumFormat: TsNumberFormat; ADecimals: Byte = 2): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumber(Result, ANumber, ANumFormat, ADecimals);
end;

{@@ ----------------------------------------------------------------------------
  Writes a floating-point number to a cell

  @param  ACell        Pointer to the cell
  @param  ANumber      Number to be written
  @param  ANumFormat   Identifier for a built-in number format, e.g. nfFixed
  @param  ADecimals    Optional number of decimal places used for formatting
                       If ANumFormat is nfFraction the ADecimals defines the
                       digits of Numerator and denominator.
  @see TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: Double;
  ANumFormat: TsNumberFormat; ADecimals: Byte = 2);
var
  fmt: TsCellFormat;
  nfs: String;
begin
  if IsDateTimeFormat(ANumFormat) or IsCurrencyFormat(ANumFormat) then
    raise Exception.Create(rsInvalidNumberFormat);

  if ACell <> nil then begin
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;

    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.NumberFormat := ANumFormat;
    if ANumFormat <> nfGeneral then begin
      Include(fmt.UsedFormattingFields, uffNumberFormat);
      if ANumFormat = nfFraction then
      begin
        if ADecimals = 0 then ADecimals := 1;
        nfs := '# ' + DupeString('?', ADecimals) + '/' + DupeString('?', ADecimals);
      end else
        nfs := BuildNumberFormatString(fmt.NumberFormat, Workbook.FormatSettings, ADecimals);
      fmt.NumberFormatIndex := Workbook.AddNumberFormat(nfs);
    end else begin
      Exclude(fmt.UsedFormattingFields, uffNumberFormat);
      fmt.NumberFormatIndex := -1;
    end;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a floating point number to the cell and uses a custom number format
  specified by the format string.
  Note that fpspreadsheet may not be able to detect the formatting when reading
  the file.

  @param  ARow              Cell row index
  @param  ACol              Cell column index
  @param  ANumber           Number to be written
  @param  ANumFormat        Format identifier (nfCustom)
  @param  ANumFormatString  String of formatting codes (such as 'dd/mmm'
  @return Pointer to cell created or used
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumber(ARow, ACol: Cardinal; ANumber: Double;
  ANumFormat: TsNumberFormat; ANumFormatString: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumber(Result, ANumber, ANumFormat, ANumFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Writes a floating point number to the cell and uses a custom number format
  specified by the format string.
  Note that fpspreadsheet may not be able to detect the formatting when reading
  the file.

  @param  ACell             Pointer to the cell considered
  @param  ANumber           Number to be written
  @param  ANumFormat        Format identifier (nfCustom)
  @param  ANumFormatString  String of formatting codes (such as 'dd/mmm' )
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumber(ACell: PCell; ANumber: Double;
  ANumFormat: TsNumberFormat; ANumFormatString: String);
var
  parser: TsNumFormatParser;
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    parser := TsNumFormatParser.Create(ANumFormatString, FWorkbook.FormatSettings);
    try
      // Format string ok?
      if parser.Status <> psOK then
        raise Exception.Create(rsNoValidNumberFormatString);
      // Make sure that we do not write a date/time value here
      if parser.IsDateTimeFormat
        then raise Exception.Create(rsInvalidNumberFormat);
      // If format string matches a built-in format use its format identifier,
      // All this is considered when calling Builtin_NumFormat of the parser.
    finally
      parser.Free;
    end;

    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := ANumber;

    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    if ANumFormat <> nfGeneral then begin
      fmt.NumberFormatIndex := Workbook.AddNumberFormat(ANumFormatString);
      Include(fmt.UsedFormattingFields, uffNumberFormat);
    end else begin
      Exclude(fmt.UsedFormattingFields, uffNumberFormat);
      fmt.NumberFormatIndex := -1;
    end;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an empty cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @return Pointer to the cell
  Note:   Empty cells are useful when, for example, a border line extends
          along a range of cells including empty cells.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBlank(ARow, ACol: Cardinal): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBlank(Result);
end;

{@@ ----------------------------------------------------------------------------
  Writes an empty cell

  @param  ACel      Pointer to the cell
  Note:   Empty cells are useful when, for example, a border line extends
          along a range of cells including empty cells.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBlank(ACell: PCell);
begin
  if ACell <> nil then begin
    ACell^.FormulaValue := '';
    if HasHyperlink(ACell) then
      WriteUTF8Text(ACell, '')  // '' will be replaced by the hyperlink target.
    else
    begin
      ACell^.ContentType := cctEmpty;
      ChangedCell(ACell^.Row, ACell^.Col);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a boolean cell

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The boolean value
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBoolValue(ARow, ACol: Cardinal; AValue: Boolean): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBoolValue(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes a boolean cell

  @param  ACell      Pointer to the cell
  @param  AValue     The boolean value
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBoolValue(ACell: PCell; AValue: Boolean);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctBool;
    ACell^.BoolValue := AValue;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes data defined as a string into a cell. Depending on the structure of the
  string, the worksheet tries to guess whether it is a number, a date/time or
  a text and calls the corresponding writing method.

  @param  ARow    Row index of the cell
  @param  ACol    Column index of the cell
  @param  AValue  Value to be written into the cell given as a string. Depending
                  on the structure of the string, however, the value is written
                  as a number, a date/time or a text.
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteCellValueAsString(ARow, ACol: Cardinal;
  AValue: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCellValueAsString(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes data defined as a string into a cell. Depending on the structure of the
  string, the worksheet tries to guess whether it is a number, a date/time or
  a text and calls the corresponding writing method.

  @param  ACell   Pointer to the cell
  @param  AValue  Value to be written into the cell given as a string. Depending
                  on the structure of the string, however, the value is written
                  as a number, a date/time or a text.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCellValueAsString(ACell: PCell; AValue: String);
var
  isPercent: Boolean;
  number: Double;
  currSym: String;
  fmt: TsCellFormat;
  numFmtParams: TsNumFormatParams;
  maxDig: Integer;
  isMixed: Boolean;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  numFmtParams := Workbook.GetNumberFormat(fmt.NumberFormatIndex);

  if AValue = '' then
  begin
    WriteUTF8Text(ACell, '');
    exit;
  end;

  isPercent := Pos('%', AValue) = Length(AValue);
  if isPercent then Delete(AValue, Length(AValue), 1);

  if TryStrToCurrency(AValue, number, currSym, FWorkbook.FormatSettings) then
  begin
    WriteCurrency(ACell, number, nfCurrencyRed, -1, currSym);
    exit;
  end;

  if TryFractionStrToFloat(AValue, number, ismixed, maxdig) then
  begin
    WriteNumber(ACell, number);
    WriteFractionFormat(ACell, ismixed, maxdig, maxdig);
    exit;
  end;

  if TryStrToFloat(AValue, number, FWorkbook.FormatSettings) then
  begin
    if isPercent then
      WriteNumber(ACell, number/100, nfPercentage)
    else
    begin
      if IsDateTimeFormat(numFmtParams) then
        WriteNumber(ACell, number, nfGeneral)
      else
        WriteNumber(ACell, number);
    end;
    exit;
  end;

  if TryStrToDateTime(AValue, number, FWorkbook.FormatSettings) then
  begin
    if number < 1.0 then          // this is a time alone
    begin
      if not IsTimeFormat(numFmtParams) then
      begin
        if IsLongTimeFormat(AValue, FWorkbook.FormatSettings.TimeSeparator) then
          WriteDateTime(ACell, number, nfLongTime)
        else
          WriteDateTime(ACell, number, nfShortTime);
      end;
    end else
    if frac(number) = 0.0 then  // this is a date alone
    begin
      if not IsDateFormat(numFmtParams) then
        WriteDateTime(ACell, number, nfShortDate);
    end else
    if not IsDateTimeFormat(fmt.NumberFormat) then
      WriteDateTime(ACell, number, nfShortDateTime)
    else
      WriteDateTime(ACell, number);
    exit;
  end;

  WriteUTF8Text(ACell, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes a currency value to a given cell. Its number format can be provided
  optionally by specifying various parameters.

  @param ARow            Cell row index
  @param ACol            Cell column index
  @param AValue          Number value to be written
  @param ANumFormat      Format identifier, must be nfCurrency, or nfCurrencyRed.
  @param ADecimals       Number of decimal places
  @param APosCurrFormat  Code specifying the order of value, currency symbol
                         and spaces (see pcfXXXX constants)
  @param ANegCurrFormat  Code specifying the order of value, currency symbol,
                         spaces, and how negative values are shown
                         (see ncfXXXX constants)
  @param ACurrencySymbol String to be shown as currency, such as '$', or 'EUR'.
                         In case of '?' the currency symbol defined in the
                         workbook's FormatSettings is used.
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
  ANumFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = 2;
  ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
  ANegCurrFormat: Integer = -1): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCurrency(Result, AValue, ANumFormat, ADecimals, ACurrencySymbol,
    APosCurrFormat, ANegCurrFormat);
end;

{@@ ----------------------------------------------------------------------------
  Writes a currency value to a given cell. Its number format can be provided
  optionally by specifying various parameters.

  @param ACell           Pointer to the cell considered
  @param AValue          Number value to be written
  @param ANumFormat      Format identifier, must be nfCurrency or nfCurrencyRed.
  @param ADecimals       Number of decimal places
  @param APosCurrFormat  Code specifying the order of value, currency symbol
                         and spaces (see pcfXXXX constants)
  @param ANegCurrFormat  Code specifying the order of value, currency symbol,
                         spaces, and how negative values are shown
                         (see ncfXXXX constants)
  @param ACurrencySymbol String to be shown as currency, such as '$', or 'EUR'.
                         In case of '?' the currency symbol defined in the
                         workbook's FormatSettings is used.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCurrency(ACell: PCell; AValue: Double;
  ANumFormat: TsNumberFormat = nfCurrency; ADecimals: Integer = -1;
  ACurrencySymbol: String = '?'; APosCurrFormat: Integer = -1;
  ANegCurrFormat: Integer = -1);
var
  nfs: String;
begin
  if ADecimals = -1 then
    ADecimals := Workbook.FormatSettings.CurrencyDecimals;
  if APosCurrFormat = -1 then
    APosCurrFormat := Workbook.FormatSettings.CurrencyFormat;
  if ANegCurrFormat = -1 then
    ANegCurrFormat := Workbook.FormatSettings.NegCurrFormat;
  if ACurrencySymbol = '?' then
    ACurrencySymbol := Workbook.FormatSettings.CurrencyString;
  RegisterCurrency(ACurrencySymbol);

  nfs := BuildCurrencyFormatString(
    ANumFormat,
    Workbook.FormatSettings,
    ADecimals,
    APosCurrFormat, ANegCurrFormat,
    ACurrencySymbol);

  WriteCurrency(ACell, AValue, ANumFormat, nfs);
end;

{@@ ----------------------------------------------------------------------------
  Writes a currency value to a given cell. Its number format is specified by
  means of a format string.

  @param ARow               Cell row index
  @param ACol               Cell column index
  @param AValue             Number value to be written
  @param ANumFormat         Format identifier, must be nfCurrency or nfCurrencyRed.
  @param ANumFormatString   String of formatting codes, including currency symbol.
                            Can contain sections for different formatting of positive
                            and negative number.
                            Example: '"EUR" #,##0.00;("EUR" #,##0.00)'
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteCurrency(ARow, ACol: Cardinal; AValue: Double;
  ANumFormat: TsNumberFormat; ANumFormatString: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteCurrency(Result, AValue, ANumFormat, ANumFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Writes a currency value to a given cell. Its number format is specified by
  means of a format string.

  @param ACell              Pointer to the cell considered
  @param AValue             Number value to be written
  @param ANumFormat         Format identifier, must be nfCurrency or nfCurrencyRed.
  @param ANumFormatString   String of formatting codes, including currency symbol.
                            Can contain sections for different formatting of positive
                            and negative number.
                            Example: '"EUR" #,##0.00;("EUR" #,##0.00)'
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCurrency(ACell: PCell; AValue: Double;
  ANumFormat: TsNumberFormat; ANumFormatString: String);
var
  fmt: TsCellFormat;
begin
  if not IsCurrencyFormat(ANumFormat) then
    raise Exception.Create('[TsWorksheet.WriteCurrency] ANumFormat can only be nfCurrency or nfCurrencyRed');

  if (ACell <> nil) then begin
    ACell^.ContentType := cctNumber;
    ACell^.NumberValue := AValue;

    fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(ANumFormatString);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt);

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value to a cell, does not change number format

  @param  ARow          The row of the cell
  @param  ACol          The column of the cell
  @param  AValue        The date/time/datetime to be written
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTime(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value to a cell. Does not change number format

  @param  ACell         Pointer to the cell considered
  @param  AValue        The date/time/datetime to be written
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDateTime(ACell: PCell; AValue: TDateTime);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctDateTime;
    ACell^.DateTimeValue := AValue;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value to a cell

  @param  ARow          The row of the cell
  @param  ACol          The column of the cell
  @param  AValue        The date/time/datetime to be written
  @param  ANumFormat    The format specifier, e.g. nfShortDate (optional)
                        If not specified format is not changed.
  @param  ANumFormatStr Format string, used only for nfCustom or nfTimeInterval.
  @return Pointer to the cell

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
  ANumFormat: TsNumberFormat; ANumFormatStr: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTime(Result, AValue, ANumFormat, ANumFormatStr);
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value to a cell

  @param  ACell         Pointer to the cell considered
  @param  AValue        The date/time/datetime to be written
  @param  ANumFormat    The format specifier, e.g. nfShortDate (optional)
                        If not specified format is not changed.
  @param  ANumFormatStr Format string, used only for nfCustom or nfTimeInterval.

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDateTime(ACell: PCell; AValue: TDateTime;
  ANumFormat: TsNumberFormat; ANumFormatStr: String = '');
var
  parser: TsNumFormatParser;
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctDateTime;
    ACell^.DateTimeValue := AValue;

    // Date/time is actually a number field in Excel.
    // To make sure it gets saved correctly, set a date format (instead of General).
    // The user can choose another date format if he wants to

    if ANumFormat = nfGeneral then begin
      if trunc(AValue) = 0 then         // time only
        ANumFormat := nfLongTime
      else if frac(AValue) = 0.0 then   // date only
        ANumFormat := nfShortDate;
    end;

    if ANumFormatStr = '' then
      ANumFormatStr := BuildDateTimeFormatString(ANumFormat, Workbook.FormatSettings, ANumFormatStr)
    else
    if ANumFormat = nfTimeInterval then
      ANumFormatStr := AddIntervalBrackets(ANumFormatStr);

    // Check whether the formatstring is for date/times.
    if ANumFormatStr <> '' then begin
      parser := TsNumFormatParser.Create(ANumFormatStr, Workbook.FormatSettings);
      try
        // Format string ok?
        if parser.Status <> psOK then
          raise Exception.Create(rsNoValidNumberFormatString);
        // Make sure that we do not use a number format for date/times values.
        if not parser.IsDateTimeFormat
          then raise Exception.Create(rsInvalidDateTimeFormat);
        // Avoid possible duplication of standard formats
        if ANumFormat = nfCustom then
          ANumFormat := parser.NumFormat;
      finally
        parser.Free;
      end;
    end;

    fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormat := ANumFormat;
    fmt.NumberFormatStr := ANumFormatStr;
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(fmt.NumberFormatStr);
    ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt);

    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value to a cell

  @param  ARow          The row index of the cell
  @param  ACol          The column index of the cell
  @param  AValue        The date/time/datetime to be written
  @param  ANumFormatStr Format string (the format identifier nfCustom is used to
                        classify the format).
  @return Pointer to the cell

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDateTime(ARow, ACol: Cardinal; AValue: TDateTime;
  ANumFormatStr: String): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTime(Result, AValue, ANumFormatStr);
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value to a cell

  @param  ACell         Pointer to the cell considered
  @param  AValue        The date/time/datetime to be written
  @param  ANumFormatStr Format string (the format identifier nfCustom is used to
                        classify the format).

  Note: at least Excel xls does not recognize a separate datetime cell type:
  a datetime is stored as a (floating point) number, and the cell is formatted
  as a date (either built-in or a custom format).
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDateTime(ACell: PCell; AValue: TDateTime;
  ANumFormatStr: String);
begin
  WriteDateTime(ACell, AValue, nfCustom, ANumFormatStr);
end;

{@@ ----------------------------------------------------------------------------
  Adds a date/time format to the formatting of a cell

  @param  ARow             The row of the cell
  @param  ACol             The column of the cell
  @param  ANumFormat       Identifier of the format to be applied (nfXXXX constant)
  @param  ANumFormatString Optional string of formatting codes. Is only considered
                           if ANumberFormat is nfCustom.
  @return Pointer to the cell

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDateTimeFormat(ARow, ACol: Cardinal;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteDateTimeFormat(Result, ANumFormat, ANumFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Adds a date/time format to the formatting of a cell

  @param  ACell             Pointer to the cell considered
  @param  ANumFormat        Identifier of the format to be applied (nxXXXX constant)
  @param  ANumFormatString  optional string of formatting codes. Is only considered
                            if ANumberFormat is nfCustom.

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDateTimeFormat(ACell: PCell;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = '');
var
  fmt: TsCellFormat;
  nfs: String;
begin
  if ACell = nil then
    exit;

  if not ((ANumFormat in [nfGeneral, nfCustom]) or IsDateTimeFormat(ANumFormat)) then
    raise Exception.Create('WriteDateTimeFormat can only be called with date/time formats.');

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  fmt.NumberFormat := ANumFormat;
  if (ANumFormat <> nfGeneral) then
  begin
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    if (ANumFormatString = '') then
      nfs := BuildDateTimeFormatString(ANumFormat, Workbook.FormatSettings)
    else
      nfs := ANumFormatString;
  end else
  begin
    Exclude(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormatStr := '';
  end;
  fmt.NumberFormat := ANumFormat;
  fmt.NumberFormatStr := nfs;
  fmt.NumberFormatIndex := Workbook.AddNumberFormat(nfs);
  ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt);

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Formats the number in a cell to show a given count of decimal places.
  Is ignored for non-decimal formats (such as most date/time formats).

  @param  ARow       Row indows of the cell considered
  @param  ACol       Column indows of the cell considered
  @param  ADecimals  Number of decimal places to be displayed
  @return Pointer to the cell
  @see    TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDecimals(ARow, ACol: Cardinal; ADecimals: Byte): PCell;
begin
  Result := FindCell(ARow, ACol);
  WriteDecimals(Result, ADecimals);
end;

{@@ ----------------------------------------------------------------------------
  Formats the number in a cell to show a given count of decimal places.
  Is ignored for non-decimal formats (such as most date/time formats).

  @param  ACell      Pointer to the cell considered
  @param  ADecimals  Number of decimal places to be displayed
  @see    TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteDecimals(ACell: PCell; ADecimals: Byte);
var
  parser: TsNumFormatParser;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams;
  numFmtStr: String;
begin
  if (ACell = nil) or (ACell^.ContentType <> cctNumber) then
    exit;

  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  numFmt := FWorkbook.GetNumberFormat(fmt.NumberFormatIndex);
  if numFmt <> nil then
    numFmtStr := numFmt.NumFormatStr
  else
    numFmtStr := '0.00';
  parser := TsNumFormatParser.Create(numFmtStr, Workbook.FormatSettings);
  try
    parser.Decimals := ADecimals;
    numFmtStr := parser.FormatString;
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(numFmtStr);
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  finally
    parser.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an error value to a cell.

  @param  ARow       The row of the cell
  @param  ACol       The column of the cell
  @param  AValue     The error code value
  @return Pointer to the cell

  @see TsErrorValue
-------------------------------------------------------------------------------}
function TsWorksheet.WriteErrorValue(ARow, ACol: Cardinal; AValue: TsErrorValue): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteErrorValue(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes an error value to a cell.

  @param  ACol       Pointer to the cell to be written
  @param  AValue     The error code value

  @see TsErrorValue
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteErrorValue(ACell: PCell; AValue: TsErrorValue);
begin
  if ACell <> nil then begin
    ACell^.ContentType := cctError;
    ACell^.ErrorValue := AValue;
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a formula to a given cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  AFormula  The formula string to be written. A leading "=" will be removed.
  @param  ALocalized If true, the formula is expected to have decimal and list
                     separators of the workbook's FormatSettings. Otherwise
                     uses dot and comma, respectively.
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFormula(ARow, ACol: Cardinal; AFormula: string;
  ALocalized: Boolean = false): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFormula(Result, AFormula, ALocalized);
end;

{@@ ----------------------------------------------------------------------------
  Writes a formula to a given cell

  @param  ACell      Pointer to the cell
  @param  AFormula   Formula string to be written. A leading '=' will be removed.
  @param  ALocalized If true, the formula is expected to have decimal and list
                     separators of the workbook's FormatSettings. Otherwise
                     uses dot and comma, respectively.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteFormula(ACell: PCell; AFormula: string;
  ALocalized: Boolean = false);
var
  parser: TsExpressionParser;
begin
  if ACell = nil then
    exit;

  // Remove '='; is not stored internally
  if (AFormula <> '') and (AFormula[1] = '=') then
    AFormula := Copy(AFormula, 2, Length(AFormula));

  // Convert "localized" formula to standard format
  if ALocalized then begin
    parser := TsSpreadsheetParser.Create(self);
    try
      parser.LocalizedExpression[Workbook.FormatSettings] := AFormula;
      AFormula := parser.Expression;
    finally
      parser.Free;
    end;
  end;

  // Store formula in cell
  ACell^.ContentType := cctFormula;
  ACell^.FormulaValue := AFormula;
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param  ARow            The row of the cell
  @param  ACol            The column of the cell
  @param  ANumFormat      Identifier of the format to be applied
  @param  ADecimals       Number of decimal places
  @param  ACurrencySymbol optional currency symbol in case of nfCurrency
  @param  APosCurrFormat  optional identifier for positive currencies
  @param  ANegCurrFormat  optional identifier for negative currencies
  @return Pointer to the cell

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumFormat: TsNumberFormat; ADecimals: Integer; ACurrencySymbol: String = '';
  APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumberFormat(Result, ANumFormat, ADecimals, ACurrencySymbol,
    APosCurrFormat, ANegCurrFormat);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param  ARow            The row of the cell
  @param  ACol            The column of the cell
  @param  ANumFormat      Identifier of the format to be applied
  @param  ADecimals       Number of decimal places
  @param  ACurrencySymbol optional currency symbol in case of nfCurrency
  @param  APosCurrFormat  optional identifier for positive currencies
  @param  ANegCurrFormat  optional identifier for negative currencies

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumberFormat(ACell: PCell;
  ANumFormat: TsNumberFormat; ADecimals: Integer; ACurrencySymbol: String = '';
  APosCurrFormat: Integer = -1; ANegCurrFormat: Integer = -1);
var
  fmt: TsCellFormat;
  fmtStr: String;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  fmt.NumberFormat := ANumFormat;
  if ANumFormat <> nfGeneral then begin
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    if IsCurrencyFormat(ANumFormat) then
    begin
      RegisterCurrency(ACurrencySymbol);
      fmtStr := BuildCurrencyFormatString(ANumFormat, Workbook.FormatSettings,
        ADecimals, APosCurrFormat, ANegCurrFormat, ACurrencySymbol);
    end else
      fmtStr := BuildNumberFormatString(ANumFormat,
        Workbook.FormatSettings, ADecimals);
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(fmtStr);
  end else begin
    Exclude(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormatIndex := -1;
  end;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Formats a number as a fraction

  @param  ARow             Row index of the cell
  @param  ACol             Column index of the cell
  @param  ANumFormat       Identifier of the format to be applied. Must be
                           either nfFraction or nfMixedFraction
  @param  ANumeratorDigts  Count of numerator digits
  @param  ADenominatorDigits Count of denominator digits
  @return Pointer to the cell

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFractionFormat(ARow, ACol: Cardinal;
  AMixedFraction: Boolean; ANumeratorDigits, ADenominatorDigits: Integer): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFractionFormat(Result, AMixedFraction, ANumeratorDigits, ADenominatorDigits);
end;

{@@ ----------------------------------------------------------------------------
  Formats a number as a fraction

  @param  ACell            Pointer to the cell to be formatted
  @param  ANumFormat       Identifier of the format to be applied. Must be
                           either nfFraction or nfMixedFraction
  @param  ANumeratorDigts  Count of numerator digits
  @param  ADenominatorDigits Count of denominator digits

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteFractionFormat(ACell: PCell;
  AMixedFraction: Boolean; ANumeratorDigits, ADenominatorDigits: Integer);
var
  fmt: TsCellFormat;
  nfs: String;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  nfs := BuildFractionFormatString(AMixedFraction, ANumeratorDigits, ADenominatorDigits);
  fmt.NumberFormatIndex := Workbook.AddNumberFormat(nfs);
  Include(fmt.UsedFormattingFields, uffNumberFormat);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param  ARow             The row of the cell
  @param  ACol             The column of the cell
  @param  ANumFormat       Identifier of the format to be applied
  @param  ANumFormatString Optional string of formatting codes. Is only considered
                           if ANumberFormat is nfCustom.
  @return Pointer to the cell

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
function TsWorksheet.WriteNumberFormat(ARow, ACol: Cardinal;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = ''): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteNumberFormat(Result, ANumFormat, ANumFormatString);
end;

{@@ ----------------------------------------------------------------------------
  Adds a number format to the formatting of a cell

  @param  ACell            Pointer to the cell considered
  @param  ANumFormat       Identifier of the format to be applied
  @param  ANumFormatString Optional string of formatting codes. Is only considered
                           if ANumberFormat is nfCustom.

  @see    TsNumberFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteNumberFormat(ACell: PCell;
  ANumFormat: TsNumberFormat; const ANumFormatString: String = '');
var
  fmt: TsCellFormat;
  fmtStr: String;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if ANumFormat <> nfGeneral then begin
    Include(fmt.UsedFormattingFields, uffNumberFormat);
    if (ANumFormatString = '') then
      fmtStr := BuildNumberFormatString(ANumFormat, Workbook.FormatSettings)
    else
      fmtStr := ANumFormatString;
    fmt.NumberFormatIndex := Workbook.AddNumberFormat(fmtStr);
  end else begin
    Exclude(fmt.UsedFormattingFields, uffNumberFormat);
    fmt.NumberFormatIndex := -1;
  end;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Writes an RPN formula to a cell. An RPN formula is an array of tokens
  describing the calculation to be performed.

  @param  ARow          Row indows of the cell considered
  @param  ACol          Column index of the cell
  @param  AFormula      Array of TsFormulaElements. The array can be created by
                        using "CreateRPNFormla".
  @return Pointer to the cell

  @see    TsNumberFormat
  @see    TsFormulaElements
  @see    CreateRPNFormula
-------------------------------------------------------------------------------}
function TsWorksheet.WriteRPNFormula(ARow, ACol: Cardinal;
  AFormula: TsRPNFormula): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteRPNFormula(Result, AFormula);
end;

{@@ ----------------------------------------------------------------------------
  Writes an RPN formula to a cell. An RPN formula is an array of tokens
  describing the calculation to be performed. In addition,the RPN formula is
  converted to a string formula.

  @param  ACell         Pointer to the cell
  @param  AFormula      Array of TsFormulaElements. The array can be created by
                        using "CreateRPNFormla".

  @see    TsNumberFormat
  @see    TsFormulaElements
  @see    CreateRPNFormula
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteRPNFormula(ACell: PCell; AFormula: TsRPNFormula);
begin
  if ACell = nil then
    exit;

  ACell^.ContentType := cctFormula;
  ACell^.FormulaValue := ConvertRPNFormulaToStringFormula(AFormula);

  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Adds font specification to the formatting of a cell. Looks in the workbook's
  FontList and creates an new entry if the font is not used so far. Returns the
  index of the font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontName   Name of the font
  @param  AFontSize   Size of the font, in points
  @param  AFontStyle  Set with font style attributes
                      (don't use those of unit "graphics" !)
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer;
begin
  Result := WriteFont(GetCell(ARow, ACol), AFontName, AFontSize, AFontStyle, AFontColor);
end;

{@@ ----------------------------------------------------------------------------
  Adds font specification to the formatting of a cell. Looks in the workbook's
  FontList and creates an new entry if the font is not used so far. Returns the
  index of the font in the font list.

  @param  ACell       Pointer to the cell considered
  @param  AFontName   Name of the font
  @param  AFontSize   Size of the font, in points
  @param  AFontStyle  Set with font style attributes
                      (don't use those of unit "graphics" !)
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFont(ACell: PCell; const AFontName: String;
  AFontSize: Single; AFontStyle: TsFontStyles; AFontColor: TsColor): Integer;
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
  begin
    Result := -1;
    Exit;
  end;

  Result := FWorkbook.FindFont(AFontName, AFontSize, AFontStyle, AFontColor);
  if Result = -1 then
    result := FWorkbook.AddFont(AFontName, AFontSize, AFontStyle, AFontColor);

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  Include(fmt.UsedFormattingFields, uffFont);
  fmt.FontIndex := Result;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Applies a font to the formatting of a cell. The font is determined by its
  index in the workbook's font list:

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontIndex  Index of the font in the workbook's font list
  @return Pointer to the cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFont(ARow, ACol: Cardinal; AFontIndex: Integer): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteFont(Result, AFontIndex);
end;

{@@ ----------------------------------------------------------------------------
  Applies a font to the formatting of a cell. The font is determined by its
  index in the workbook's font list:

  @param  ACell       Pointer to the cell considered
  @param  AFontIndex  Index of the font in the workbook's font list
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteFont(ACell: PCell; AFontIndex: Integer);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;

  if (AFontIndex < 0) or (AFontIndex >= Workbook.GetFontCount) or (AFontIndex = 4) then
    // note: Font index 4 is not defined in BIFF
    raise Exception.Create(rsInvalidFontIndex);

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  Include(fmt.UsedFormattingFields, uffFont);
  fmt.FontIndex := AFontIndex;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the text color used in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontColor  RGB value of the new text color
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontColor(ARow, ACol: Cardinal; AFontColor: TsColor): Integer;
begin
  Result := WriteFontColor(GetCell(ARow, ACol), AFontColor);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the text color used in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell
  @param  AFontColor  RGB value of the new text color
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontColor(ACell: PCell; AFontColor: TsColor): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  Result := WriteFont(ACell, fnt.FontName, fnt.Size, fnt.Style, AFontColor);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the font used in formatting of a cell considering only the font face
  and leaving font size, style and color unchanged. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AFontName   Name of the new font to be used
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontName(ARow, ACol: Cardinal; AFontName: String): Integer;
begin
  result := WriteFontName(GetCell(ARow, ACol), AFontName);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the font used in formatting of a cell considering only the font face
  and leaving font size, style and color unchanged. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell
  @param  AFontName   Name of the new font to be used
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontName(ACell: PCell; AFontName: String): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  result := WriteFont(ACell, AFontName, fnt.Size, fnt.Style, fnt.Color);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the font size in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  ASize       Size of the font to be used (in points).
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontSize(ARow, ACol: Cardinal; ASize: Single): Integer;
begin
  Result := WriteFontSize(GetCell(ARow, ACol), ASize);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the font size in formatting of a cell. Looks in the workbook's
  font list if this modified font has already been used. If not a new font entry
  is created. Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell
  @param  ASize       Size of the font to be used (in points).
  @return Index of the font in the workbook's font list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontSize(ACell: PCell; ASize: Single): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := 0;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  Result := WriteFont(ACell, fnt.FontName, ASize, fnt.Style, fnt.Color);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the font style (bold, italic, etc) in formatting of a cell.
  Looks in the workbook's font list if this modified font has already been used.
  If not a new font entry is created.
  Returns the index of this font in the font list.

  @param  ARow        The row of the cell
  @param  ACol        The column of the cell
  @param  AStyle      New font style to be used
  @return Index of the font in the workbook's font list.

  @see TsFontStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontStyle(ARow, ACol: Cardinal;
  AStyle: TsFontStyles): Integer;
begin
  Result := WriteFontStyle(GetCell(ARow, ACol), AStyle);
end;

{@@ ----------------------------------------------------------------------------
  Replaces the font style (bold, italic, etc) in formatting of a cell.
  Looks in the workbook's font list if this modified font has already been used.
  If not a new font entry is created.
  Returns the index of this font in the font list.

  @param  ACell       Pointer to the cell considered
  @param  AStyle      New font style to be used
  @return Index of the font in the workbook's font list.

  @see TsFontStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteFontStyle(ACell: PCell; AStyle: TsFontStyles): Integer;
var
  fnt: TsFont;
begin
  if ACell = nil then begin
    Result := -1;
    exit;
  end;
  fnt := ReadCellFont(ACell);
  Result := WriteFont(ACell, fnt.FontName, fnt.Size, AStyle, fnt.Color);
end;

{@@ ----------------------------------------------------------------------------
  Adds text rotation to the formatting of a cell

  @param  ARow      The row of the cell
  @param  ACol      The column of the cell
  @param  ARotation How to rotate the text
  @return Pointer to cell

  @see    TsTextRotation
-------------------------------------------------------------------------------}
function TsWorksheet.WriteTextRotation(ARow, ACol: Cardinal;
  ARotation: TsTextRotation): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteTextRotation(Result, ARotation);
end;

{@@ ----------------------------------------------------------------------------
  Adds text rotation to the formatting of a cell

  @param  ACell      Pointer to the cell
  @param  ARotation  How to rotate the text

  @see    TsTextRotation
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteTextRotation(ACell: PCell; ARotation: TsTextRotation);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;

  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  Include(fmt.UsedFormattingFields, uffTextRotation);
  fmt.TextRotation := ARotation;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);

  ChangedFont(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Directly modifies the used formatting fields of a cell.
  Only formatting corresponding to items included in this set is executed.

  @param  ARow            The row of the cell
  @param  ACol            The column of the cell
  @param  AUsedFormatting set of the used formatting fields
  @return Pointer to the (existing or created) cell

  @see    TsUsedFormattingFields
  @see    TCell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteUsedFormatting(ARow, ACol: Cardinal;
  AUsedFormatting: TsUsedFormattingFields): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteUsedFormatting(Result, AUsedFormatting);
end;

{@@ ----------------------------------------------------------------------------
  Directly modifies the used formatting fields of an existing cell.
  Only formatting corresponding to items included in this set is executed.

  @param  ACell           Pointer to the cell to be modified
  @param  AUsedFormatting set of the used formatting fields

  @see    TsUsedFormattingFields
  @see    TCell
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteUsedFormatting(ACell: PCell;
  AUsedFormatting: TsUsedFormattingFields);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := FWorkbook.GetCellFormat(ACell^.FormatIndex);
  fmt.UsedFormattingFields := AUsedFormatting;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Defines a background pattern for a cell

  @param  ARow              Row index of the cell
  @param  ACol              Column index of the cell
  @param  AFillStyle        Fill style to be used - see TsFillStyle
  @param  APatternColor     RGB value of the pattern color
  @param  ABackgroundColor  RGB value of the background color
  @return Pointer to cell

  @NOTE Is replaced by uniform fill if WriteBackgroundColor is called later.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBackground(ARow, ACol: Cardinal; AStyle: TsFillStyle;
  APatternColor, ABackgroundColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBackground(Result, AStyle, APatternColor, ABackgroundColor);
end;

{@@ ----------------------------------------------------------------------------
  Defines a background pattern for a cell

  @param  ACell             Pointer to the cell
  @param  AStyle            Fill style ("pattern") to be used - see TsFillStyle
  @param  APatternColor     RGB value of the pattern color
  @param  ABackgroundColor  RGB value of the background color

  @NOTE Is replaced by uniform fill if WriteBackgroundColor is called later.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBackground(ACell: PCell; AStyle: TsFillStyle;
  APatternColor: TsColor = scTransparent; ABackgroundColor: TsColor = scTransparent);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    if (AStyle = fsNoFill) or
       ((APatternColor = scTransparent) and (ABackgroundColor = scTransparent))
    then
      Exclude(fmt.UsedFormattingFields, uffBackground)
    else
    begin
      Include(fmt.UsedFormattingFields, uffBackground);
      fmt.Background.Style := AStyle;
      fmt.Background.FgColor := APatternColor;
      if (AStyle = fsSolidFill) and (ABackgroundColor = scTransparent) then
        fmt.Background.BgColor := APatternColor
      else
        fmt.Background.BgColor := ABackgroundColor;
    end;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets a uniform background color of a cell.

  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  AColor     RGB value of the new background color.
                     Use the value "scTransparent" to clear an existing
                     background color.
  @return Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBackgroundColor(ARow, ACol: Cardinal;
  AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBackgroundColor(Result, AColor);
end;

{@@ ----------------------------------------------------------------------------
  Sets a uniform background color of a cell.

  @param  ACell      Pointer to cell
  @param  AColor     RGB value of the new background color.
                     Use the value "scTransparent" to clear an existing
                     background color.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBackgroundColor(ACell: PCell; AColor: TsColor);
begin
  if ACell <> nil then begin
    if AColor = scTransparent then
      WriteBackground(ACell, fsNoFill)
    else
      WriteBackground(ACell, fsSolidFill, AColor, AColor);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown!

  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  AColor     RGB value of the new border color
  @return Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderColor(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderColor(Result, ABorder, AColor);
end;

{@@ ----------------------------------------------------------------------------
  Sets the color of a cell border line.
  Note: the border must be included in Borders set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  AColor     RGB value of the new border color
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderColor(ACell: PCell; ABorder: TsCellBorder;
  AColor: TsColor);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder].Color := AColor;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  ALineStyle Identifier of the new line style to be applied.
  @return Pointer to cell

  @see    TsLineStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderLineStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderLineStyle(Result, ABorder, ALineStyle);
end;

{@@ ----------------------------------------------------------------------------
  Sets the linestyle of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Indicates to which border (left/top etc) this color is
                     to be applied
  @param  ALineStyle Identifier of the new line style to be applied.

  @see    TsLineStyle
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderLineStyle(ACell: PCell;
  ABorder: TsCellBorder; ALineStyle: TsLineStyle);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder].LineStyle := ALineStyle;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Shows the cell borders included in the set ABorders. No border lines are drawn
  for those not included.

  The borders are drawn using the "BorderStyles" assigned to the cell.

  @param  ARow      Row index of the cell
  @param  ACol      Column index of the cell
  @param  ABorders  Set with elements to identify the border(s) to will be shown
  @return Pointer to cell
  @see    TsCellBorder
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorders(ARow, ACol: Cardinal; ABorders: TsCellBorders): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorders(Result, ABorders);
end;

{@@ ----------------------------------------------------------------------------
  Shows the cell borders included in the set ABorders. No border lines are drawn
  for those not included.

  The borders are drawn using the "BorderStyles" assigned to the cell.

  @param  ACell     Pointer to cell
  @param  ABorders  Set with elements to identify the border(s) to will be shown
  @see    TsCellBorder
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorders(ACell: PCell; ABorders: TsCellBorders);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    if ABorders = [] then
      Exclude(fmt.UsedFormattingFields, uffBorder)
    else
      Include(fmt.UsedFormattingFields, uffBorder);
    fmt.Border := ABorders;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ARow       Row index of the cell considered
  @param  ACol       Column index of the cell considered
  @param  ABorder    Identifies the border to be modified (left/top/right/bottom)
  @param  AStyle     record of parameters controlling how the border line is drawn
                     (line style, line color)
  @result Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; AStyle: TsCellBorderStyle): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyle(Result, ABorder, AStyle);
end;

{@@ ----------------------------------------------------------------------------
  Sets the style of a cell border, i.e. line style and line color.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Identifies the border to be modified (left/top/right/bottom)
  @param  AStyle     record of parameters controlling how the border line is drawn
                     (line style, line color)
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
  AStyle: TsCellBorderStyle);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder] := AStyle;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ARow       Row index of the considered cell
  @param  ACol       Column index of the considered cell
  @param  ABorder    Identifier of the border to be modified
  @param  ALineStyle Identifier for the new line style of the border
  @param  AColor     RGB value of the border line color
  @return Pointer to cell

  @see WriteBorderStyles
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderStyle(ARow, ACol: Cardinal;
  ABorder: TsCellBorder; ALineStyle: TsLinestyle; AColor: TsColor): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyle(Result, ABorder, ALineStyle, AColor);
end;

{@@ ----------------------------------------------------------------------------
  Sets line style and line color of a cell border.
  Note: the border must be included in the "Borders" set in order to be shown!

  @param  ACell      Pointer to cell
  @param  ABorder    Identifier of the border to be modified
  @param  ALineStyle Identifier for the new line style of the border
  @param  AColor     RGB value of the color of the border line

  @see WriteBorderStyles
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderStyle(ACell: PCell; ABorder: TsCellBorder;
  ALineStyle: TsLinestyle; AColor: TsColor);
var
  fmt: TsCellFormat;
begin
  if ACell <> nil then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    fmt.BorderStyles[ABorder].LineStyle := ALineStyle;
    fmt.BorderStyles[ABorder].Color := AColor;
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown!

  @param  ARow    Row index of the considered cell
  @param  ACol    Column index of the considered cell
  @param  AStyles Array of CellBorderStyles for each cell border.
  @return Pointer to cell

  @see WriteBorderStyle
-------------------------------------------------------------------------------}
function TsWorksheet.WriteBorderStyles(ARow, ACol: Cardinal;
  const AStyles: TsCellBorderStyles): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteBorderStyles(Result, AStyles);
end;

{@@ ----------------------------------------------------------------------------
  Sets the style of all cell border of a cell, i.e. line style and line color.
  Note: Only those borders included in the "Borders" set are shown!

  @param  ACell   Pointer to cell
  @param  ACol    Column index of the considered cell
  @param  AStyles Array of CellBorderStyles for each cell border.

  @see WriteBorderStyle
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteBorderStyles(ACell: PCell;
  const AStyles: TsCellBorderStyles);
var
  b: TsCellBorder;
  fmt: TsCellFormat;
begin
  if Assigned(ACell) then begin
    fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
    for b in TsCellBorder do fmt.BorderStyles[b] := AStyles[b];
    ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Assigns a complete cell format record to a cell

  @param ACell        Pointer to the cell to be modified
  @param ACellFormat  Cell format record to be used by the cell

  @see TsCellFormat
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteCellFormat(ACell: PCell;
  const ACellFormat: TsCellFormat);
begin
  if Assigned(ACell) then begin
    ACell^.FormatIndex := Workbook.AddCellFormat(ACellFormat);
    ChangedCell(ACell^.Row, ACell^.Col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Defines the horizontal alignment of text in a cell.

  @param ARow    Row index of the cell considered
  @param ACol    Column index of the cell considered
  @param AValue  Parameter for horizontal text alignment
                 (haDefault, vaLeft, haCenter, haRight)
                 By default, texts are left-aligned, numbers and dates are
                 right-aligned.
  @return Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteHorAlignment(ARow, ACol: Cardinal; AValue: TsHorAlignment): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteHorAlignment(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Defines the horizontal alignment of text in a cell.

  @param ACell   Pointer to the cell considered
  @param AValue  Parameter for horizontal text alignment
                 (haDefault, vaLeft, haCenter, haRight)
                 By default, texts are left-aligned, numbers and dates are
                 right-aligned.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteHorAlignment(ACell: PCell; AValue: TsHorAlignment);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue = haDefault then
    Exclude(fmt.UsedFormattingFields, uffHorAlign)
  else
    Include(fmt.UsedFormattingFields, uffHorAlign);
  fmt.HorAlignment := AValue;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Defines the vertical alignment of text in a cell.

  @param ARow    Row index of the cell considered
  @param ACol    Column index of the cell considered
  @param AValue  Parameter for vertical text alignment
                 (vaDefault, vaTop, vaCenter, vaBottom)
                 By default, texts are bottom-aligned.
  @return Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteVertAlignment(ARow, ACol: Cardinal;
  AValue: TsVertAlignment): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteVertAlignment(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Defines the vertical alignment of text in a cell.

  @param ACell   Poiner to the cell considered
  @param AValue  Parameter for vertical text alignment
                 (vaDefault, vaTop, vaCenter, vaBottom)
                 By default, texts are bottom-aligned.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteVertAlignment(ACell: PCell; AValue: TsVertAlignment);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue = vaDefault then
    Exclude(fmt.UsedFormattingFields, uffVertAlign)
  else
    Include(fmt.UsedFormattingFields, uffVertAlign);
  fmt.VertAlignment := AValue;
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

{@@ ----------------------------------------------------------------------------
  Enables or disables the word-wrapping feature for a cell.

  @param  ARow    Row index of the cell considered
  @param  ACol    Column index of the cell considered
  @param  AValue  true = word-wrapping enabled, false = disabled.
  @return Pointer to cell
-------------------------------------------------------------------------------}
function TsWorksheet.WriteWordwrap(ARow, ACol: Cardinal; AValue: boolean): PCell;
begin
  Result := GetCell(ARow, ACol);
  WriteWordWrap(Result, AValue);
end;

{@@ ----------------------------------------------------------------------------
  Enables or disables the word-wrapping feature for a cell.

  @param ACel    Pointer to the cell considered
  @param AValue  true = word-wrapping enabled, false = disabled.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteWordwrap(ACell: PCell; AValue: boolean);
var
  fmt: TsCellFormat;
begin
  if ACell = nil then
    exit;
  fmt := Workbook.GetCellFormat(ACell^.FormatIndex);
  if AValue then
    Include(fmt.UsedFormattingFields, uffWordwrap)
  else
    Exclude(fmt.UsedFormattingFields, uffWordwrap);
  ACell^.FormatIndex := Workbook.AddCellFormat(fmt);
  ChangedCell(ACell^.Row, ACell^.Col);
end;

function TsWorksheet.GetFormatSettings: TFormatSettings;
begin
  Result := FWorkbook.FormatSettings;
end;

{@@ ----------------------------------------------------------------------------
  Calculates the optimum height of a given row. Depends on the font size
  of the individual cells in the row.

  @param  ARow   Index of the row to be considered
  @return Row height in line count of the default font.
-------------------------------------------------------------------------------}
function TsWorksheet.CalcAutoRowHeight(ARow: Cardinal): Single;
var
  cell: PCell;
  h0: Single;
begin
  Result := 0;
  h0 := Workbook.GetDefaultFontSize;
  for cell in Cells.GetRowEnumerator(ARow) do
    Result := Max(Result, ReadCellFont(cell).Size / h0);
end;

{@@ ----------------------------------------------------------------------------
 Checks if a row record exists for the given row index and returns a pointer
 to the row record, or nil if not found

 @param  ARow   Index of the row looked for
 @return        Pointer to the row record with this row index, or nil if not
                found
-------------------------------------------------------------------------------}
function TsWorksheet.FindRow(ARow: Cardinal): PRow;
var
  LElement: TRow;
  AVLNode: TAVGLVLTreeNode;
begin
  Result := nil;
  LElement.Row := ARow;
  AVLNode := FRows.Find(@LElement);
  if Assigned(AVLNode) then
    result := PRow(AVLNode.Data);
end;

{@@ ----------------------------------------------------------------------------
 Checks if a column record exists for the given column index and returns a
 pointer to the TCol record, or nil if not found

 @param  ACol   Index of the column looked for
 @return        Pointer to the column record with this column index, or
                nil if not found
-------------------------------------------------------------------------------}
function TsWorksheet.FindCol(ACol: Cardinal): PCol;
var
  LElement: TCol;
  AVLNode: TAVGLVLTreeNode;
begin
  Result := nil;
  LElement.Col := ACol;
  AVLNode := FCols.Find(@LElement);
  if Assigned(AVLNode) then
    result := PCol(AVLNode.Data);
end;

{@@ ----------------------------------------------------------------------------
 Checks if a row record exists for the given row index and creates it if not
 found.

 @param  ARow   Index of the row looked for
 @return        Pointer to the row record with this row index. It can safely be
                assumed that this row record exists.
-------------------------------------------------------------------------------}
function TsWorksheet.GetRow(ARow: Cardinal): PRow;
begin
  Result := FindRow(ARow);
  if (Result = nil) then begin
    Result := GetMem(SizeOf(TRow));
    FillChar(Result^, SizeOf(TRow), #0);
    Result^.Row := ARow;
    FRows.Add(Result);
    if FLastRowIndex = 0 then
      FLastRowIndex := GetLastRowIndex(true)
    else
      FLastRowIndex := Max(FLastRowIndex, ARow);
  end;
end;

{@@ ----------------------------------------------------------------------------
 Checks if a column record exists for the given column index and creates it
 if not found.

 @param  ACol   Index of the column looked for
 @return        Pointer to the TCol record with this column index. It can
                safely be assumed that this column record exists.
-------------------------------------------------------------------------------}
function TsWorksheet.GetCol(ACol: Cardinal): PCol;
begin
  Result := FindCol(ACol);
  if (Result = nil) then begin
    Result := GetMem(SizeOf(TCol));
    FillChar(Result^, SizeOf(TCol), #0);
    Result^.Col := ACol;
    FCols.Add(Result);
    if FFirstColIndex = 0
      then FFirstColIndex := GetFirstColIndex(true)
      else FFirstColIndex := Min(FFirstColIndex, ACol);
    if FLastColIndex = 0
      then FLastColIndex := GetLastColIndex(true)
      else FLastColIndex := Max(FLastColIndex, ACol);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Counts how many cells exist in the given column. Blank cells do contribute
  to the sum, as well as formatted cells.

  @param  ACol  Index of the column considered
  @return Count of cells with value or format in this column
-------------------------------------------------------------------------------}
function TsWorksheet.GetCellCountInCol(ACol: Cardinal): Cardinal;
var
  cell: PCell;
  r: Cardinal;
  row: PRow;
begin
  Result := 0;
  for r := GetFirstRowIndex to GetLastRowIndex do begin
    cell := FindCell(r, ACol);
    if cell <> nil then
      inc(Result)
    else begin
      row := FindRow(r);
      if row <> nil then inc(Result);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Counts how many cells exist in the given row. Blank cells do contribute
  to the sum, as well as formatted cell.s

  @param  ARow  Index of the row considered
  @return Count of cells with value or format in this row
-------------------------------------------------------------------------------}
function TsWorksheet.GetCellCountInRow(ARow: Cardinal): Cardinal;
var
  cell: PCell;
  c: Cardinal;
  col: PCol;
begin
  Result := 0;
  for c := 0 to GetLastColIndex do begin
    cell := FindCell(ARow, c);
    if cell <> nil then
      inc(Result)
    else begin
      col := FindCol(c);
      if col <> nil then inc(Result);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the width of the given column. If there is no column record then
  the default column width is returned.

  @param  ACol  Index of the column considered
  @return Width of the column (in count of "0" characters of the default font)
-------------------------------------------------------------------------------}
function TsWorksheet.GetColWidth(ACol: Cardinal): Single;
var
  col: PCol;
begin
  col := FindCol(ACol);
  if col <> nil then
    Result := col^.Width
  else
    Result := FDefaultColWidth;
end;

{@@ ----------------------------------------------------------------------------
  Returns the height of the given row. If there is no row record then the
  default row height is returned

  @param  ARow  Index of the row considered
  @return Height of the row (in line count of the default font).
-------------------------------------------------------------------------------}
function TsWorksheet.GetRowHeight(ARow: Cardinal): Single;
var
  row: PRow;
begin
  row := FindRow(ARow);
  if row <> nil then
    Result := row^.Height
  else
    //Result := CalcAutoRowHeight(ARow);
    Result := FDefaultRowHeight;
end;

{@@ ----------------------------------------------------------------------------
  Deletes the column at the index specified. Cells with greader column indexes
  are moved one column to the left. Merged cell blocks and cell references in
  formulas are considered as well.

  @param   ACol   Index of the column to be deleted
-------------------------------------------------------------------------------}
procedure TsWorksheet.DeleteCol(ACol: Cardinal);
var
  col: PCol;
  i: Integer;
  r: Cardinal;
  cell: PCell;
  firstRow, lastRow: Cardinal;
begin
  lastRow := GetLastOccupiedRowIndex;
  firstRow := GetFirstRowIndex;

  // Fix merged cells
  FMergedCells.DeleteRowOrCol(ACol, false);

  // Fix comments
  FComments.DeleteRowOrCol(ACol, false);

  // Fix hyperlinks
  FHyperlinks.DeleteRowOrCol(ACol, false);

  // Delete cells
  for r := lastRow downto firstRow do
    RemoveAndFreeCell(r, ACol);

  // Update column index of cell records
  for cell in FCells do
    DeleteColCallback(cell, {%H-}pointer(PtrInt(ACol)));

  // Update column index of col records
  for i:=FCols.Count-1 downto 0 do begin
    col := PCol(FCols.Items[i]);
    if col^.Col > ACol then
      dec(col^.Col)
    else
      break;
  end;

  // Update first and last column index
  UpDateCaches;

  ChangedCell(0, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Deletes the row at the index specified. Cells with greader row indexes are
  moved one row up. Merged cell blocks and cell references in formulas
  are considered as well.

  @param   ARow   Index of the row to be deleted
-------------------------------------------------------------------------------}
procedure TsWorksheet.DeleteRow(ARow: Cardinal);
var
  row: PRow;
  i: Integer;
  c: Cardinal;
  firstCol, lastCol: Cardinal;
  cell: PCell;
begin
  firstCol := GetFirstColIndex;
  lastCol := GetLastOccupiedColIndex;

  // Fix merged cells
  FMergedCells.DeleteRowOrCol(ARow, true);

  // Fix comments
  FComments.DeleteRowOrCol(ARow, true);

  // Fix hyperlinks
  FHyperlinks.DeleteRowOrCol(ARow, true);

  // Delete cells
  for c := lastCol downto firstCol do
    RemoveAndFreeCell(ARow, c);

  // Update row index of cell records
  for cell in FCells do
    DeleteRowCallback(cell, {%H-}pointer(PtrInt(ARow)));

  // Update row index of row records
  for i:=FRows.Count-1 downto 0 do
  begin
    row := PRow(FRows.Items[i]);
    if row^.Row > ARow then
      dec(row^.Row)
    else
      break;
  end;

  // Update first and last row index
  UpdateCaches;

  ChangedCell(ARow, 0);
end;

{@@ ----------------------------------------------------------------------------
  Inserts a column BEFORE the index specified. Cells with greater column indexes
  are moved one column to the right. Merged cell blocks and cell references in
  formulas are considered as well.

  @param   ACol   Index of the column before which a new column is inserted.
-------------------------------------------------------------------------------}
procedure TsWorksheet.InsertCol(ACol: Cardinal);
var
  col: PCol;
  i: Integer;
  cell: PCell;
  rng: PsCellRange;
begin
  // Update column index of comments
  FComments.InsertRowOrCol(ACol, false);

  // Update column index of hyperlinks
  FHyperlinks.InsertRowOrCol(ACol, false);

  // Update column index of cell records
  for cell in FCells do
    InsertColCallback(cell, {%H-}pointer(PtrInt(ACol)));

  // Update column index of column records
  for i:=0 to FCols.Count-1 do begin
    col := PCol(FCols.Items[i]);
    if col^.Col >= ACol then inc(col^.Col);
  end;

  // Update first and last column index
  UpdateCaches;

  // Fix merged cells
  for rng in FMergedCells do
  begin
    // The new column is at the LEFT of the merged block
    // --> Shift entire range to the right by 1 column
    if (ACol < rng^.Col1) then
    begin
      // The former first column is no longer merged --> un-tag its cells
      for cell in Cells.GetColEnumerator(rng^.Col1, rng^.Row1, rng^.Row2) do
        Exclude(cell^.Flags, cfMerged);

      // Shift merged block to the right
      // Don't call "MergeCells" here - this would add a new merged block
      // because of the new merge base! --> infinite loop!
      inc(rng^.Col1);
      inc(rng^.Col2);
      // The right column needs to be tagged
      for cell in Cells.GetColEnumerator(rng^.Col2, rng^.Row1, rng^.Row2) do
        Include(cell^.Flags, cfMerged);
    end else
    // The new column goes through this cell block --> Shift only the right
    // column of the range to the right by 1
    if (ACol >= rng^.Col1) and (ACol <= rng^.Col2) then
      MergeCells(rng^.Row1, rng^.Col1, rng^.Row2, rng^.Col2+1);
  end;

  ChangedCell(0, ACol);
end;

procedure TsWorksheet.InsertColCallback(data, arg: Pointer);
var
  cell: PCell;
  col: Cardinal;
  formula: TsRPNFormula;
  i: Integer;
begin
  col := LongInt({%H-}PtrInt(arg));
  cell := PCell(data);
  if cell = nil then   // This should not happen. Just to make sure...
    exit;

  // Update row index of moved cells
  if cell^.Col >= col then
    inc(cell^.Col);

  // Update formulas
  if HasFormula(cell) and (cell^.FormulaValue <> '' ) then
  begin
    // (1) create an rpn formula
    formula := BuildRPNFormula(cell);
    // (2) update cell addresses affected by the insertion of a column
    for i:=0 to Length(formula)-1 do
    begin
      case formula[i].ElementKind of
        fekCell, fekCellRef:
          if formula[i].Col >= col then inc(formula[i].Col);
        fekCellRange:
          begin
            if formula[i].Col >= col then inc(formula[i].Col);
            if formula[i].Col2 >= col then inc(formula[i].Col2);
          end;
      end;
    end;
    // (3) convert rpn formula back to string formula
    cell^.FormulaValue := ConvertRPNFormulaToStringFormula(formula);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inserts a row BEFORE the row specified. Cells with greater row indexes are
  moved one row down. Merged cell blocks and cell references in formulas are
  considered as well.

  @param   ARow   Index of the row before which a new row is inserted.
-------------------------------------------------------------------------------}
procedure TsWorksheet.InsertRow(ARow: Cardinal);
var
  row: PRow;
  i: Integer;
  cell: PCell;
  rng: PsCellRange;
begin
  // Update row index of cell comments
  FComments.InsertRowOrCol(ARow, true);

  // Update row index of cell hyperlinks
  FHyperlinks.InsertRowOrCol(ARow, true);

  // Update row index of cell records
  for cell in FCells do
    InsertRowCallback(cell, {%H-}pointer(PtrInt(ARow)));

  // Update row index of row records
  for i:=0 to FRows.Count-1 do begin
    row := PRow(FRows.Items[i]);
    if row^.Row >= ARow then inc(row^.Row);
  end;

  // Update first and last row index
  UpdateCaches;

  // Fix merged cells
  for rng in FMergedCells do
  begin
    // The new row is ABOVE the merged block --> Shift entire range down by 1 row
    if (ARow < rng^.Row1) then
    begin
      // The formerly first row is no longer merged --> un-tag its cells
      for cell in Cells.GetRowEnumerator(rng^.Row1, rng^.Col1, rng^.Col2) do
        Exclude(cell^.Flags, cfMerged);

      // Shift merged block down
      // (Don't call "MergeCells" here - this would add a new merged block
      // because of the new merge base! --> infinite loop!)
      inc(rng^.Row1);
      inc(rng^.Row2);
      // The last row needs to be tagged
      for cell in Cells.GetRowEnumerator(rng^.Row2, rng^.Col1, rng^.Col2) do
        Include(cell^.Flags, cfMerged);
    end else
    // The new row goes through this cell block --> Shift only the bottom row
    // of the range down by 1
    if (ARow >= rng^.Row1) and (ARow <= rng^.Row2) then
      MergeCells(rng^.Row1, rng^.Col1, rng^.Row2+1, rng^.Col2);
  end;

  ChangedCell(ARow, 0);
end;

procedure TsWorksheet.InsertRowCallback(data, arg: Pointer);
var
  cell: PCell;
  row: Cardinal;
  i: Integer;
  formula: TsRPNFormula;
begin
  row := LongInt({%H-}PtrInt(arg));
  cell := PCell(data);
  if cell = nil then   // This should not happen. Just to make sure...
    exit;

  // Update row index of moved cells
  if cell^.Row >= row then
    inc(cell^.Row);

  // Update formulas
  if HasFormula(cell) then
  begin
    // (1) create an rpn formula
    formula := BuildRPNFormula(cell);
    // (2) update cell addresses affected by the insertion of a column
    for i:=0 to Length(formula)-1 do begin
      case formula[i].ElementKind of
        fekCell, fekCellRef:
          if formula[i].Row >= row then inc(formula[i].Row);
        fekCellRange:
          begin
            if formula[i].Row >= row then inc(formula[i].Row);
            if formula[i].Row2 >= row then inc(formula[i].Row2);
          end;
      end;
    end;
    // (3) convert rpn formula back to string formula
    cell^.FormulaValue := ConvertRPNFormulaToStringFormula(formula);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes all row records from the worksheet and frees the occupied memory.
  Note: Cells are retained.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveAllRows;
var
  Node: Pointer;
  i: Integer;
begin
  for i := FRows.Count-1 downto 0 do begin
    Node := FRows.Items[i];
    FreeMem(Node, SizeOf(TRow));
  end;
  FRows.Clear;
end;

{@@ ----------------------------------------------------------------------------
  Removes all column records from the worksheet and frees the occupied memory.
  Note: Cells are retained.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveAllCols;
var
  Node: Pointer;
  i: Integer;
begin
  for i := FCols.Count-1 downto 0 do begin
    Node := FCols.Items[i];
    FreeMem(Node, SizeOf(TCol));
  end;
  FCols.Clear;
end;

{@@ ----------------------------------------------------------------------------
  Removes a specified column record from the worksheet and frees the occupied
  memory. This resets its column width to default.

  Note: Cells in that column are retained.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveCol(ACol: Cardinal);
var
  AVLNode: TAVGLVLTreeNode;
  lCol: TCol;
begin
  lCol.Col := ACol;
  AVLNode := FCols.Find(@lCol);
  if Assigned(AVLNode) then
  begin
    FreeMem(PCol(AVLNode.Data), SizeOf(TCol));
    FCols.Delete(AVLNode);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes a specified row record from the worksheet and frees the occupied memory.
  This resets the its row height to default.
  Note: Cells in that row are retained.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveRow(ARow: Cardinal);
var
  AVLNode: TAVGLVLTreeNode;
  lRow: TRow;
begin
  lRow.Row := ARow;
  AVLNode := FRows.Find(@lRow);
  if Assigned(AVLNode) then
  begin
    FreeMem(PRow(AVLNode.Data), SizeOf(TRow));
    FRows.Delete(AVLNode);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a row record for the row at a given index to the spreadsheet.
  Currently the row record contains only the row height (and the row index,
  of course).

  Creates a new row record if it does not yet exist.

  @param  ARow   Index of the row record which will be created or modified
  @param  AData  Data to be written.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteRowInfo(ARow: Cardinal; AData: TRow);
var
  AElement: PRow;
begin
  AElement := GetRow(ARow);
  AElement^.Height := AData.Height;
end;

{@@ ----------------------------------------------------------------------------
  Sets the row height for a given row. Creates a new row record if it
  does not yet exist.

  @param  ARow     Index of the row to be considered
  @param  AHeight  Row height to be assigned to the row. The row height is
                   expressed as the line count of the default font size.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteRowHeight(ARow: Cardinal; AHeight: Single);
var
  AElement: PRow;
begin
  AElement := GetRow(ARow);
  AElement^.Height := AHeight;
end;

{@@ ----------------------------------------------------------------------------
  Writes a column record for the column at a given index to the spreadsheet.
  Currently the column record contains only the column width (and the column
  index, of course).

  Creates a new column record if it does not yet exist.

  @param  ACol   Index of the column record which will be created or modified
  @param  AData  Data to be written (essentially column width).
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteColInfo(ACol: Cardinal; AData: TCol);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AData.Width;
end;

{@@ ----------------------------------------------------------------------------
  Sets the column width for a given column. Creates a new column record if it
  does not yet exist.

  @param  ACol     Index of the column to be considered
  @param  AWidth   Width to be assigned to the column. The column width is
                   expressed as the count of "0" characters of the default font.
-------------------------------------------------------------------------------}
procedure TsWorksheet.WriteColWidth(ACol: Cardinal; AWidth: Single);
var
  AElement: PCol;
begin
  AElement := GetCol(ACol);
  AElement^.Width := AWidth;
end;


{*******************************************************************************
*                              TsWorkbook                                      *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Helper method called before reading the workbook. Clears the error log.
-------------------------------------------------------------------------------}
procedure TsWorkbook.PrepareBeforeReading;
begin
  // Initializes fonts
  InitFonts;

  // Clear error log
  FLog.Clear;

  // Abort if virtual mode is active without an event handler
  if (boVirtualMode in FOptions) and not Assigned(FOnReadCellData) then
    raise Exception.Create('[TsWorkbook.PrepareBeforeReading] Event handler "OnReadCellData" required for virtual mode.');
end;

{@@ ----------------------------------------------------------------------------
  Helper method called before saving the workbook. Clears the error log, and
  calculates the formulas in all worksheets if workbook option soCalcBeforeSaving
  is set.
-------------------------------------------------------------------------------}
procedure TsWorkbook.PrepareBeforeSaving;
var
  sheet: TsWorksheet;
begin
  // Clear error log
  FLog.Clear;

  // Updates fist/last column/row index
  UpdateCaches;

  // Calculated formulas (if requested)
  if (boCalcBeforeSaving in FOptions) then
    for sheet in FWorksheets do
      sheet.CalcFormulas;

  // Abort if virtual mode is active without an event handler
  if (boVirtualMode in FOptions) and not Assigned(FOnWriteCellData) then
    raise Exception.Create('[TsWorkbook.PrepareBeforeWriting] Event handler "OnWriteCellData" required for virtual mode.');
end;

{@@ ----------------------------------------------------------------------------
  Recalculates rpn formulas in all worksheets
-------------------------------------------------------------------------------}
procedure TsWorkbook.Recalc;
var
  sheet: TsWorksheet;
begin
  for sheet in FWorksheets do
    sheet.CalcFormulas;
end;

{@@ ----------------------------------------------------------------------------
  Helper method for clearing the spreadsheet list.
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveWorksheetsCallback(data, arg: pointer);
begin
  Unused(arg);
  TsWorksheet(data).Free;
end;

{@@ ----------------------------------------------------------------------------
  Helper method to update internal caching variables
-------------------------------------------------------------------------------}
procedure TsWorkbook.UpdateCaches;
var
  sheet: TsWorksheet;
begin
  for sheet in FWorksheets do
    sheet.UpdateCaches;
end;

{@@ ----------------------------------------------------------------------------
  Constructor of the workbook class. Among others, it initializes the built-in
  fonts, defines the default font, and sets up the FormatSettings for
  localization of some number formats.
-------------------------------------------------------------------------------}
constructor TsWorkbook.Create;
var
  fmt: TsCellFormat;
begin
  inherited Create;
  FWorksheets := TFPList.Create;
  FLog := TStringList.Create;
  FFormat := sfExcel8;

  FormatSettings := UTF8FormatSettings;
  FormatSettings.ShortDateFormat := MakeShortDateFormat(FormatSettings.ShortDateFormat);
  FormatSettings.LongDateFormat := MakeLongDateFormat(FormatSettings.ShortDateFormat);

  FFontList := TFPList.Create;
  SetDefaultFont(DEFAULT_FONTNAME, DEFAULT_FONTSIZE);
  InitFonts;

  FNumFormatList := TsNumFormatList.Create(self, true);
  FCellFormatList := TsCellFormatList.Create(false);

  // Add default cell format
  InitFormatRecord(fmt);
  AddCellFormat(fmt);
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the workbook class
-------------------------------------------------------------------------------}
destructor TsWorkbook.Destroy;
begin
  RemoveAllWorksheets;
  RemoveAllFonts;

  FWorksheets.Free;
  FCellFormatList.Free;
  FNumFormatList.Free;
  FFontList.Free;

  FLog.Free;
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Helper method for determining the spreadsheet type. Read the first few bytes
  of a file and determines the spreadsheet type from the characteristic
  signature. Only implemented for xls files where several file types have the
  same extension
-------------------------------------------------------------------------------}
class function TsWorkbook.GetFormatFromFileHeader(const AFileName: TFileName;
  out SheetType: TsSpreadsheetFormat): Boolean;
const
  BIFF2_HEADER: array[0..3] of byte = (
    $09,$00, $04,$00);  // they are common to all BIFF2 files that I've seen
  BIFF58_HEADER: array[0..7] of byte = (
    $D0,$CF, $11,$E0, $A1,$B1, $1A,$E1);

  function ValidOLEStream(AStream: TStream; AName: String): Boolean;
  var
    fsOLE: TVirtualLayer_OLE;
  begin
    AStream.Position := 0;
    fsOLE := TVirtualLayer_OLE.Create(AStream);
    try
      fsOLE.Initialize;
      Result := fsOLE.FileExists('/'+AName);
    finally
      fsOLE.Free;
    end;
  end;

var
  buf: packed array[0..7] of byte = (0,0,0,0,0,0,0,0);
  stream: TStream;
  i: Integer;
  ok: Boolean;
begin
  Result := false;
  stream := TFileStream.Create(AFileName, fmOpenRead + fmShareDenyNone);
  try
    // Read first 8 bytes
    stream.ReadBuffer(buf, 8);

    // Check for Excel 2
    ok := true;
    for i:=0 to High(BIFF2_HEADER) do
      if buf[i] <> BIFF2_HEADER[i] then
      begin
        ok := false;
        break;
      end;
    if ok then
    begin
      SheetType := sfExcel2;
      Exit(True);
    end;

    // Check for Excel 5 or 8
    for i:=0 to High(BIFF58_HEADER) do
      if buf[i] <> BIFF58_HEADER[i] then
        exit;

    // Now we know that the file is a Microsoft compound document.

    // We check for Excel 5 in which the stream is named "Book"
    if ValidOLEStream(stream, 'Book') then begin
      SheetType := sfExcel5;
      exit(True);
    end;

    // Now we check for Excel 8 which names the stream "Workbook"
    if ValidOLEStream(stream, 'Workbook') then begin
      SheetType := sfExcel8;
      exit(True);
    end;

  finally
    stream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Convenience method which creates the correct reader object for a given
  spreadsheet format.

  @param  AFormat  File format which is assumed when reading a document into
                   to workbook. An exception is raised when the document has
                   a different format.

  @return An instance of a TsCustomSpreadReader descendent which is able to
          read the given file format.
-------------------------------------------------------------------------------}
function TsWorkbook.CreateSpreadReader(AFormat: TsSpreadsheetFormat): TsBasicSpreadReader;
var
  i: Integer;
begin
  Result := nil;

  for i := 0 to Length(GsSpreadFormats) - 1 do
    if GsSpreadFormats[i].Format = AFormat then
    begin
      Result := GsSpreadFormats[i].ReaderClass.Create(self);
      Break;
    end;

  if Result = nil then
    raise Exception.Create(rsUnsupportedReadFormat);
end;

{@@ ----------------------------------------------------------------------------
  Convenience method which creates the correct writer object for a given
  spreadsheet format.

  @param  AFormat  File format to be used for writing the workbook

  @return An instance of a TsCustomSpreadWriter descendent which is able to
          write the given file format.
-------------------------------------------------------------------------------}
function TsWorkbook.CreateSpreadWriter(AFormat: TsSpreadsheetFormat): TsBasicSpreadWriter;
var
  i: Integer;
begin
  Result := nil;

  for i := 0 to Length(GsSpreadFormats) - 1 do
    if GsSpreadFormats[i].Format = AFormat then
    begin
      Result := GsSpreadFormats[i].WriterClass.Create(self);
      Break;
    end;
    
  if Result = nil then
    raise Exception.Create(rsUnsupportedWriteFormat);
end;

{@@ ----------------------------------------------------------------------------
  Determines the maximum index of used columns and rows in all sheets of this
  workbook. Respects VirtualMode.
  Is needed to disable saving when limitations of the format is exceeded.
-------------------------------------------------------------------------------}
procedure TsWorkbook.GetLastRowColIndex(out ALastRow, ALastCol: Cardinal);
var
  i: Integer;
  sheet: TsWorksheet;
begin
  if (boVirtualMode in Options) then
  begin
    ALastRow := FVirtualRowCount - 1;
    ALastCol := FVirtualColCount - 1;
  end else
  begin
    ALastRow := 0;
    ALastCol := 0;
    for i:=0 to GetWorksheetCount-1 do
    begin
      sheet := GetWorksheetByIndex(i);
      ALastRow := Max(ALastRow, sheet.GetLastRowIndex);
      ALastCol := Max(ALastCol, sheet.GetLastColIndex);
    end;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Reads the document from a file. It is assumed to have the given file format.

  @param  AFileName  Name of the file to be read
  @param  AFormat    File format assumed
-------------------------------------------------------------------------------}
procedure TsWorkbook.ReadFromFile(AFileName: string;
  AFormat: TsSpreadsheetFormat);
var
  AReader: TsBasicSpreadReader;
  ok: Boolean;
begin
  if not FileExists(AFileName) then
    raise Exception.CreateFmt(rsFileNotFound, [AFileName]);

  AReader := CreateSpreadReader(AFormat);
  try
    FFileName := AFileName;
    PrepareBeforeReading;
    ok := false;
    FReadWriteFlag := rwfRead;
    inc(FLockCount);          // This locks various notifications from being sent
    try
      AReader.ReadFromFile(AFileName);
      ok := true;
      UpdateCaches;
      if (boAutoCalc in Options) then
        Recalc;
      FFormat := AFormat;
    finally
      FReadWriteFlag := rwfNormal;
      dec(FLockCount);
      if ok and Assigned(FOnOpenWorkbook) then   // ok is true if file has been read successfully
        FOnOpenWorkbook(self);   // send common notification
    end;
  finally
    AReader.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the document from a file. This method will try to guess the format from
  the extension. In the case of the ambiguous xls extension, it will simply
  assume that it is BIFF8. Note that it could be BIFF2 or 5 as well.
-------------------------------------------------------------------------------}
procedure TsWorkbook.ReadFromFile(AFileName: string); overload;
var
  SheetType: TsSpreadsheetFormat;
  valid: Boolean;
  lException: Exception = nil;
begin
  if not FileExists(AFileName) then
    raise Exception.CreateFmt(rsFileNotFound, [AFileName]);

  // .xls files can contain several formats. We look into the header first.
  if Lowercase(ExtractFileExt(AFileName))=STR_EXCEL_EXTENSION then
  begin
    valid := GetFormatFromFileHeader(AFileName, SheetType);
    // It is possible that valid xls files are not detected correctly. Therefore,
    // we open them explicitly by trial and error - see below.
    if not valid then
      SheetType := sfExcel8;
    valid := true;
  end else
    valid := GetFormatFromFileName(AFileName, SheetType);

  if valid then
  begin
    if SheetType = sfExcel8 then
    begin
      // Here is the trial-and-error loop checking for the various biff formats.
      while True do
      begin
        try
          ReadFromFile(AFileName, SheetType);
          valid := True;
        except
          on E: Exception do
          begin
            if SheetType = sfExcel8 then lException := E;
            valid := False
          end;
        end;
        if valid or (SheetType = sfExcel2) then Break;
        SheetType := Pred(SheetType);
      end;

      // A failed attempt to read a file should bring an exception, so re-raise
      // the exception if necessary. We re-raise the exception brought by Excel 8,
      // since this is the most common format
      if (not valid) and (lException <> nil) then raise lException;
    end
    else
      ReadFromFile(AFileName, SheetType);
  end else
    raise Exception.CreateFmt(rsNoValidSpreadsheetFile, [AFileName]);
end;

{@@ ----------------------------------------------------------------------------
  Reads the document from a file, but ignores the extension.
-------------------------------------------------------------------------------}
procedure TsWorkbook.ReadFromFileIgnoringExtension(AFileName: string);
var
  SheetType: TsSpreadsheetFormat;
  lException: Exception;
begin
  lException := pointer(1);  // Must not be nil initially
  SheetType := sfExcel8;
  while (SheetType in [sfExcel2..sfExcel8, sfOpenDocument, sfOOXML]) and (lException <> nil) do
  begin
    try
      Dec(SheetType);
      ReadFromFile(AFileName, SheetType);
      lException := nil;
    except
      on E: Exception do { do nothing } ;
    end;
    if lException = nil then Break;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the document from a seekable stream.

  @param  AStream  Stream being read
  @param  AFormat  File format assumed.
-------------------------------------------------------------------------------}
procedure TsWorkbook.ReadFromStream(AStream: TStream;
  AFormat: TsSpreadsheetFormat);
var
  AReader: TsBasicSpreadReader;
  ok: Boolean;
begin
  AReader := CreateSpreadReader(AFormat);
  try
    PrepareBeforeReading;
    inc(FLockCount);
    try
      ok := false;
      AReader.ReadFromStream(AStream);
      ok := true;
    finally
      dec(FLockCount);
      if ok and Assigned(FOnOpenWorkbook) then
        FOnOpenWorkbook(self);
    end;
    UpdateCaches;
    if (boAutoCalc in Options) then
      Recalc;
  finally
    AReader.Free;
  end;
end;

procedure TsWorkbook.SetVirtualColCount(AValue: Cardinal);
begin
  if FReadWriteFlag = rwfWrite then exit;
  FVirtualColCount := AValue;
end;

procedure TsWorkbook.SetVirtualRowCount(AValue: Cardinal);
begin
  if FReadWriteFlag = rwfWrite then exit;
  FVirtualRowCount := AValue;
end;

{@@ ----------------------------------------------------------------------------
  Writes the document to a file. If the file doesn't exist, it will be created.

  @param  AFileName  Name of the file to be written
  @param  AFormat    The file will be written in this file format
  @param  AOverwriteExisting  If the file is already existing it will be
                     overwritten in case of AOverwriteExisting = true.
                     If false an exception will be raised.
-------------------------------------------------------------------------------}
procedure TsWorkbook.WriteToFile(const AFileName: string;
 const AFormat: TsSpreadsheetFormat; const AOverwriteExisting: Boolean = False);
var
  AWriter: TsBasicSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);
  try
    FFileName := AFileName;
    FFormat := AFormat;
    PrepareBeforeSaving;
    AWriter.CheckLimitations;
    FReadWriteFlag := rwfWrite;
    AWriter.WriteToFile(AFileName, AOverwriteExisting);
  finally
    FReadWriteFlag := rwfNormal;
    AWriter.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the document to file based on the extension.
  If this was an earlier sfExcel type file, it will be upgraded to sfExcel8.

  @param  AFileName  Name of the destination file
  @param  AOverwriteExisting  If the file already exists it will be overwritten
                     of AOverwriteExisting is true. In case of false, an
                     exception will be raised.
-------------------------------------------------------------------------------}
procedure TsWorkbook.WriteToFile(const AFileName: String;
  const AOverwriteExisting: Boolean);
var
  SheetType: TsSpreadsheetFormat;
  valid: Boolean;
begin
  valid := GetFormatFromFileName(AFileName, SheetType);
  if valid then
    WriteToFile(AFileName, SheetType, AOverwriteExisting)
  else
    raise Exception.Create(Format(rsInvalidExtension, [
      ExtractFileExt(AFileName)
    ]));
end;

{@@ ----------------------------------------------------------------------------
  Writes the document to a stream

  @param  AStream   Instance of the stream being written to
  @param  AFormat   File format to be written.
-------------------------------------------------------------------------------}
procedure TsWorkbook.WriteToStream(AStream: TStream; AFormat: TsSpreadsheetFormat);
var
  AWriter: TsBasicSpreadWriter;
begin
  AWriter := CreateSpreadWriter(AFormat);
  try
    PrepareBeforeSaving;
    AWriter.CheckLimitations;
    FReadWriteFlag := rwfWrite;
    AWriter.WriteToStream(AStream);
  finally
    FReadWriteFlag := rwfNormal;
    AWriter.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Adds a new worksheet to the workbook.
  It is put to the end of the worksheet list.

  @param  AName                The name of the new worksheet
  @param  ReplaceDupliateName  If true and the sheet name already exists then
                               a number is added to the sheet name to make it
                               unique.
  @return The instance of the newly created worksheet
  @see    TsWorksheet
-------------------------------------------------------------------------------}
function TsWorkbook.AddWorksheet(AName: string;
  ReplaceDuplicateName: Boolean = false): TsWorksheet;
begin
  // Check worksheet name
  if not ValidWorksheetName(AName, ReplaceDuplicateName) then
    raise Exception.CreateFmt(rsInvalidWorksheetName, [AName]);

  // Create worksheet...
  Result := TsWorksheet.Create;

  // Add it to the internal worksheet list
  FWorksheets.Add(Pointer(Result));

  // Remember the workbook to which it belongs (This must occur before
  // setting the workbook name because the workbook is needed there).
  Result.FWorkbook := Self;
  Result.FActiveCellRow := 0;
  Result.FActiveCellCol := 0;

  // Set the name of the new worksheet.
  // For this we turn off notification of listeners. This is not necessary here
  // because it will be repeated at end when OnAddWorksheet is executed below.
  inc(FLockCount);
  try
    Result.Name := AName;
  finally
    dec(FLockCount);
  end;

  // Send notification for new worksheet to listeners. They get the worksheet
  // name here as well.
  if (FLockCount = 0) and Assigned(FOnAddWorksheet) then
    FOnAddWorksheet(self, Result);
end;

{@@ ----------------------------------------------------------------------------
  Copies a worksheet (even from an external workbook) and adds it to the
  current workbook

  @param  AWorksheet            Worksheet to be copied. Can be in a different
                                workbook.
  @param  ReplaceDuplicateName  The copied worksheet gets the name of the original.
                                If ReplaceDuplicateName is true and this sheet
                                name already exists then a number is added to
                                the sheet name to make it unique.
  @return The instance of the newly created worksheet
  @see    TsWorksheet
-------------------------------------------------------------------------------}
function TsWorkbook.CopyWorksheetFrom(AWorksheet: TsWorksheet;
  ReplaceDuplicateName: boolean = false): TsWorksheet;
var
  r, c: Cardinal;
  cell: PCell;
begin
  Result := nil;
  if (AWorksheet = nil) then
    exit;

  Result := AddWorksheet(AWorksheet.Name, ReplaceDuplicateName);
  inc(FLockCount);
  try
    for cell in AWorksheet.Cells do
    begin
      r := cell^.Row;
      c := cell^.Col;
      Result.CopyCell(r, c, r, c, AWorksheet);
    end;
  finally
    dec(FLockCount);
  end;

  Result.ChangedCell(r, c);
end;

{@@ ----------------------------------------------------------------------------
  Quick helper routine which returns the first worksheet

  @return A TsWorksheet instance if at least one is present.
          nil otherwise.

  @see    TsWorkbook.GetWorksheetByIndex
  @see    TsWorkbook.GetWorksheetByName
  @see    TsWorksheet
-------------------------------------------------------------------------------}
function TsWorkbook.GetFirstWorksheet: TsWorksheet;
begin
  Result := TsWorksheet(FWorksheets.First);
end;

{@@ ----------------------------------------------------------------------------
  Gets the worksheet with a given index

  The index is zero-based, so the first worksheet
  added has index 0, the second 1, etc.

  @param  AIndex    The index of the worksheet (0-based)

  @return A TsWorksheet instance if one is present at that index.
          nil otherwise.

  @see    TsWorkbook.GetFirstWorksheet
  @see    TsWorkbook.GetWorksheetByName
  @see    TsWorksheet
-------------------------------------------------------------------------------}
function TsWorkbook.GetWorksheetByIndex(AIndex: Integer): TsWorksheet;
begin
  if (integer(AIndex) < FWorksheets.Count) and (integer(AIndex)>=0) then
    Result := TsWorksheet(FWorksheets.Items[AIndex])
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Gets the worksheet with a given worksheet name

  @param  AName    The name of the worksheet
  @return A TsWorksheet instance if one is found with that name,
          nil otherwise. Case is ignored.

  @see    TsWorkbook.GetFirstWorksheet
  @see    TsWorkbook.GetWorksheetByIndex
  @see    TsWorksheet
-------------------------------------------------------------------------------}
function TsWorkbook.GetWorksheetByName(AName: String): TsWorksheet;
var
  i:integer;
begin
  Result := nil;
  for i:=0 to FWorksheets.Count-1 do
  begin
    if UTF8CompareText(TsWorkSheet(FWorkSheets.Items[i]).Name, AName) = 0 then
    begin
      Result := TsWorksheet(FWorksheets.Items[i]);
      exit;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  The number of worksheets on the workbook

  @see    TsWorksheet
-------------------------------------------------------------------------------}
function TsWorkbook.GetWorksheetCount: Integer;
begin
  Result := FWorksheets.Count;
end;

{@@ ----------------------------------------------------------------------------
  Returns the index of a worksheet in the worksheet list
-------------------------------------------------------------------------------}
function TsWorkbook.GetWorksheetIndex(AWorksheet: TsWorksheet): Integer;
begin
  Result := FWorksheets.IndexOf(AWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Clears the list of Worksheets and releases their memory.

  NOTE: This procedure conflicts with the WorkbookLink mechanism which requires
  at least 1 worksheet per workbook!
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveAllWorksheets;
begin
  FWorksheets.ForEachCall(RemoveWorksheetsCallback, nil);
end;

{@@ ----------------------------------------------------------------------------
  Removes the specified worksheet: Removes the sheet from the internal sheet
  list, generates an event OnRemoveWorksheet, and releases all memory.
  The event handler specifies the index of the deleted worksheet; the worksheet
  itself does no longer exist.
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveWorksheet(AWorksheet: TsWorksheet);
var
  i: Integer;
begin
  if GetWorksheetCount > 1 then     // There must be at least 1 worksheet!
  begin
    i := GetWorksheetIndex(AWorksheet);
    if (i <> -1) and (AWorksheet <> nil) then
    begin
      if Assigned(FOnRemovingWorksheet) then
        FOnRemovingWorksheet(self, AWorksheet);
      FWorksheets.Delete(i);
      AWorksheet.Free;
      if Assigned(FOnRemoveWorksheet) then
        FOnRemoveWorksheet(self, i);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Makes the specified worksheet "active". Only needed for visual controls.
  The active worksheet is displayed in a TsWorksheetGrid and in the selected
  tab of a TsWorkbookTabControl.
-------------------------------------------------------------------------------}
procedure TsWorkbook.SelectWorksheet(AWorksheet: TsWorksheet);
begin
  if (AWorksheet <> nil) and (FWorksheets.IndexOf(AWorksheet) = -1) then
    raise Exception.Create('[TsWorkbook.SelectSheet] Worksheet does not belong to the workbook');
  FActiveWorksheet := AWorksheet;
  if Assigned(FOnSelectWorksheet) then FOnSelectWorksheet(self, AWorksheet);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the passed string is a valid worksheet name according to Excel
  (ODS seems to be a bit less restrictive, but if we follow Excel's convention
  we always have valid sheet names independent of the format.

  @param   AName                Name to be checked.
  @param   ReplaceDuplicateName If there exists already a sheet name equal to
                                AName then a number is added to AName such that
                                the name is unique.
  @return  TRUE if it is a valid worksheet name, FALSE otherwise
-------------------------------------------------------------------------------}
function TsWorkbook.ValidWorksheetName(var AName: String;
  ReplaceDuplicateName: Boolean = false): Boolean;
// see: http://stackoverflow.com/questions/451452/valid-characters-for-excel-sheet-names
var
  INVALID_CHARS: array [0..6] of char = ('[', ']', ':', '*', '?', '/', '\');
var
  i: Integer;
  unique: Boolean;
begin
  Result := false;

  // Name must not be empty
  if (AName = '') then
    exit;

  // Length must be less than 31 characters
  if UTF8Length(AName) > 31 then
    exit;

  // Name must not contain any of the INVALID_CHARS
  for i:=0 to High(INVALID_CHARS) do
    if UTF8Pos(INVALID_CHARS[i], AName) > 0 then
      exit;

  // Name must be unique
  unique := (GetWorksheetByName(AName) = nil);
  if not unique then
  begin
    if ReplaceDuplicateName then
    begin
      i := 0;
      repeat
        inc(i);
        unique := (GetWorksheetByName(AName + IntToStr(i)) = nil);
      until unique;
      AName := AName + IntToStr(i);
    end else
      exit;
  end;

  Result := true;
end;


{ String-to-cell/range conversion }

{@@ ----------------------------------------------------------------------------
  Analyses a string which can contain an array of cell ranges along with a
  worksheet name. Extracts the worksheet (if missing the "active" worksheet of
  the workbook is returned) and the cell's row and column indexes.

  @param  AText        General cell range string in Excel notation,
                       i.e. worksheet name + ! + cell in A1 notation.
                       Example: Sheet1!A1:A10; A1:A10 or A1 are valid as well.
  @param  AWorksheet   Pointer to the worksheet referred to by AText. If AText
                       does not contain the worksheet name, the active worksheet
                       of the workbook is returned
  @param  ARow, ACol   Zero-based row and column index of the cell identified
                       by ATest. If AText contains one ore more cell ranges
                       then the upper left corner of the first range is returned.
  @param  AListSeparator  Character to separate the cell blocks in the text
                       If #0 then the ListSeparator of the workbook's FormatSettings
                       is used.
  @returns TRUE if AText is a valid list of cell ranges, FALSE if not. If the
           result is FALSE then AWorksheet, ARow and ACol may have unpredictable
           values.
-------------------------------------------------------------------------------}
function TsWorkbook.TryStrToCell(AText: String; out AWorksheet: TsWorksheet;
  out ARow,ACol: Cardinal; AListSeparator: Char = #0): Boolean;
var
  ranges: TsCellRangeArray;
begin
  Result := TryStrToCellRanges(AText, AWorksheet, ranges, AListSeparator);
  if Result then
  begin
    ARow := ranges[0].Row1;
    ACol := ranges[0].Col1;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Analyses a string which can contain an array of cell ranges along with a
  worksheet name. Extracts the worksheet (if missing the "active" worksheet of
  the workbook is returned) and the cell range (or the first cell range, if there
  are several ranges).

  @param  AText        General cell range string in Excel notation,
                       i.e. worksheet name + ! + cell in A1 notation.
                       Example: Sheet1!A1:A10; A1:A10 or A1 are valid as well.
  @param  AWorksheet   Pointer to the worksheet referred to by AText. If AText
                       does not contain the worksheet name, the active worksheet
                       of the workbook is returned
  @param  ARange       TsCellRange records identifying the cell block. If AText
                       contains several cell ranges the first one is returned.
  @param  AListSeparator  Character to separate the cell blocks in the text
                       If #0 then the ListSeparator of the workbook's FormatSettings
                       is used.
  @returns TRUE if AText is a valid cell range, FALSE if not. If the
           result is FALSE then AWorksheet and ARange may have unpredictable
           values.
-------------------------------------------------------------------------------}
function TsWorkbook.TryStrToCellRange(AText: String; out AWorksheet: TsWorksheet;
  out ARange: TsCellRange; AListSeparator: Char = #0): Boolean;
var
  ranges: TsCellRangeArray;
begin
  Result := TryStrToCellRanges(AText, AWorksheet, ranges, AListSeparator);
  if Result then ARange := ranges[0];
end;

{@@ ----------------------------------------------------------------------------
  Analyses a string which can contain an array of cell ranges along with a
  worksheet name. Extracts the worksheet (if missing the "active" worksheet of
  the workbook is returned) and the range array.

  @param  AText        General cell range string in Excel notation,
                       i.e. worksheet name + ! + cell in A1 notation.
                       Example: Sheet1!A1:A10; A1:A10 or A1 are valid as well.
  @param  AWorksheet   Pointer to the worksheet referred to by AText. If AText
                       does not contain the worksheet name, the active worksheet
                       of the workbook is returned
  @param  ARanges      Array of TsCellRange records identifying the cell blocks
  @param  AListSeparator  Character to separate the cell blocks in the text
                       If #0 then the ListSeparator of the workbook's FormatSettings
                       is used.
  @returns TRUE if AText is a valid list of cell ranges, FALSE if not. If the
           result is FALSE then AWorksheet and ARanges may have unpredictable
           values.
-------------------------------------------------------------------------------}
function TsWorkbook.TryStrToCellRanges(AText: String; out AWorksheet: TsWorksheet;
  out ARanges: TsCellRangeArray; AListSeparator: Char = #0): Boolean;
var
  i: Integer;
  L: TStrings;
  sheetname: String;
begin
  Result := false;
  AWorksheet := nil;
  SetLength(ARanges, 0);

  if AText = '' then
    exit;

  i := pos(SHEETSEPARATOR, AText);
  if i = 0 then
    AWorksheet := FActiveWorksheet
  else begin
    sheetname := Copy(AText, 1, i-1);
    if (sheetname <> '') and (sheetname[1] = '''') then
      Delete(sheetname, 1, 1);
    if (sheetname <> '') and (sheetname[Length(sheetname)] = '''') then
      Delete(sheetname, Length(sheetname), 1);
    AWorksheet := GetWorksheetByName(sheetname);
    if AWorksheet = nil then
      exit;
    AText := Copy(AText, i+1, Length(AText));
  end;

  L := TStringList.Create;
  try
    if AListSeparator = #0 then
      L.Delimiter := FormatSettings.ListSeparator
    else
      L.Delimiter := AListSeparator;
    L.StrictDelimiter := true;
    L.DelimitedText := AText;
    if L.Count = 0 then
    begin
      AWorksheet := nil;
      exit;
    end;
    SetLength(ARanges, L.Count);
    for i:=0 to L.Count-1 do begin
      if pos(':', L[i]) = 0 then begin
        Result := ParseCellString(L[i], ARanges[i].Row1, ARanges[i].Col1);
        if Result then begin
          ARanges[i].Row2 := ARanges[i].Row1;
          ARanges[i].Col2 := ARanges[i].Col1;
        end;
      end else
        Result := ParseCellRangeString(L[i], ARanges[i]);
      if not Result then begin
        SetLength(ARanges, 0);
        AWorksheet := nil;
        exit;
      end;
    end;
  finally
    L.Free;
  end;
end;


{ Format handling }

{@@ ----------------------------------------------------------------------------
  Adds the specified format record to the internal list and returns the index
  in the list. If the record had already been added before the function only
  returns the index.
-------------------------------------------------------------------------------}
function TsWorkbook.AddCellFormat(const AValue: TsCellFormat): Integer;
begin
  Result := FCellFormatList.Add(AValue);
end;

{@@ ----------------------------------------------------------------------------
  Returns the contents of the format record with the specified index.
-------------------------------------------------------------------------------}
function TsWorkbook.GetCellFormat(AIndex: Integer): TsCellFormat;
begin
  Result := FCellFormatList.Items[AIndex]^;
end;

{@@ ----------------------------------------------------------------------------
  Returns a string describing the cell format with the specified index.
-------------------------------------------------------------------------------}
function TsWorkbook.GetCellFormatAsString(AIndex: Integer): String;
var
  fmt: PsCellFormat;
  cb: TsCellBorder;
  s: String;
  numFmt: TsNumFormatParams;
begin
  Result := '';
  fmt := GetPointerToCellFormat(AIndex);
  if fmt = nil then
    exit;

  if (uffFont in fmt^.UsedFormattingFields) then
    Result := Format('%s; Font%d', [Result, fmt^.FontIndex]);
  if (uffBackground in fmt^.UsedFormattingFields) then begin
    Result := Format('%s; Bg %s', [Result, GetColorName(fmt^.Background.BgColor)]);
    Result := Format('%s; Fg %s', [Result, GetColorName(fmt^.Background.FgColor)]);
    Result := Format('%s; Pattern %s', [Result, GetEnumName(TypeInfo(TsFillStyle), ord(fmt^.Background.Style))]);
  end;
  if (uffHorAlign in fmt^.UsedFormattingfields) then
    Result := Format('%s; %s', [Result, GetEnumName(TypeInfo(TsHorAlignment), ord(fmt^.HorAlignment))]);
  if (uffVertAlign in fmt^.UsedFormattingFields) then
    Result := Format('%s; %s', [Result, GetEnumName(TypeInfo(TsVertAlignment), ord(fmt^.VertAlignment))]);
  if (uffWordwrap in fmt^.UsedFormattingFields) then
    Result := Format('%s; Word-wrap', [Result]);
  if (uffNumberFormat in fmt^.UsedFormattingFields) then
  begin
    numFmt := GetNumberFormat(fmt^.NumberFormatIndex);
    if numFmt <> nil then
      Result := Format('%s; %s (%s)', [Result,
        GetEnumName(TypeInfo(TsNumberFormat), ord(numFmt.NumFormat)),
        numFmt.NumFormatStr
      ])
    else
      Result := Format('%s; %s', [Result, 'nfGeneral']);
  end else
    Result := Format('%s; %s', [Result, 'nfGeneral']);
  if (uffBorder in fmt^.UsedFormattingFields) then
  begin
    s := '';
    for cb in fmt^.Border do
      if s = '' then s := GetEnumName(TypeInfo(TsCellBorder), ord(cb))
        else s := s + '+' + GetEnumName(TypeInfo(TsCellBorder), ord(cb));
    Result := Format('%s; %s', [Result, s]);
  end;
  if Result <> '' then Delete(Result, 1, 2);
end;

{@@ ----------------------------------------------------------------------------
  Returns the count of format records used all over the workbook
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumCellFormats: Integer;
begin
  Result := FCellFormatList.Count;
end;

{@@ ----------------------------------------------------------------------------
  Returns a pointer to the format record with the specified index
-------------------------------------------------------------------------------}
function TsWorkbook.GetPointerToCellFormat(AIndex: Integer): PsCellFormat;
begin
  Result := FCellFormatList.Items[AIndex];
end;


{ Font handling }

{@@ ----------------------------------------------------------------------------
  Adds a font to the font list. Returns the index in the font list.

  @param AFontName  Name of the font (like 'Arial')
  @param ASize      Size of the font in points
  @param AStyle     Style of the font, a combination of TsFontStyle elements
  @param AColor     RGB valoe of the font color
  @return           Index of the font in the workbook's font list
-------------------------------------------------------------------------------}
function TsWorkbook.AddFont(const AFontName: String; ASize: Single;
  AStyle: TsFontStyles; AColor: TsColor): Integer;
var
  fnt: TsFont;
begin
  fnt := TsFont.Create;
  fnt.FontName := AFontName;
  fnt.Size := ASize;
  fnt.Style := AStyle;
  fnt.Color := AColor;
  Result := AddFont(fnt);
end;

{@@ ----------------------------------------------------------------------------
  Adds a font to the font list. Returns the index in the font list.

  @param AFont      TsFont record containing all font parameters
  @return           Index of the font in the workbook's font list
-------------------------------------------------------------------------------}
function TsWorkbook.AddFont(const AFont: TsFont): Integer;
begin
  result := FFontList.Add(AFont);
end;

{@@ ----------------------------------------------------------------------------
  Deletes a font.
  Use with caution because this will screw up the font assignment to cells.
  The only legal reason to call this method is from a reader of a file format
  in which the missing font #4 of BIFF does exist.
-------------------------------------------------------------------------------}
procedure TsWorkbook.DeleteFont(AFontIndex: Integer);
var
  fnt: TsFont;
begin
  if AFontIndex < FFontList.Count then
  begin
    fnt := TsFont(FFontList.Items[AFontIndex]);
    if fnt <> nil then fnt.Free;
    FFontList.Delete(AFontIndex);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the font with the given specification is already contained in
  the font list. Returns the index, or -1 if not found.

  @param AFontName  Name of the font (like 'Arial')
  @param ASize      Size of the font in points
  @param AStyle     Style of the font, a combination of TsFontStyle elements
  @param AColor     RGB value of the font color
  @return           Index of the font in the font list, or -1 if not found.
-------------------------------------------------------------------------------}
function TsWorkbook.FindFont(const AFontName: String; ASize: Single;
  AStyle: TsFontStyles; AColor: TsColor): Integer;
const
  EPS = 1e-3;
var
  fnt: TsFont;
begin
  for Result := 0 to FFontList.Count-1 do
  begin
    fnt := TsFont(FFontList.Items[Result]);
    if (fnt <> nil) and
       SameText(AFontName, fnt.FontName) and
       SameValue(ASize, fnt.Size, EPS) and   // careful when comparing floating point numbers
      (AStyle = fnt.Style) and
      (AColor = fnt.Color)
    then
      exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Initializes the font list by adding 5 fonts:

    0: default font
    1: like default font, but blue and underlined (for hyperlinks)
    2: like default font, but bold
    3: like default font, but italic
-------------------------------------------------------------------------------}
procedure TsWorkbook.InitFonts;
var
  fntName: String;
  fntSize: Single;
begin
  // Memorize old default font
  with TsFont(FFontList.Items[0]) do
  begin
    fntName := FontName;
    fntSize := Size;
  end;

  // Remove current font list
  RemoveAllFonts;

  // Build new font list
  SetDefaultFont(fntName, fntSize);                      // FONT0: Default font
  AddFont(fntName, fntSize, [fssUnderline], scBlue);     // FONT1: Hyperlink font = blue & underlined
  AddFont(fntName, fntSize, [fssBold], scBlack);         // FONT2: Bold font
  AddFont(fntName, fntSize, [fssItalic], scBlack);       // FONT3: Italic font (not used directly)

  FBuiltinFontCount := FFontList.Count;
end;

{@@ ----------------------------------------------------------------------------
  Clears the list of fonts and releases their memory.
-------------------------------------------------------------------------------}
procedure TsWorkbook.RemoveAllFonts;
var
  i: Integer;
  fnt: TsFont;
begin
  for i := FFontList.Count-1 downto 0 do
  begin
    fnt := TsFont(FFontList.Items[i]);
    fnt.Free;
    FFontList.Delete(i);
  end;
  FBuiltinFontCount := 0;
end;

{@@ ----------------------------------------------------------------------------
  Replaces the built-in font at a specific index with different font parameters
-------------------------------------------------------------------------------}
procedure TsWorkbook.ReplaceFont(AFontIndex: Integer; AFontName: String;
  ASize: Single; AStyle: TsFontStyles; AColor: TsColor);
var
  fnt: TsFont;
begin
  if (AFontIndex < FBuiltinFontCount) and (AFontIndex <> 4) then
  begin
    fnt := TsFont(FFontList[AFontIndex]);
    fnt.FontName := AFontName;
    fnt.Size := ASize;
    fnt.Style := AStyle;
    fnt.Color := AColor;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Defines the default font. This is the font with index 0 in the FontList.
  The next built-in fonts will have the same font name and size
-------------------------------------------------------------------------------}
procedure TsWorkbook.SetDefaultFont(const AFontName: String; ASize: Single);
var
  i: Integer;
begin
  if FFontList.Count = 0 then
    AddFont(AFontName, ASize, [], scBlack)
  else
  for i:=0 to FBuiltinFontCount-1 do
    if (i <> 4) and (i < FFontList.Count) then
      with TsFont(FFontList[i]) do
      begin
        FontName := AFontName;
        Size := ASize;
      end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the count of built-in fonts (default font, hyperlink font, bold font
  by default).
-------------------------------------------------------------------------------}
function TsWorkbook.GetBuiltinFontCount: Integer;
begin
  Result := FBuiltinFontCount;
end;

{@@ ----------------------------------------------------------------------------
  Returns the default font. This is the first font (index 0) in the font list
-------------------------------------------------------------------------------}
function TsWorkbook.GetDefaultFont: TsFont;
begin
  Result := GetFont(0);
end;

{@@ ----------------------------------------------------------------------------
  Returns the point size of the default font
-------------------------------------------------------------------------------}
function TsWorkbook.GetDefaultFontSize: Single;
begin
  Result := GetFont(0).Size;
end;

{@@ ----------------------------------------------------------------------------
  Returns the font with the given index.

  @param  AIndex   Index of the font to be considered
  @return Record containing all parameters of the font (or nil if not found).
-------------------------------------------------------------------------------}
function TsWorkbook.GetFont(AIndex: Integer): TsFont;
begin
  if (AIndex >= 0) and (AIndex < FFontList.Count) then
    Result := FFontList.Items[AIndex]
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Returns a string which identifies the font with a given index.

  @param  AIndex    Index of the font
  @return String with font name, font size etc.
-------------------------------------------------------------------------------}
function TsWorkbook.GetFontAsString(AIndex: Integer): String;
var
  fnt: TsFont;
begin
  fnt := GetFont(AIndex);
  if fnt <> nil then begin
    Result := Format('%s; size %.1g; %s', [
      fnt.FontName, fnt.Size, GetColorName(fnt.Color)]);
    if (fssBold in fnt.Style) then Result := Result + '; bold';
    if (fssItalic in fnt.Style) then Result := Result + '; italic';
    if (fssUnderline in fnt.Style) then Result := Result + '; underline';
    if (fssStrikeout in fnt.Style) then result := Result + '; strikeout';
  end else
    Result := '';
end;

{@@ ----------------------------------------------------------------------------
  Returns the count of registered fonts
-------------------------------------------------------------------------------}
function TsWorkbook.GetFontCount: Integer;
begin
  Result := FFontList.Count;
end;

{@@ ----------------------------------------------------------------------------
  Returns the hypertext font. This is font with index 6 in the font list
-------------------------------------------------------------------------------}
function TsWorkbook.GetHyperlinkFont: TsFont;
begin
  Result := GetFont(HYPERLINK_FONTINDEX);
end;


{@@ ----------------------------------------------------------------------------
  Adds a number format to the internal list. Returns the list index if already
  present, or creates a new format item and returns its index.
-------------------------------------------------------------------------------}
function TsWorkbook.AddNumberFormat(AFormatStr: String): Integer;
begin
  if AFormatStr = '' then
    Result := -1  // General number format is not stored
  else
    Result := TsNumFormatList(FNumFormatList).AddFormat(AFormatStr);
end;

{@@ ----------------------------------------------------------------------------
  Returns the parameters of the number format stored in the NumFormatList at the
  specified index.
  "General" number format is returned as nil.
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumberFormat(AIndex: Integer): TsNumFormatParams;
begin
  if (AIndex >= 0) and (AIndex < FNumFormatList.Count) then
    Result := TsNumFormatParams(FNumFormatList.Items[AIndex])
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Returns the count of number format records stored in the NumFormatList
-------------------------------------------------------------------------------}
function TsWorkbook.GetNumberFormatCount: Integer;
begin
  Result := FNumFormatList.Count;
end;
                            (*
{@@ ----------------------------------------------------------------------------
  Adds a color to the palette and returns its palette index, but only if the
  color does not already exist - in this case, it returns the index of the
  existing color entry.
  The color must in little-endian notation (like TColor of the graphics units)

  @param   AColorValue   Number containing the rgb code of the color to be added
  @return  Index of the new (or already existing) color item
-------------------------------------------------------------------------------}
function TsWorkbook.AddColorToPalette(AColorValue: TsColorValue): TsColor;
var
  n: Integer;
begin
  n := Length(FPalette);

  // Look look for the color. Is it already in the existing palette?
  if n > 0 then
    for Result := 0 to n-1 do
      if FPalette[Result] = AColorValue then
        exit;

  // No --> Add it to the palette.

  // Do not overwrite Excel's built-in system colors
  case n of
    DEF_FOREGROUND_COLOR:
      begin
        SetLength(FPalette, n+3);
        FPalette[n] := DEF_FOREGROUND_COLORVALUE;
        FPalette[n+1] := DEF_BACKGROUND_COLORVALUE;
        FPalette[n+2] := AColorValue;
      end;
    DEF_BACKGROUND_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_BACKGROUND_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    DEF_CHART_FOREGROUND_COLOR:
      begin
        SetLength(FPalette, n+4);
        FPalette[n] := DEF_CHART_FOREGROUND_COLORVALUE;
        FPalette[n+1] := DEF_CHART_BACKGROUND_COLORVALUE;
        FPalette[n+2] := DEF_CHART_NEUTRAL_COLORVALUE;
        FPalette[n+3] := AColorValue;
      end;
    DEF_CHART_BACKGROUND_COLOR:
      begin
        SetLength(FPalette, n+3);
        FPalette[n] := DEF_CHART_BACKGROUND_COLORVALUE;
        FPalette[n+1] := DEF_CHART_NEUTRAL_COLORVALUE;
        FPalette[n+2] := AColorValue;
      end;
    DEF_CHART_NEUTRAL_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_CHART_NEUTRAL_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    DEF_TOOLTIP_TEXT_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_TOOLTIP_TEXT_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    DEF_FONT_AUTOMATIC_COLOR:
      begin
        SetLength(FPalette, n+2);
        FPalette[n] := DEF_FONT_AUTOMATIC_COLORVALUE;
        FPalette[n+1] := AColorValue;
      end;
    else
      begin
        SetLength(FPalette, n+1);
        FPalette[n] := AColorValue;
      end;
  end;
  Result := Length(FPalette) - 1;

  if Assigned(FOnChangePalette) then FOnChangePalette(self);
end;
                          *)
{@@ ----------------------------------------------------------------------------
  Adds a (simple) error message to an internal list

  @param   AMsg   Error text to be stored in the list
-------------------------------------------------------------------------------}
procedure TsWorkbook.AddErrorMsg(const AMsg: String);
begin
  FLog.Add(AMsg);
end;

{@@ ----------------------------------------------------------------------------
  Adds an error message composed by means of format codes to an internal list

  @param   AMsg   Error text to be stored in the list
  @param   Args   Array of arguments to be used by the Format() function
-------------------------------------------------------------------------------}
procedure TsWorkbook.AddErrorMsg(const AMsg: String; const Args: Array of const);
begin
  FLog.Add(Format(AMsg, Args));
end;

{@@ ----------------------------------------------------------------------------
  Clears the internal error message list
-------------------------------------------------------------------------------}
procedure TsWorkbook.ClearErrorList;
begin
  FLog.Clear;
end;

{@@ ----------------------------------------------------------------------------
  Getter to retrieve the error messages collected during reading/writing
-------------------------------------------------------------------------------}
function TsWorkbook.GetErrorMsg: String;
begin
  Result := FLog.Text;
end;
                           (*
{@@ ----------------------------------------------------------------------------
  Converts a fpspreadsheet color into into a string RRGGBB.
  Note that colors are written to xls files as ABGR (where A is 0).
  if the color is scRGBColor the color value is taken from the argument
  ARGBColor, otherwise from the palette entry for the color index.
-------------------------------------------------------------------------------}
function TsWorkbook.FPSColorToHexString(AColor: TsColor;
  ARGBColor: TFPColor): string;
type
  TRgba = packed record Red, Green, Blue, A: Byte end;
var
  colorvalue: TsColorValue;
  r,g,b: Byte;
begin
  if AColor = scRGBColor then
  begin
    r := ARGBColor.Red div $100;
    g := ARGBColor.Green div $100;
    b := ARGBColor.Blue div $100;
  end else
  begin
    colorvalue := GetPaletteColor(AColor);
    r := TRgba(colorvalue).Red;
    g := TRgba(colorvalue).Green;
    b := TRgba(colorvalue).Blue;
  end;
  Result := Format('%.2x%.2x%.2x', [r, g, b]);
end;

{@@ ----------------------------------------------------------------------------
  Returns the name of the color pointed to by the given color index.
  If the name is not known the hex string is returned as RRGGBB.

  @param   AColorIndex   Palette index of the color considered
  @return  String identifying the color (a color name or, if unknown, a
           string showing the rgb components
-------------------------------------------------------------------------------}
function TsWorkbook.GetColorName(AColorIndex: TsColor): string;
begin
  case AColorIndex of
    scTransparent:
      Result := 'transparent';
    scNotDefined:
      Result := 'not defined';
    else
      GetColorName(GetPaletteColor(AColorIndex), Result);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the name of an rgb color value.
  If the name is not known the hex string is returned as RRGGBB.

  @param   AColorValue  rgb value of the color considered
  @param   AName        String identifying the color (a color name or, if
                        unknown, a string showing the rgb components
-------------------------------------------------------------------------------}
procedure TsWorkbook.GetColorName(AColorValue: TsColorValue; out AName: String);
type
  TRgba = packed record R,G,B,A: Byte; end;
var
  i: Integer;
begin
  // Find color value in default palette
  for i:=0 to High(DEFAULT_PALETTE) do
    // if found: get the color name from the default color names array
    if DEFAULT_PALETTE[i] = AColorValue then
    begin
      AName := DEFAULT_COLORNAMES[i];
      exit;
    end;

  // if not found: construct a string from rgb byte values.
  with TRgba(AColorValue) do
    AName := Format('%.2x%.2x%.2x', [R, G, B]);
end;

{@@ ----------------------------------------------------------------------------
  Converts the palette color of the given index to a string that can be used
  in HTML code. For ODS.

  @param  AColorIndex Index of the color considered
  @return A HTML-compatible string identifying the color.
          "Red", for example, is returned as '#FF0000';
-------------------------------------------------------------------------------}
function TsWorkbook.GetPaletteColorAsHTMLStr(AColorIndex: TsColor): String;
begin
  Result := ColorToHTMLColorStr(GetPaletteColor(AColorIndex));
end;

{@@ ----------------------------------------------------------------------------
  Instructs the workbook to take colors from the default palette. Is called
  from ODS reader because ODS does not have a palette. Without a palette the
  color constants (scRed etc.) would not be correct any more.
-------------------------------------------------------------------------------}
procedure TsWorkbook.UseDefaultPalette;
begin
  UsePalette(@DEFAULT_PALETTE, Length(DEFAULT_PALETTE), false);
end;

{@@ ----------------------------------------------------------------------------
  Instructs the Workbook to take colors from the palette pointed to by the
  parameter APalette
  This palette is only used for writing. When reading the palette found in the
  file is used.

  @param  APalette      Pointer to the array of TsColorValue numbers which will
                        become the new palette
  @param  APaletteCount Count of numbers in the source palette
  @param  ABigEnding    If true, indicates that the source palette is in
                        big-endian notation. The methods inverts the rgb
                        components to little-endian which is used by
                        fpspreadsheet internally.
-------------------------------------------------------------------------------}
procedure TsWorkbook.UsePalette(APalette: PsPalette; APaletteCount: Word;
  ABigEndian: Boolean);
var
  i: Integer;
begin
  if APaletteCount > 64 then
    raise Exception.Create('Due to Excel-compatibility, palettes cannot have more then 64 colors.');

 {$IFOPT R+}
  {$DEFINE RNGCHECK}
 {$ENDIF}
  SetLength(FPalette, APaletteCount);
  if ABigEndian then
    for i:=0 to APaletteCount-1 do
     {$IFDEF RNGCHECK}
      {$R-}
     {$ENDIF}
      FPalette[i] := LongRGBToExcelPhysical(APalette^[i])
     {$IFDEF RNGCHECK}
      {$R+}
     {$ENDIF}
  else
    for i:=0 to APaletteCount-1 do
     {$IFDEF RNGCHECK}
      {$R-}
     {$ENDIF}
      FPalette[i] := APalette^[i];
     {$IFDEF RNGCHECK}
      {$R+}
     {$ENDIF}

  if Assigned(FOnChangePalette) then FOnChangePalette(self);
end;

{@@ ----------------------------------------------------------------------------
  Checks whether a given color is used somewhere within the entire workbook

  @param  AColorIndex   Palette index of the color
  @result True if the color is used by at least one cell, false if not.
-------------------------------------------------------------------------------}
function TsWorkbook.UsesColor(AColorIndex: TsColor): Boolean;
var
  sheet: TsWorksheet;
  cell: PCell;
  i: Integer;
  fnt: TsFont;
  b: TsCellBorder;
  fmt: PsCellFormat;
begin
  Result := true;
  for i:=0 to GetWorksheetCount-1 do
  begin
    sheet := GetWorksheetByIndex(i);
    for cell in sheet.Cells do
    begin
      fmt := GetPointerToCellFormat(cell^.FormatIndex);
      if (uffBackground in fmt^.UsedFormattingFields) then
      begin
        if fmt^.Background.BgColor = AColorIndex then exit;
        if fmt^.Background.FgColor = AColorIndex then exit;
      end;
      if (uffBorder in fmt^.UsedFormattingFields) then
        for b in TsCellBorders do
          if fmt^.BorderStyles[b].Color = AColorIndex then
            exit;
      if (uffFont in fmt^.UsedFormattingFields) then
      begin
        fnt := GetFont(fmt^.FontIndex);
        if fnt.Color = AColorIndex then
          exit;
      end;
    end;
  end;
  Result := false;
end;
   *)

{*******************************************************************************
*                          TsBasicSpreadReaderWriter                           *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Constructor of the reader/writer. Has the workbook to be read/written as a
  parameter to apply the localization information found in its FormatSettings.

  @param AWorkbook  Workbook into which the file is being read or from with the
                    file is written. This parameter is passed from the workbook
                    which creates the reader/writer.
-------------------------------------------------------------------------------}
constructor TsBasicSpreadReaderWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  { A good starting point valid for many formats ... }
  FLimitations.MaxColCount := 256;
  FLimitations.MaxRowCount := 65536;
  FLimitations.MaxPaletteSize := MaxInt;
end;

{@@ ----------------------------------------------------------------------------
  Returns a record containing limitations of the specific file format of the
  writer.
-------------------------------------------------------------------------------}
function TsBasicSpreadReaderWriter.Limitations: TsSpreadsheetFormatLimitations;
begin
  Result := FLimitations;
end;


{*******************************************************************************
*                             TsBasicSpreadWriter                              *
*******************************************************************************}

{@@ ----------------------------------------------------------------------------
  Checks limitations of the writer, e.g max row/column count
-------------------------------------------------------------------------------}
procedure TsBasicSpreadWriter.CheckLimitations;
var
  lastCol, lastRow: Cardinal;
begin
  Workbook.GetLastRowColIndex(lastRow, lastCol);

  // Check row count
  if lastRow >= FLimitations.MaxRowCount then
    Workbook.AddErrorMsg(rsMaxRowsExceeded, [lastRow+1, FLimitations.MaxRowCount]);

  // Check column count
  if lastCol >= FLimitations.MaxColCount then
    Workbook.AddErrorMsg(rsMaxColsExceeded, [lastCol+1, FLimitations.MaxColCount]);
end;


initialization

finalization
  SetLength(GsSpreadFormats, 0);

end.

