unit sCSVParamsForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ButtonPanel, ExtCtrls, ComCtrls, StdCtrls,
  fpsCSV,
  sCtrls, sFormatsettingsForm;

type

  { TCSVParamsForm }

  TCSVParamsForm = class(TForm)
    ButtonPanel: TButtonPanel;
    CbAutoDetectNumberFormat: TCheckBox;
    CbLongDateFormat: TComboBox;
    CbLongTimeFormat: TComboBox;
    EdCurrencySymbol: TEdit;
    CbShortTimeFormat: TComboBox;
    CbShortDateFormat: TComboBox;
    CbDecimalSeparator: TComboBox;
    CbDateSeparator: TComboBox;
    CbTimeSeparator: TComboBox;
    CbThousandSeparator: TComboBox;
    CbLineEnding: TComboBox;
    CbQuoteChar: TComboBox;
    CbDelimiter: TComboBox;
    EdTRUE: TEdit;
    EdFALSE: TEdit;
    EdNumFormat: TEdit;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    LblDateTimeSample: TLabel;
    LblDecimalSeparator: TLabel;
    LblDecimalSeparator1: TLabel;
    LblDecimalSeparator2: TLabel;
    LblCurrencySymbol: TLabel;
    LblShortMonthNames: TLabel;
    LblLongDayNames: TLabel;
    LblShortDayNames: TLabel;
    LblNumFormat1: TLabel;
    LblNumFormat2: TLabel;
    LblNumFormat3: TLabel;
    LblNumFormat4: TLabel;
    LblLongMonthNames: TLabel;
    LblThousandSeparator: TLabel;
    LblNumFormat: TLabel;
    LblQuoteChar: TLabel;
    LblNumFormatInfo: TLabel;
    PageControl: TPageControl;
    PgGeneralParams: TTabSheet;
    PgNumberParams: TTabSheet;
    PgDateTimeParams: TTabSheet;
    PgBoolParams: TTabSheet;
    RgDetectContentType: TRadioGroup;
    PgCurrency: TTabSheet;
    procedure DateTimeFormatChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
    FSampleDateTime: TDateTime;
    FDateFormatSample: String;
    FTimeFormatSample: String;
    FEdLongMonthNames: TMonthDayNamesEdit;
    FEdShortMonthNames: TMonthDayNamesEdit;
    FEdLongDayNames: TMonthDayNamesEdit;
    FEdShortDayNames: TMonthDayNamesEdit;
    procedure DateSeparatorToFormatSettings(var ASettings: TFormatSettings);
    procedure DecimalSeparatorToFormatSettings(var ASettings: TFormatSettings);
//    function GetCurrencySymbol: String;
    procedure ThousandSeparatorToFormatSettings(var ASettings: TFormatSettings);
    procedure TimeSeparatorToFormatSettings(var ASettings: TFormatSettings);
  public
    { public declarations }
    procedure GetParams(var AParams: TsCSVParams);
    procedure SetParams(const AParams: TsCSVParams);
  end;

var
  CSVParamsForm: TCSVParamsForm;

implementation

uses
  fpsUtils;

resourcestring
  rsLikeSpreadsheet = 'like spreadsheet';

  {
const
  CURR_VALUE = 100.0;
   }

var
  CSVParamsPageIndex: Integer = 0;


{ TCSVParamsForm }

procedure TCSVParamsForm.DateSeparatorToFormatSettings(var ASettings: TFormatSettings);
begin
  case CbDateSeparator.ItemIndex of
    0: ASettings.DateSeparator := #0;
    1: ASettings.DateSeparator := '.';
    2: ASettings.DateSeparator := '-';
    3: ASettings.DateSeparator := '/';
    else ASettings.DateSeparator := CbDateSeparator.Text[1];
  end;
end;

procedure TCSVParamsForm.DecimalSeparatorToFormatSettings(var ASettings: TFormatSettings);
begin
  case CbDecimalSeparator.ItemIndex of
    0: ASettings.DecimalSeparator := #0;
    1: ASettings.DecimalSeparator := '.';
    2: ASettings.DecimalSeparator := ',';
    else ASettings.DecimalSeparator := CbDecimalSeparator.Text[1];
  end;
end;

procedure TCSVParamsForm.DateTimeFormatChange(Sender: TObject);
var
  fs: TFormatSettings;
  ctrl: TWinControl;
  dt: TDateTime;
  arr: Array[1..12] of String;
  i: Integer;
begin
  fs := DefaultFormatSettings;
  if CbLongDateFormat.ItemIndex <> 0 then
    fs.LongDateFormat := CbLongDateFormat.Text;
  if CbShortDateFormat.ItemIndex <> 0 then
    fs.ShortDateFormat := CbShortDateFormat.Text;
  if CbLongTimeFormat.ItemIndex <> 0 then
    fs.LongTimeFormat := CbLongTimeFormat.Text;
  if CbShortTimeFormat.ItemIndex <> 0 then
    fs.ShortTimeFormat := CbShortTimeFormat.Text;
  if CbDateSeparator.ItemIndex <> 0 then
    DateSeparatorToFormatSettings(fs);
  if CbTimeSeparator.ItemIndex <> 0 then
    TimeSeparatorToFormatSettings(fs);

  if FEdLongMonthNames.Text <> rsLikeSpreadsheet then begin
    arr[1] := '';  // to silence the compiler
    FEdLongMonthNames.GetNames(arr);
    for i:=1 to 12 do
      if arr[i] <> '' then fs.LongMonthNames[i] := arr[i];
  end;
  if FEdShortMonthNames.Text <> rsLikeSpreadsheet then begin
    FEdShortMonthNames.GetNames(arr);
    for i:=1 to 12 do
      if arr[i] <> '' then fs.ShortMonthNames[i] := arr[i];
  end;
  if FEdLongDayNames.Text <> rsLikeSpreadsheet then begin
    FEdLongDayNames.GetNames(arr);
    for i:=1 to 7 do
      if arr[i] <> '' then fs.LongDayNames[i] := arr[i];
  end;
  if FEdShortDayNames.Text <> rsLikeSpreadsheet then begin
    FEdShortDayNames.GetNames(arr);
    for i:=1 to 7 do
      if arr[i] <> '' then fs.ShortDayNames[i] := arr[i];
  end;

  dt := FSampleDateTime;
  ctrl := ActiveControl;
  if (ctrl = CbLongDateFormat) then
  begin
    FDateFormatSample := fs.LongDateFormat;
    LblDateTimeSample.Caption := FormatDateTime(FDateFormatSample, dt, fs);
  end
  else
  if (ctrl = CbShortDateFormat) then
  begin
    FDateFormatSample := fs.ShortDateFormat;
    LblDateTimeSample.Caption := FormatDateTime(FDateFormatSample, dt, fs);
  end
  else
  if (ctrl = CbDateSeparator) then
    LblDateTimeSample.Caption := FormatDateTime(FDateFormatSample, dt, fs)
  else
  if (ctrl = CbLongTimeFormat) then
  begin
    FTimeFormatSample := fs.LongTimeFormat;
    LblDateTimeSample.Caption := FormatDateTime(FTimeFormatSample, dt, fs);
  end
  else
  if (ctrl = CbShortTimeFormat) then
  begin
    FTimeFormatSample := fs.ShortTimeFormat;
    LblDateTimeSample.Caption := FormatDateTime(FTimeFormatSample, dt, fs);
  end
  else
  if (ctrl = CbTimeSeparator) then
    LblDateTimeSample.Caption := FormatDateTime(FTimeFormatSample, dt, fs)
  else
    LblDateTimeSample.Caption := FormatDateTime('c', dt, fs);

  Application.ProcessMessages;
end;

procedure TCSVParamsForm.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  Unused(Sender, CanClose);
  CSVParamsPageIndex := PageControl.ActivePageIndex;
end;

(*
procedure TCSVParamsForm.EdCurrencySymbolChange(Sender: TObject);
var
  sel: Integer;
begin
  sel := CbPosCurrencyFormat.ItemIndex;
  CbPosCurrencyFormat.Items.BeginUpdate;
  try
    CbPosCurrencyFormat.Items.Clear;
    BuildCurrencyFormatList(CbPosCurrencyFormat.Items, true, CURR_VALUE, GetCurrencySymbol);
    CbPosCurrencyFormat.Items.Insert(0, rsLikeSpreadsheet);
    CbPosCurrencyFormat.ItemIndex := sel;
  finally
    CbPosCurrencyFormat.Items.EndUpdate;
  end;

  sel := CbNegCurrencyFormat.ItemIndex;
  CbNegCurrencyFormat.Items.BeginUpdate;
  try
    CbNegCurrencyFormat.Items.Clear;
    BuildCurrencyFormatList(CbNegCurrencyFormat.Items, false, CURR_VALUE, GetCurrencySymbol);
    CbNegCurrencyFormat.Items.Insert(0, rsLikeSpreadsheet);
    CbNegCurrencyFormat.ItemIndex := sel;
  finally
    CbNegCurrencyFormat.Items.EndUpdate;
  end;
end;
*)

procedure TCSVParamsForm.FormCreate(Sender: TObject);
begin
  PageControl.ActivePageIndex := CSVParamsPageIndex;

//  CbNegCurrencyFormat.DropdownCount := 32;

  FEdLongMonthNames := TMonthDayNamesEdit.Create(self);
  with FEdLongMonthNames do
  begin
    Parent := PgDateTimeParams;
    Left :=  CbDateSeparator.Left;
    Top := CbDateSeparator.Top + 32;
   {$IFDEF LCL_FULLVERSION AND LCL_FULLVERSION > 1020600}
    Width := CbDateSeparator.Width;
   {$ELSE}
    Width := CbDateSeparator.Width - Button.Width;
   {$ENDIF}
    OnChange := @DateTimeFormatChange;
    OnEnter := @DateTimeFormatChange;
    TabOrder := CbDateSeparator.TabOrder + 1;
  end;
  LblLongMonthNames.FocusControl := FEdLongMonthNames;

  FEdShortMonthNames := TMonthDayNamesEdit.Create(self);
  with FEdShortMonthNames do
  begin
    Parent := PgDateTimeParams;
    Left :=  CbDateSeparator.Left;
    Top := CbDateSeparator.Top + 32*2;
    Width := FEdLongMonthNames.Width;
    TabOrder := CbDateSeparator.TabOrder + 2;
    OnChange := @DateTimeFormatChange;
    OnEnter := @DateTimeFormatChange;
  end;
  LblShortMonthNames.FocusControl := FEdShortMonthNames;

  FEdLongDayNames := TMonthDayNamesEdit.Create(self);
  with FEdLongDayNames do
  begin
    Parent := PgDateTimeParams;
    Left :=  CbDateSeparator.Left;
    Top := CbDateSeparator.Top + 32*3;
    Width := FEdLongMonthNames.Width;
    TabOrder := CbDateSeparator.TabOrder + 3;
    OnChange := @DateTimeFormatChange;
    OnEnter := @DateTimeFormatChange;
  end;
  LblLongDayNames.FocusControl := FEdLongDayNames;

  FEdShortDayNames := TMonthDayNamesEdit.Create(self);
  with FEdShortDayNames do
  begin
    Parent := PgDateTimeParams;
    Left :=  CbDateSeparator.Left;
    Top := CbDateSeparator.Top + 32*4;
    Width := FEdLongMonthNames.Width;
    TabOrder := CbDateSeparator.TabOrder + 4;
    OnChange := @DateTimeFormatChange;
    OnEnter := @DateTimeFormatChange;
  end;
  LblShortDayNames.FocusControl := FEdShortDayNames;

  FDateFormatSample := DefaultFormatSettings.LongDateFormat;
  FTimeFormatSample := DefaultFormatSettings.LongTimeFormat;
  FSampleDateTime := now();
end;

procedure TCSVParamsForm.GetParams(var AParams: TsCSVParams);
begin
  // Line endings
  case CbLineEnding.ItemIndex of
    0: AParams.LineEnding := leSystem;
    1: AParams.LineEnding := leCRLF;
    2: AParams.LineEnding := leCR;
    3: AParams.LineEnding := leLF;
  end;

  // Column delimiter
  case CbDelimiter.ItemIndex of
    0: AParams.Delimiter := ',';
    1: AParams.Delimiter := ';';
    2: AParams.Delimiter := ':';
    3: AParams.Delimiter := '|';
    4: AParams.Delimiter := #9;
  end;

  // Quote character
  case CbQuoteChar.ItemIndex of
    0: AParams.QuoteChar := #0;
    1: AParams.QuoteChar := '"';
    2: AParams.QuoteChar := '''';
  end;

  // Detect content type and convert
  AParams.DetectContentType := RgDetectContentType.ItemIndex <> 0;

  // Auto-detect number format
  AParams.AutoDetectNumberFormat := CbAutoDetectNumberFormat.Checked;

  // Number format
  AParams.NumberFormat := EdNumFormat.Text;

  // Decimal separator
  DecimalSeparatorToFormatSettings(AParams.FormatSettings);

  // Thousand separator
  ThousandSeparatorToFormatSettings(AParams.FormatSettings);

  // Currency symbol
  if (EdCurrencySymbol.Text = '') or (EdCurrencySymbol.Text = rsLikeSpreadsheet) then
    AParams.FormatSettings.CurrencyString := ''
  else
    AParams.FormatSettings.CurrencyString := UTF8ToAnsi(EdCurrencySymbol.Text);

  // Long date format string
  if (CbLongDateFormat.ItemIndex = 0) or (CbLongDateFormat.Text = '') then
    AParams.FormatSettings.LongDateFormat := ''
  else
    AParams.FormatSettings.LongDateFormat := CbLongDateFormat.Text;

  // Short date format string
  if (CbShortDateFormat.ItemIndex = 0) or (CbShortDateFormat.Text = '') then
    AParams.FormatSettings.ShortDateFormat := ''
  else
    AParams.FormatSettings.ShortDateFormat := CbShortDateFormat.Text;

  // Date separator
  DateSeparatorToFormatSettings(AParams.FormatSettings);

  // Long month names
  FEdLongMonthNames.GetNames(AParams.FormatSettings.LongMonthNames);

  // Short month names
  FEdShortMonthNames.GetNames(AParams.FormatSettings.ShortMonthNames);

  // Long day names
  FEdLongDayNames.GetNames(AParams.FormatSettings.LongDayNames);

  // Short day names
  FEdShortDayNames.GetNames(AParams.FormatSettings.ShortDayNames);

  // Long time format string
  if CbLongTimeFormat.ItemIndex = 0 then
    AParams.FormatSettings.LongTimeFormat := ''
  else
    AParams.FormatSettings.LongTimeFormat := CbLongTimeFormat.Text;

  // Short time format string
  if CbShortTimeFormat.ItemIndex = 0 then
    AParams.FormatSettings.ShortTimeFormat := ''
  else
    AParams.FormatSettings.ShortTimeFormat := CbShortTimeFormat.Text;

  // Time separator
  TimeSeparatorToFormatSettings(AParams.FormatSettings);

  // Text for "TRUE"
  AParams.TrueText := EdTRUE.Text;

  // Test for "FALSE"
  AParams.FalseText := EdFALSE.Text;
end;

procedure TCSVParamsForm.SetParams(const AParams: TsCSVParams);
begin
  // Line endings
  case AParams.LineEnding of
    leSystem: CbLineEnding.ItemIndex := 0;
    leCRLF  : CbLineEnding.ItemIndex := 1;
    leCR    : CbLineEnding.ItemIndex := 2;
    leLF    : CbLineEnding.ItemIndex := 3;
  end;

  // Column delimiter
  case AParams.Delimiter of
    ',' : CbDelimiter.ItemIndex := 0;
    ';' : CbDelimiter.ItemIndex := 1;
    ':' : CbDelimiter.ItemIndex := 2;
    '|' : CbDelimiter.ItemIndex := 3;
    #9  : CbDelimiter.ItemIndex := 4;
  end;

  // Quote character
  case AParams.QuoteChar of
    #0   : CbQuoteChar.ItemIndex := 0;
    '"'  : CbQuoteChar.ItemIndex := 1;
    '''' : CbQuoteChar.ItemIndex := 2;
  end;

  // Detect content type
  RgDetectContentType.ItemIndex := ord(AParams.DetectContentType);

  // Auto-detect number format
  CbAutoDetectNumberFormat.Checked := AParams.AutoDetectNumberFormat;

  // Number format
  EdNumFormat.Text := AParams.NumberFormat;

  // Decimal separator
  case AParams.FormatSettings.DecimalSeparator of
    #0  : CbDecimalSeparator.ItemIndex := 0;
    '.' : CbDecimalSeparator.ItemIndex := 1;
    ',' : CbDecimalSeparator.ItemIndex := 2;
    else  CbDecimalSeparator.Text := AParams.FormatSettings.DecimalSeparator;
  end;

  // Thousand separator
  case AParams.FormatSettings.ThousandSeparator of
    #0  : CbThousandSeparator.ItemIndex := 0;
    '.' : CbThousandSeparator.ItemIndex := 1;
    ',' : CbThousandSeparator.ItemIndex := 2;
    ' ' : CbThousandSeparator.ItemIndex := 3;
    else  CbThousandSeparator.Text := AParams.FormatSettings.ThousandSeparator;
  end;

  // Currency symbol
  if AParams.FormatSettings.CurrencyString = '' then
    EdCurrencySymbol.Text := rsLikeSpreadsheet
  else
    EdCurrencySymbol.Text := AnsiToUTF8(AParams.FormatSettings.CurrencyString);

  // Long date format
  if AParams.FormatSettings.LongDateFormat = '' then
    CbLongDateFormat.ItemIndex := 0
  else
    CbLongDateFormat.Text := AParams.FormatSettings.LongDateFormat;

  // Short date format
  if AParams.FormatSettings.ShortDateFormat = '' then
    CbShortDateFormat.ItemIndex := 0
  else
    CbShortDateFormat.Text := AParams.FormatSettings.ShortDateFormat;

  // Date separator
  case AParams.FormatSettings.DateSeparator of
    #0  : CbDateSeparator.ItemIndex := 0;
    '.' : CbDateSeparator.ItemIndex := 1;
    '-' : CbDateSeparator.ItemIndex := 2;
    '/' : CbDateSeparator.ItemIndex := 3;
    else  CbDateSeparator.Text := AParams.FormatSettings.DateSeparator;
  end;

  // Long month names
  FEdLongMonthNames.SetNames(AParams.FormatSettings.LongMonthNames, 12, false, rsLikeSpreadsheet);

  // Short month names
  FEdShortMonthNames.SetNames(AParams.FormatSettings.ShortMonthNames, 12, true, rsLikeSpreadsheet);

  // Long day names
  FEdLongDayNames.SetNames(AParams.FormatSettings.LongDayNames, 7, false, rsLikeSpreadsheet);

  // Short month names
  FEdShortDayNames.SetNames(AParams.FormatSettings.ShortDayNames, 7, true, rsLikeSpreadsheet);

  // Long time format
  if AParams.FormatSettings.LongTimeFormat = '' then
    CbLongTimeFormat.ItemIndex := 0
  else
    CbLongTimeFormat.Text := AParams.FormatSettings.LongTimeFormat;

  // Short time format
  if AParams.FormatSettings.ShortTimeFormat = '' then
    CbShortTimeFormat.ItemIndex := 0
  else
    CbShortTimeFormat.Text := AParams.FormatSettings.ShortTimeFormat;

  // Time separator
  case AParams.FormatSettings.TimeSeparator of
    #0  : CbTimeSeparator.ItemIndex := 0;
    '.' : CbTimeSeparator.ItemIndex := 1;
    '-' : CbTimeSeparator.ItemIndex := 2;
    '/' : CbTimeSeparator.ItemIndex := 3;
    ':' : CbTimeSeparator.ItemIndex := 4;
    else  CbTimeSeparator.Text := AParams.FormatSettings.TimeSeparator;
  end;

  // Text for "TRUE"
  EdTRUE.Text := AParams.TrueText;

  // Test for "FALSE"
  EdFALSE.Text := AParams.FalseText;

  // Update date/time sample display
  DateTimeFormatChange(nil);
end;

procedure TCSVParamsForm.ThousandSeparatorToFormatSettings(var ASettings: TFormatSettings);
begin
  case CbThousandSeparator.ItemIndex of
    0: ASettings.ThousandSeparator := #0;
    1: ASettings.ThousandSeparator := '.';
    2: ASettings.ThousandSeparator := ',';
    3: ASettings.ThousandSeparator := ' ';
    else ASettings.ThousandSeparator := CbThousandSeparator.Text[1];
  end;
end;

procedure TCSVParamsForm.TimeSeparatorToFormatSettings(var ASettings: TFormatSettings);
begin
  case CbTimeSeparator.ItemIndex of
    0: ASettings.TimeSeparator := #0;
    1: ASettings.TimeSeparator := '.';
    2: ASettings.TimeSeparator := '-';
    3: ASettings.TimeSeparator := '/';
    4: ASettings.TimeSeparator := ':';
    else ASettings.TimeSeparator := CbTimeSeparator.Text[1];
  end;
end;

initialization
  {$I scsvparamsform.lrs}

end.

