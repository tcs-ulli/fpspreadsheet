unit sNumFormatForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ButtonPanel,
  ExtCtrls, StdCtrls, Spin, Buttons, types, contnrs, inifiles,
  fpsTypes, fpsNumFormat, fpSpreadsheet;

type
  TsNumFormatCategory = (nfcNumber, nfcPercent, nfcScientific, nfcFraction,
    nfcCurrency, nfcDate, nfcTime);

  { TNumFormatForm }

  TNumFormatForm = class(TForm)
    ButtonPanel1: TButtonPanel;
    CbThousandSep: TCheckBox;
    CbNegRed: TCheckBox;
    CbCurrSymbol: TComboBox;
    EdNumFormatStr: TEdit;
    GbOptions: TGroupBox;
    GbFormatString: TGroupBox;
    GroupBox3: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    DetailsPanel: TPanel;
    Sample: TLabel;
    Label5: TLabel;
    LbCategory: TListBox;
    LbFormat: TListBox;
    Panel1: TPanel;
    Panel2: TPanel;
    EdDecimals: TSpinEdit;
    CurrSymbolPanel: TPanel;
    BtnAddCurrSymbol: TSpeedButton;
    Shape1: TShape;
    BtnAddFormat: TSpeedButton;
    BtnDeleteFormat: TSpeedButton;
    procedure BtnAddCurrSymbolClick(Sender: TObject);
    procedure BtnAddFormatClick(Sender: TObject);
    procedure BtnDeleteFormatClick(Sender: TObject);
    procedure CbCurrSymbolSelect(Sender: TObject);
    procedure CbNegRedClick(Sender: TObject);
    procedure CbThousandSepClick(Sender: TObject);
    procedure EdDecimalsChange(Sender: TObject);
    procedure EdNumFormatStrChange(Sender: TObject);
    procedure LbCategoryClick(Sender: TObject);
    procedure LbFormatClick(Sender: TObject);
    procedure LbFormatDrawItem(Control: TWinControl; Index: Integer;
      ARect: TRect; State: TOwnerDrawState);
  private
    { private declarations }
    FWorkbook: TsWorkbook;
    FSampleValue: Double;
    FGenerator: array[TsNumFormatCategory] of Double;
    FNumFormatStrOfList: String;
    FLockCount: Integer;
    function GetNumFormatStr: String;
    procedure SetNumFormatStr(const AValue: String);
  protected
    function FindNumFormat(ACategory: TsNumFormatCategory;
      ANumFormatStr: String): Integer;
    function FormatStrOfListIndex(AIndex: Integer): String;
    procedure ReplaceCurrSymbol;
    procedure ReplaceDecs;
    procedure SelectCategory(ACategory: TsNumFormatCategory);
    procedure SelectFormat(AIndex: Integer);
    procedure UpdateControls(ANumFormatParams: TsNumFormatParams);
    procedure UpdateSample(ANumFormatParams: TsNumFormatParams);
  public
    { public declarations }
    constructor Create(AOwner: TComponent); override;
    procedure SetData(ANumFormatStr: String; AWorkbook: TsWorkbook; ASample: Double);
    property NumFormatStr: String read GetNumFormatStr;
  end;

var
  NumFormatForm: TNumFormatForm;

procedure ReadNumFormatsFromIni(const AIniFile: TCustomIniFile);
procedure WriteNumFormatsToIni(const AIniFile: TCustomIniFile);

implementation

{$R *.lfm}

uses
  LCLType, Math, DateUtils, TypInfo,
  fpsUtils, fpsNumFormatParser, fpsCurrency,
  sCurrencyForm;

const
  BUILTIN_OFFSET = 1;
  USER_OFFSET = 1000;

var
  NumFormats: TStringList = nil;

procedure AddToList(ACategory: TsNumFormatCategory; AFormatStr: String;
  AOffset: Integer = BUILTIN_OFFSET);
begin
  if NumFormats.IndexOf(AFormatStr) = -1 then
    NumFormats.AddObject(AFormatStr, TObject(PtrInt(AOffset + ord(ACategory))));
end;

procedure InitNumFormats(AFormatSettings: TFormatSettings);
var
  copiedFormats: TStringList;
  nfs: String;
  data: PtrInt;
  i: Integer;
  fs: TFormatSettings absolute AFormatSettings;
begin
  copiedFormats := nil;

  // Store user-defined formats already added to NumFormats list
  if NumFormats <> nil then
  begin
    copiedFormats := TStringList.Create;
    for i:=0 to NumFormats.Count-1 do
    begin
      nfs := NumFormats.Strings[i];
      data := PtrInt(NumFormats.Objects[i]);
      if data >= USER_OFFSET then
        copiedFormats.AddObject(nfs, TObject(data));
    end;
    NumFormats.Free;
  end;

  NumFormats := TStringList.Create;

  // Add built-in formats
  AddToList(nfcNumber, 'General');
  AddToList(nfcNumber, '0');
  AddToList(nfcNumber, '0.0');
  AddToList(nfcNumber, '0.00');
  AddToList(nfcNumber, '0.000');
  AddToList(nfcNumber, '#,##0');
  AddToList(nfcNumber, '#,##0.0');
  AddToList(nfcNumber, '#,##0.00');
  AddToList(nfcNumber, '#,##0.000');

  AddToList(nfcPercent, '0%');
  AddToList(nfcPercent, '0.0%');
  AddToList(nfcPercent, '0.00%');
  AddToList(nfcPercent, '0.000%');

  AddToList(nfcScientific, '0E+0');
  AddToList(nfcScientific, '0E+00');
  AddToList(nfcScientific, '0E+000');
  AddToList(nfcScientific, '0.0E+0');
  AddToList(nfcScientific, '0.0E+00');
  AddToList(nfcScientific, '0.0E+000');
  AddToList(nfcScientific, '0.00E+0');
  AddToList(nfcScientific, '0.00E+00');
  AddToList(nfcScientific, '0.00E+000');
  AddToList(nfcScientific, '0.000E+0');
  AddToList(nfcScientific, '0.000E+00');
  AddToList(nfcScientific, '0.000E+000');
  AddToList(nfcScientific, '0E-0');
  AddToList(nfcScientific, '0E-00');
  AddToList(nfcScientific, '0E-000');
  AddToList(nfcScientific, '0.0E-0');
  AddToList(nfcScientific, '0.0E-00');
  AddToList(nfcScientific, '0.0E-000');
  AddToList(nfcScientific, '0.00E-0');
  AddToList(nfcScientific, '0.00E-00');
  AddToList(nfcScientific, '0.00E-000');
  AddToList(nfcScientific, '0.000E-0');
  AddToList(nfcScientific, '0.000E-00');
  AddToList(nfcScientific, '0.000E-000');

  AddToList(nfcFraction, '# ?/?');
  AddToList(nfcFraction, '# ??/??');
  AddToList(nfcFraction, '# ???/???');
  AddToList(nfcFraction, '# ?/2');
  AddToList(nfcFraction, '# ?/4');
  AddToList(nfcFraction, '# ?/8');
  AddToList(nfcFraction, '# ?/16');
  AddToList(nfcFraction, '# ?/32');
  AddToList(nfcFraction, '?/?');
  AddToList(nfcFraction, '?/??');
  AddToList(nfcFraction, '?/???');
  AddToList(nfcFraction, '?/2');
  AddToList(nfcFraction, '?/4');
  AddToList(nfcFraction, '?/8');
  AddToList(nfcFraction, '?/16');
  AddToList(nfcFraction, '?/32');

  AddToList(nfcCurrency, '#,##0 [$$];-#,##0 [$$]');
  AddToList(nfcCurrency, '#,##0.00 [$$];-#,##0.00 [$$]');
  AddToList(nfcCurrency, '#,##0 [$$];(#,##0) [$$]');
  AddToList(nfcCurrency, '#,##0.00 [$$];(#,##0.00) [$$]');
  AddToList(nfcCurrency, '#,##0 [$$];[red]-#,##0 [$$]');
  AddToList(nfcCurrency, '#,##0.00 [$$];[red]-#,##0.00 [$$]');
  AddToList(nfcCurrency, '#,##0 [$$];[red](#,##0) [$$]');
  AddToList(nfcCurrency, '#,##0.00 [$$];[red]-#,##0.00 [$$]');
  AddToList(nfcCurrency, '[$$] #,##0;[$$] -#,##0');
  AddToList(nfcCurrency, '[$$] #,##0.00;[$$] -#,##0.00');
  AddToList(nfcCurrency, '[$$] #,##0;[$$] (#,##0)');
  AddToList(nfcCurrency, '[$$] #,##0.00;[$$] (#,##0.00)');
  AddToList(nfcCurrency, '[$$] #,##0;[red][$$] -#,##0');
  AddToList(nfcCurrency, '[$$] #,##0.00;[red][$$] -#,##0.00');
  AddToList(nfcCurrency, '[$$] #,##0;[red][$$] (#,##0)');
  AddToList(nfcCurrency, '[$$] #,##0.00;[red][$$] -#,##0.00');

  AddToList(nfcDate, 'dddd, '+fs.LongDateFormat + ' ' + fs.ShortTimeFormat);
  AddToList(nfcDate, 'dddd, '+fs.ShortDateFormat + ' ' + fs.ShortTimeFormat);
  AddToList(nfcDate, 'dddd, '+fs.LongDateFormat);
  AddToList(nfcDate, 'dddd, '+fs.ShortDateFormat);
  AddToList(nfcDate, 'ddd., '+fs.LongDateFormat + ' ' + fs.ShortTimeFormat);
  AddToList(nfcDate, 'ddd., '+fs.ShortDateFormat + ' ' + fs.ShortTimeFormat);
  AddToList(nfcDate, 'ddd., '+fs.LongDateFormat);
  AddToList(nfcDate, 'ddd., '+fs.ShortDateFormat);
  AddToList(nfcDate, fs.LongDateFormat + ' ' + fs.ShortTimeFormat);
  AddToList(nfcDate, fs.ShortDateFormat + ' ' + fs.ShortTimeFormat);
  AddToList(nfcDate, fs.LongDateFormat);
  AddToList(nfcDate, fs.ShortDateFormat);
  AddToList(nfcDate, 'dd. mmmm');
  AddToList(nfcDate, 'dd. mmm.');
  AddToList(nfcDate, 'd. mmmm');
  AddToList(nfcDate, 'd. mmm.');
  AddToList(nfcDate, 'mmmm dd');
  AddToList(nfcDate, 'mmmm d');
  AddToList(nfcDate, 'mmm. dd');
  AddToList(nfcDate, 'mmm. d');
  AddToList(nfcDate, 'mmmm yyyy');
  AddToList(nfcDate, 'mmm. yy');
  AddToList(nfcDate, 'yyyy-mmm');
  AddToList(nfcDate, 'yy-mmm');

  AddToList(nfcTime, fs.LongTimeFormat);
  AddToList(nfcTime, fs.ShortTimeFormat);
  AddToList(nfcTime, AddAMPM(fs.LongTimeFormat, fs));
  AddToList(nfcTime, AddAMPM(fs.ShortTimeFormat, fs));
  AddToList(nfcTime, 'nn:ss');
  AddToList(nfcTime, 'nn:ss.0');
  AddToList(nfcTime, 'nn:ss.00');
  AddToList(nfcTime, 'nn:ss.000');
  AddToList(nfcTime, '[h]:nn');
  AddToList(nfcTime, '[h]:nn:ss');

  // Add user-defined formats
  if copiedFormats <> nil then
  begin
    for i:=0 to copiedFormats.Count-1 do begin
      nfs := copiedFormats.Strings[i];
      data := PtrInt(copiedFormats.Objects[i]);
      NumFormats.AddObject(nfs, TObject(PtrInt(data)));
    end;
    copiedFormats.Free;
  end;
end;

procedure DestroyNumFormats;
begin
  NumFormats.Free;
end;

{ Reads the user-defined number format strings from an ini file. }
procedure ReadNumFormatsFromIni(const AIniFile: TCustomIniFile);
var
  section: String;
  list: TStringList;
  cat: TsNumFormatCategory;
  i: Integer;
  nfs: String;
  scat: String;
begin
  if NumFormats = nil
    then NumFormats := TStringList.Create
    else NumFormats.Clear;

  list := TStringList.Create;
  try
    section := 'Built-in number formats';
    AIniFile.ReadSection(section, list);
    for i:=0 to list.Count-1 do begin
      scat := list.Names[i];
      nfs := list.Values[scat];
      cat := TsNumFormatCategory(GetEnumValue(TypeInfo(TsNumFormatCategory), scat));
      AddToList(cat, nfs, BUILTIN_OFFSET);
    end;

    list.Clear;
    section := 'User-defined number formats';
    AIniFile.ReadSection(section, list);
    for i:=0 to list.Count-1 do begin
      scat := list.Names[i];
      nfs := list.Values[scat];
      cat := TsNumFormatCategory(GetEnumValue(TypeInfo(TsNumFormatCategory), scat));
      AddToList(cat, nfs, USER_OFFSET);
    end;

  finally
    list.Free;
  end;
end;

procedure WriteNumFormatsToIni(const AIniFile: TCustomIniFile);
var
  data: PtrInt;
  section: String;
  i: Integer;
  cat: TsNumFormatCategory;
  scat: String;
  nfs: String;
begin
  section := 'Built-in number formats';
  for i:=0 to NumFormats.Count-1 do
  begin
    data := PtrInt(NumFormats.Objects[i]);
    if data < USER_OFFSET then
    begin
      cat := TsNumFormatCategory(data - BUILTIN_OFFSET);
      scat := Copy(GetEnumName(TypeInfo(TsNumFormatCategory), ord(cat)), 3, MaxInt);
      nfs := NumFormats.Strings[i];
      AIniFile.WriteString(section, scat, nfs);
    end;
  end;

  section := 'User-defined number formats';
  for i:=0 to NumFormats.Count-1 do
  begin
    data := PtrInt(NumFormats.Objects[i]);
    if data >= USER_OFFSET then
    begin
      cat := TsNumFormatCategory(data - USER_OFFSET);
      scat := Copy(GetEnumName(TypeInfo(TsNumFormatCategory), ord(cat)), 3, MaxInt);
      nfs := NumFormats.Strings[i];
      AIniFile.WriteString(section, scat, nfs);
    end;
  end;
end;



{ TNumFormatForm }

constructor TNumFormatForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FGenerator[nfcNumber] := -1234.123456;
  FGenerator[nfcPercent] := -0.123456789;
  FGenerator[nfcScientific] := -1234.5678;
  FGenerator[nfcFraction] := -1234; //-1.23456;
  FGenerator[nfcCurrency] := -1234.56789;
  FGenerator[nfcDate] := EncodeDate(YearOf(date), 1, 1);
  FGenerator[nfcTime] := EncodeTime(9, 0, 2, 235);
  GetRegisteredCurrencies(CbCurrSymbol.Items);
end;

procedure TNumFormatForm.BtnAddCurrSymbolClick(Sender: TObject);
var
  F: TCurrencyForm;
begin
  F := TCurrencyForm.Create(nil);
  try
    if F.ShowModal = mrOK then
    begin
      GetRegisteredCurrencies(CbCurrSymbol.Items);
      CbCurrSymbol.ItemIndex := CbCurrSymbol.Items.IndexOf(F.CurrencySymbol);
      ReplaceCurrSymbol;
    end;
  finally
    F.Free;
  end;
end;

procedure TNumFormatForm.BtnAddFormatClick(Sender: TObject);
var
  cat: TsNumFormatCategory;
  idx: Integer;
  nfs: String;
begin
  if LbCategory.ItemIndex > -1 then begin
    cat := TsNumFormatCategory(LbCategory.ItemIndex);
    nfs := EdNumFormatStr.Text;
    if nfs = '' then nfs := 'General';
    if NumFormats.IndexOf(nfs) = -1 then
    begin
      AddToList(cat, nfs, USER_OFFSET);
      SelectCategory(cat);    // Rebuilds the "Format" listbox
      idx := FindNumFormat(cat, nfs);
      SelectFormat(idx);
    end;
  end;
end;

procedure TNumFormatForm.BtnDeleteFormatClick(Sender: TObject);
var
  cat: TsNumFormatCategory;
  idx: Integer;
  nfs: String;
  n, i: Integer;
begin
  if LbCategory.ItemIndex > -1 then begin
    // Find in internal template list
    idx := NumFormats.IndexOf(EdNumFormatStr.Text);
    if idx > -1 then begin
      nfs := NumFormats.Strings[idx];
      n := PtrInt(NumFormats.Objects[idx]);
      if n >= USER_OFFSET
        then cat := TsNumFormatCategory(n - USER_OFFSET)
        else cat := TsNumFormatCategory(n - BUILTIN_OFFSET);
      i := FindNumFormat(cat, nfs);  // Index in format listbox
      // Delete from internal template list
      NumFormats.Delete(idx);

      // Rebuild format listbox (without the deleted item)
      SelectCategory(cat);
      if i >= LbFormat.Items.Count
        then SelectFormat(LbFormat.Items.Count-1)
        else SelectFormat(i);
    end;
  end;
end;
                                                      (*
{ The global stringlist "NumFormats" contains to format string templates along
  with information on the category and of being built-in or user-defined. }
procedure TNumFormatForm.BuildNumFormatLists(AWorkbook: TsWorkbook);
var
  cat: TsNumFormatCategory;
  nfs: String;
  n, i: Integer;
  isUserDef: Boolean;
  copiedNumFormats: TStringList;
begin
  copiedNumFormats := TStringList.Create;

  for cat in TsNumFormatCategory do
  begin
    FreeAndNil(FNumFormatLists[cat]);
    FNumFormatLists[cat] := TsNumFormatList.Create(AWorkbook, true);
  end;

  for i:=0 to NumFormats.Count-1 do
  begin
    nfs := NumFormats.Strings[i];
    n := PtrInt(NumFormats.Objects[i]);
    if n >= USER_OFFSET then
    begin
      isUserDef := true;
      // The category numbers of user-defined template items are offset by USER_OFFSET
      cat := TsNumFormatCategory(n - USER_OFFSET);
    end else
    begin
      isUserDef := false;
      // The category numbers of built-in template items are offset by BUILTIN_OFFSET
      cat := TsNumFormatCategory(n - BUILTIN_OFFSET);
    end;
    FNumFormatLists[cat].AddFormat(nfs);
  end;
end;
                *)

procedure TNumFormatForm.CbCurrSymbolSelect(Sender: TObject);
begin
  ReplaceCurrSymbol;
end;

procedure TNumFormatForm.CbNegRedClick(Sender: TObject);
var
  nfs: String;
  nfp: TsNumFormatParams;
begin
  if FLockCount > 0 then
    exit;

  if EdNumFormatStr.Text = '' then nfs := 'General' else nfs := EdNumFormatStr.Text;
  nfp := CreateNumFormatParams(nfs, FWorkbook.FormatSettings);
  if nfp <> nil then
    try
      nfp.SetNegativeRed(CbNegRed.Checked);
      EdNumFormatStr.Text := nfp.NumFormatStr;
      SelectCategory(TsNumFormatCategory(LbCategory.ItemIndex));  // to rebuild the format listbox
      UpdateSample(nfp);
    finally
      nfp.Free;
    end;
end;

procedure TNumFormatForm.CbThousandSepClick(Sender: TObject);
var
  nfs: String;
  nfp: TsNumFormatParams;
begin
  if FLockCount > 0 then
    exit;

  if EdNumFormatStr.Text = '' then nfs := 'General' else nfs := EdNumFormatStr.Text;
  nfp := CreateNumFormatParams(nfs, FWorkbook.FormatSettings);
  if nfp <> nil then
    try
      nfp.SetThousandSep(CbThousandSep.Checked);
      EdNumFormatStr.Text := nfp.NumFormatStr;
      SelectCategory(TsNumFormatCategory(LbCategory.ItemIndex));  // to rebuild the format listbox
      UpdateSample(nfp);
    finally
      nfp.Free;
    end;
end;

procedure TNumFormatForm.EdDecimalsChange(Sender: TObject);
begin
  if FLockCount > 0 then
    exit;
  ReplaceDecs;
end;

procedure TNumFormatForm.EdNumFormatStrChange(Sender: TObject);
var
  nfp: TsNumFormatParams;
begin
  nfp := CreateNumFormatParams(EdNumFormatStr.Text, FWorkbook.FormatSettings);
  try
    UpdateControls(nfp);
  finally
    nfp.Free;
  end;
end;

{ Returns the index of a specific number format string in the format listbox
  shown for a particular category }
function TNumFormatForm.FindNumFormat(ACategory: TsNumFormatCategory;
  ANumFormatStr: String): Integer;
var
  i: Integer;
  data: PtrInt;
  cat: TsNumFormatCategory;
  nfs: String;
begin
  Result := -1;
  if ANumFormatStr = '' then ANumFormatStr := 'General';
  for i := 0 to NumFormats.Count-1 do begin
    nfs := NumFormats.Strings[i];
    data := PtrInt(NumFormats.Objects[i]);
    if data >= USER_OFFSET then
      cat := TsNumFormatCategory(data - USER_OFFSET)
    else
      cat := TsNumFormatCategory(data - BUILTIN_OFFSET);
    if (cat = ACategory) then
      inc(Result);
    if SameText(nfs, ANumFormatStr) then
      exit;
  end;
end;

function TNumFormatForm.FormatStrOfListIndex(AIndex: Integer): String;
var
  idx: PtrInt;
begin
  if (AIndex >= 0) and (AIndex < LbFormat.Count) then
  begin
    idx := PtrInt(LbFormat.Items.Objects[AIndex]);
    Result := NumFormats.Strings[idx];
  end else
    Result := '';
end;

function TNumFormatForm.GetNumFormatStr: String;
begin
  Result := EdNumFormatStr.Text;
end;

procedure TNumFormatForm.LbCategoryClick(Sender: TObject);
begin
  SelectCategory(TsNumFormatCategory(LbCategory.ItemIndex));
end;

procedure TNumFormatForm.LbFormatClick(Sender: TObject);
begin
  SelectFormat(LbFormat.ItemIndex);
end;

procedure TNumFormatForm.LbFormatDrawItem(Control: TWinControl; Index: Integer;
  ARect: TRect; State: TOwnerDrawState);
var
  s: String;
  nfs: String;
  nfp: TsNumFormatParams;
  idx: PtrInt;
begin
  LbFormat.Canvas.Brush.Color := clWindow;
  LbFormat.Canvas.Font.Assign(LbFormat.Font);
  if State * [odSelected, odFocused] <> [] then
  begin
    LbFormat.Canvas.Font.Color := clHighlightText;
    LbFormat.Canvas.Brush.Color := clHighlight;
  end;
  if (Index > -1) and (Index < LbFormat.Items.Count) then
  begin
    s := LbFormat.Items[Index];
    idx := PtrInt(LbFormat.Items.Objects[Index]);
    nfs := NumFormats.Strings[idx];
    nfp := CreateNumFormatParams(nfs, FWorkbook.FormatSettings);
    try
      if (nfp <> nil) and (Length(nfp.Sections) > 1) and (nfp.Sections[1].Color = scRed) then
        LbFormat.Canvas.Font.Color := clRed;
    finally
      nfp.Free;
    end;
  end else
    s := '';
  LbFormat.Canvas.FillRect(ARect);
  LbFormat.Canvas.TextRect(ARect, ARect.Left+1, ARect.Top+1, s);
end;

procedure TNumFormatForm.ReplaceCurrSymbol;
var
   cs: String;
  i: Integer;
  nfp: TsNumFormatParams;
  data: PtrInt;
  cat: TsNumFormatCategory;
begin
  cs := CbCurrSymbol.Items[CbCurrSymbol.ItemIndex];
  for i:=0 to NumFormats.Count-1 do
  begin
    data := PtrInt(NumFormats.Objects[i]);
    if (data >= USER_OFFSET) then
      cat := TsNumFormatCategory(data - USER_OFFSET)
    else
      cat := TsNumFormatCategory(data - BUILTIN_OFFSET);
    if cat = nfcCurrency then
    begin
      nfp := CreateNumFormatParams(NumFormats.Strings[i], FWorkbook.FormatSettings);
      if (nfp <> nil) then
        try
          nfp.SetCurrSymbol(cs);
        finally
          nfp.Free;
        end;
    end;
  end;
  SelectCategory(TsNumFormatCategory(LbCategory.ItemIndex));  // to rebuild the format listbox
end;

procedure TNumFormatForm.ReplaceDecs;
var
  nfp: TsNumFormatParams;
begin
  if EdDecimals.Text = '' then
    exit;

  nfp := CreateNumFormatParams(EdNumFormatStr.Text, FWorkbook.FormatSettings);
  try
    nfp.SetDecimals(EdDecimals.Value);
    EdNumFormatStr.Text := nfp.NumFormatStr;
    UpdateSample(nfp);
  finally
    nfp.Free;
  end;
end;

procedure TNumFormatForm.SelectCategory(ACategory: TsNumFormatCategory);
var
  nfp: TsNumFormatParams;
  i, digits, numdigits: Integer;
  data: PtrInt;
  s: String;
  genvalue: Double;
  cat: TsNumFormatCategory;
begin
  LbCategory.ItemIndex := ord(ACategory);
  with LbFormat.Items do
  begin
    Clear;
    for i:=0 to NumFormats.Count-1 do
    begin
      data := PtrInt(NumFormats.Objects[i]);
      if data >= USER_OFFSET then
        cat := TsNumFormatCategory(data - USER_OFFSET)
      else
        cat := TsNumFormatCategory(data - BUILTIN_OFFSET);
      if cat = ACategory then
      begin
        nfp := CreateNumFormatParams(NumFormats.Strings[i], FWorkbook.FormatSettings);
        try
          genValue := FGenerator[ACategory];
          if nfkTimeInterval in nfp.Sections[0].Kind then
            genvalue := genValue + 1.0;
          if ACategory = nfcFraction then
          begin
            digits := nfp.Sections[0].FracInt;
            numdigits := nfp.Sections[0].FracDenominator;
            genvalue := 1.0 / (IntPower(10, numdigits) - 3);
            if digits <> 0 then genvalue := -(1234 + genValue);
          end;
          s := ConvertFloatToStr(genValue, nfp, FWorkbook.FormatSettings);
          if s = '' then s := 'General';
          LbFormat.Items.AddObject(s, TObject(PtrInt(i)));
        finally
          nfp.Free;
        end;
      end;
    end;
    {
    for i:=0 to FNumFormatLists[ACategory].Count-1 do
    begin
      nfp := FNumFormatLists[ACategory].Items[i];
      genvalue := FGenerator[ACategory];
      if nfkTimeInterval in nfp.Sections[0].Kind then
        genvalue := genValue + 1.0;
      if (ACategory = nfcFraction) then begin
        digits := nfp.Sections[0].FracInt;
        numdigits := nfp.Sections[0].FracDenominator;
        genvalue := 1 / (IntPower(10, numdigits) - 3);
        if digits <> 0 then genvalue := -(1234 + genValue);
      end;
      s := ConvertFloatToStr(genvalue, nfp, FWorkbook.FormatSettings);
      if s = '' then s := 'General';
      Add(s);
    end;
    }
  end;
  CurrSymbolPanel.Visible := (ACategory = nfcCurrency);
  GbOptions.Visible := not (ACategory in [nfcDate, nfcTime]);
end;

procedure TNumFormatForm.SelectFormat(AIndex: Integer);
var
  nfp: TsNumFormatParams;
begin
  if LbCategory.ItemIndex = -1 then
    exit;

  LbFormat.ItemIndex := AIndex;
  if AIndex >= 0 then begin
    FNumFormatStrOfList := NumFormats.Strings[PtrInt(LbFormat.Items.Objects[AIndex])];
    nfp := CreateNumFormatParams(FNumFormatStrOfList, FWorkbook.FormatSettings);
    try
      UpdateControls(nfp);
    finally
      nfp.Free;
    end;
  end;
end;

procedure TNumFormatForm.SetData(ANumFormatStr: String; AWorkbook: TsWorkbook;
  ASample: Double);
var
  cs: String;
begin
  FWorkbook := AWorkbook;
  cs := FWorkbook.FormatSettings.CurrencyString;
  if (cs = '?') or (cs = '') then
    cs := DefaultFormatSettings.CurrencyString;
  CbCurrSymbol.ItemIndex := CbCurrSymbol.Items.IndexOf(cs);

  FSampleValue := ASample;
  InitNumFormats(FWorkbook.FormatSettings);
  SetNumFormatStr(ANumFormatStr);
end;

procedure TNumFormatForm.SetNumFormatStr(const AValue: String);
var
  nfs: String;
  nfp: TsNumFormatParams;
  cat: TsNumFormatCategory;
  i: Integer;
begin
  if AValue = '' then
    i := NumFormats.IndexOf('General')
  else
    i := NumFormats.IndexOf(AValue);
  if i = -1 then
    exit;

  nfs := NumFormats.Strings[i];
  nfp := CreateNumFormatParams(nfs, FWorkbook.FormatSettings);
  try
    if nfkPercent in nfp.Sections[0].Kind then
      cat := nfcPercent
    else
    if nfkExp in nfp.Sections[0].Kind then
      cat := nfcScientific
    else
    if nfkCurrency in nfp.Sections[0].Kind then
      cat := nfcCurrency
    else
    if nfkFraction in nfp.Sections[0].Kind then
      cat := nfcFraction
    else
    if nfkDate in nfp.Sections[0].Kind then
      cat := nfcDate
    else
    if (nfp.Sections[0].Kind * [nfkDate, nfkTime] = [nfkTime]) then
      cat := nfcTime
    else
      cat := nfcNumber;
    SelectCategory(cat);
    SelectFormat(FindNumFormat(cat, AValue));
    UpdateControls(nfp);
    ReplaceCurrSymbol;
  finally
    nfp.Free;
  end;
end;

procedure TNumFormatForm.UpdateControls(ANumFormatParams: TsNumFormatParams);
var
  cs: String;
  i: Integer;
begin
  if ANumFormatParams = nil then
  begin
    EdNumFormatStr.Text := 'General';
    GbOptions.Hide;
  end else
  begin
    EdNumFormatStr.Text := ANumFormatParams.NumFormatStr;
    if (ANumFormatParams.Sections[0].Kind * [nfkDate, nfkTime] <> []) then
      GbOptions.Hide
    else begin
      GbOptions.Show;
      inc(FLockCount);
      EdDecimals.Value := ANumFormatParams.Sections[0].Decimals;
      CbNegRed.Checked := (Length(ANumFormatParams.Sections) > 1) and
                          (ANumFormatParams.Sections[1].Color = scRed);
      CbThousandSep.Checked := nfkHasThSep in ANumFormatParams.Sections[0].Kind;
      dec(FLockCount);
    end;
    if (nfkCurrency in ANumFormatParams.Sections[0].Kind) then
    begin
      cs := ANumFormatParams.Sections[0].CurrencySymbol;
      if cs <> '' then
      begin
        i := CbCurrSymbol.Items.IndexOf(cs);
        if i = -1 then begin
          RegisterCurrency(cs);
          i := CbCurrSymbol.Items.Add(cs);
        end;
        CbCurrSymbol.ItemIndex := i;
      end;
    end;
  end;
  UpdateSample(ANumFormatParams);
end;

procedure TNumFormatForm.UpdateSample(ANumFormatParams: TsNumFormatParams);
begin
  if (FSampleValue < 0) and
     (Length(ANumFormatParams.Sections) > 1) and
     (ANumFormatParams.Sections[1].Color = scRed)
  then
    Sample.Font.Color := clRed
  else
    Sample.Font.Color := clWindowText;

  Sample.Caption := ConvertFloatToStr(FSampleValue, ANumFormatParams,
    FWorkbook.FormatSettings);

  BtnAddFormat.Enabled := (EdNumFormatStr.Text <> FNumFormatStrOfList);
end;


initialization

finalization
  DestroyNumFormats;

end.

