unit sHyperlinkForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ButtonPanel,
  ExtCtrls, Buttons, StdCtrls, ComCtrls,
  fpsTypes, fpspreadsheet;

type

  { THyperlinkForm }

  THyperlinkForm = class(TForm)
    Bevel1: TBevel;
    BtnBrowseFile: TButton;
    ButtonPanel1: TButtonPanel;
    CbFileName1: TComboBox;
    CbFileName2: TComboBox;
    CbFileName3: TComboBox;
    CbWorksheets: TComboBox;
    CbCellAddress: TComboBox;
    CbFileName: TComboBox;
    CbMailRecipient: TComboBox;
    EdMailSubject: TEdit;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    GbMailRecipient: TGroupBox;
    GroupBox8: TGroupBox;
    Images: TImageList;
    HyperlinkInfo: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Notebook: TNotebook;
    OpenDialog: TOpenDialog;
    PgInternal: TPage;
    Page2: TPage;
    Page3: TPage;
    Page4: TPage;
    Panel2: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    ToolBar: TToolBar;
    TbInternal: TToolButton;
    TbFile: TToolButton;
    TbInternet: TToolButton;
    TbMail: TToolButton;
    procedure BtnBrowseFileClick(Sender: TObject);
    procedure CbCellAddressEditingDone(Sender: TObject);
    procedure CbMailRecipientEditingDone(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
    procedure ToolButtonClick(Sender: TObject);
    procedure UpdateHyperlinkInfo(Sender: TObject);
  private
    { private declarations }
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
    function GetHyperlinkTarget: String;
    function GetHyperlinkTooltip: String;
    procedure SetHyperlinkKind(AValue: Integer);
    procedure SetHyperlinkTarget(const AValue: String);
    procedure SetHyperlinkTooltip(const AValue: String);
    procedure SetWorksheet(AWorksheet: TsWorksheet);
  protected
    function GetHyperlinkKind: Integer;
    function ValidData(out AControl: TWinControl; out AMsg: String): Boolean;
  public
    { public declarations }
    procedure GetHyperlink(out AHyperlink: TsHyperlink);
    procedure SetHyperlink(AWorksheet: TsWorksheet; const AHyperlink: TsHyperlink);
  end;

var
  HyperlinkForm: THyperlinkForm;

implementation

{$R *.lfm}

uses
  URIParser,
  fpsUtils;

const
  TAG_INTERNAL = 0;
  TAG_FILE = 1;
  TAG_INTERNET = 2;
  TAG_MAIL = 3;

{ THyperlinkForm }

procedure THyperlinkForm.BtnBrowseFileClick(Sender: TObject);
begin
  with OpenDialog do begin
    Filename := CbFileName.Text;
    if Execute then begin
      CbFileName.Text := FileName;
      if CbFileName.Items.IndexOf(FileName) = -1 then
        CbFilename.Items.Add(FileName);
    end;
  end;
end;

procedure THyperlinkForm.CbCellAddressEditingDone(Sender: TObject);
begin
  CbCellAddress.Text := Uppercase(CbCellAddress.Text);
end;

procedure THyperlinkForm.CbMailRecipientEditingDone(Sender: TObject);
begin
  if (CbMailRecipient.Text <> '') and
     (CbMaiLRecipient.Items.IndexOf(CbMailRecipient.Text) = -1)
  then
    CbMailRecipient.Items.Insert(0, CbMailRecipient.Text);
end;

procedure THyperlinkForm.GetHyperlink(out AHyperlink: TsHyperlink);
begin
  AHyperlink.Target := GetHyperlinkTarget;
  AHyperlink.Tooltip := GetHyperlinkTooltip;
end;

function THyperlinkForm.GetHyperlinkKind: Integer;
begin
  for Result := 0 to Toolbar.ButtonCount-1 do
    if Toolbar.Buttons[Result].Down then
      exit;
  Result := -1;
end;

function THyperlinkForm.GetHyperlinkTarget: String;
begin
  case GetHyperlinkKind of
    TAG_INTERNAL:
      begin //internal
        if (CbWorksheets.ItemIndex > 0) and (CbCellAddress.Text <> '') then
          Result := '#' + CbWorksheets.Text + '!' + Uppercase(CbCellAddress.Text)
        else if (CbWorksheets.ItemIndex > 0) then
          Result := '#' + CbWorksheets.Text + '!'
        else if (CbCellAddress.Text <> '') then
          Result := '#' + Uppercase(CbCellAddress.Text)
        else
          Result := '';
      end;
    TAG_FILE:
      begin  // File
        if (FWorkbook = nil) or (FWorkbook.FileName = '') then
          Result := FilenameToURI(CbFilename.Text)
        else
          Result := '';
      end;
    TAG_INTERNET:
      ;
    TAG_MAIL:
      begin  // Mail
        if EdMailSubject.Text <> '' then
          Result := Format('mailto:%s?subject=%s', [CbMailRecipient.Text, EdMailSubject.Text])
        else
          Result := Format('mailto:%s', [CbMailRecipient.Text]);
      end;
  end;
end;

function THyperlinkForm.GetHyperlinkTooltip: String;
begin
  //
end;

procedure THyperlinkForm.OKButtonClick(Sender: TObject);
var
  C: TWinControl;
  msg: String;
begin
  if not ValidData(C, msg) then begin
    C.SetFocus;
    MessageDlg(msg, mtError, [mbOK], 0);
    ModalResult := mrNone;
  end;
end;

procedure THyperlinkForm.SetHyperlink(AWorksheet: TsWorksheet;
  const AHyperlink: TsHyperlink);
begin
  SetWorksheet(AWorksheet);
  SetHyperlinkTarget(AHyperlink.Target);
  SetHyperlinkTooltip(AHyperlink.Tooltip);
end;

procedure THyperlinkForm.SetHyperlinkKind(AValue: Integer);
var
  i: Integer;
begin
  for i:=0 to Toolbar.ButtonCount-1 do
    Toolbar.Buttons[i].Down := (AValue = Toolbar.Buttons[i].Tag);
  Notebook.PageIndex := AValue;
end;

procedure THyperlinkForm.SetHyperlinkTarget(const AValue: String);
var
  u: TURI;
  sheet: TsWorksheet;
  c,r: Cardinal;
  i, idx: Integer;
  p: Integer;
begin
  if AValue = '' then
  begin
    CbWorksheets.ItemIndex := 0;
    CbCellAddress.Text := '';

    CbMailRecipient.Text := '';
    EdMailSubject.Text := '';

    UpdateHyperlinkInfo(nil);
    exit;
  end;

  // Internal link
  if pos('#', AValue) = 1 then begin
    SetHyperlinkKind(TAG_INTERNAL);
    if FWorkbook.TryStrToCell(Copy(AValue, 2, Length(AValue)), sheet, r, c) then
    begin
      if (sheet = nil) or (sheet = FWorksheet) then
        CbWorksheets.ItemIndex := 0
      else
      begin
        idx := 0;
        for i:=1 to CbWorksheets.Items.Count-1 do
          if CbWorksheets.Items[i] = sheet.Name then
          begin
            idx := i;
            break;
          end;
        CbWorksheets.ItemIndex := idx;
      end;
      CbCellAddress.Text := GetCellString(r, c);
      UpdateHyperlinkInfo(nil);
    end else begin
      HyperlinkInfo.Caption := AValue;
      MessageDlg(Format('Sheet not found in hyperlink "%s"', [AValue]), mtError,
        [mbOK], 0);
    end;
    exit;
  end;

  // external links
  u := ParseURI(AValue);

  // Mail
  if SameText(u.Protocol, 'mailto') then
  begin
    SetHyperlinkKind(TAG_MAIL);
    CbMailRecipient.Text := u.Document;
    if CbMailRecipient.Items.IndexOf(u.Document) = -1 then
      CbMailRecipient.Items.Insert(0, u.Document);
    if (u.Params <> '') then
    begin
      p := pos('subject=', u.Params);
      if p <> 0 then
        EdMailSubject.Text := copy(u.Params, p+Length('subject='), MaxInt);
    end;
    UpdateHyperlinkInfo(nil);
    exit;
  end;
end;

procedure THyperlinkForm.SetHyperlinkTooltip(const AValue: String);
begin
  //
end;

procedure THyperlinkForm.SetWorksheet(AWorksheet: TsWorksheet);
var
  i: Integer;
begin
  FWorksheet := AWorksheet;
  if FWorksheet = nil then
    raise Exception.Create('[THyperlinkForm.SetWorksheet] Worksheet cannot be nil.');
  FWorkbook := FWorksheet.Workbook;

  CbWorksheets.Items.Clear;
  CbWorksheets.Items.Add('(current worksheet)');
  for i:=0 to FWorkbook.GetWorksheetCount-1 do
    CbWorksheets.Items.Add(FWorkbook.GetWorksheetByIndex(i).Name);
end;

procedure THyperlinkForm.ToolButtonClick(Sender: TObject);
var
  i: Integer;
begin
  Notebook.PageIndex := TToolButton(Sender).Tag;
  for i:=0 to Toolbar.ButtonCount-1 do
    Toolbar.Buttons[i].Down := Toolbar.Buttons[i].Tag = TToolbutton(Sender).Tag;
  UpdateHyperlinkInfo(nil);
end;

procedure THyperlinkForm.UpdateHyperlinkInfo(Sender: TObject);
begin
  HyperlinkInfo.Caption := GetHyperlinkTarget;
end;

function THyperlinkForm.ValidData(out AControl: TWinControl;
  out AMsg: String): Boolean;
var
  r,c: Cardinal;
begin
  Result := false;
  AMsg := '';
  AControl := nil;

  case GetHyperlinkKind of
    TAG_INTERNAL:
      begin
        if CbCellAddress.Text = '' then
        begin
          AMsg := 'No cell address specified.';
          AControl := CbCellAddress;
          exit;
        end;
        if not ParseCellString(CbCellAddress.Text, r, c) then
        begin
          AMsg := Format('"%s" is not a valid cell address.', [CbCellAddress.Text]);
          AControl := CbCellAddress;
          exit;
        end;
        if (CbWorksheets.Items.IndexOf(CbWorksheets.Text) = -1) and (CbWorksheets.ItemIndex <> 0) then
        begin
          AMsg := Format('Worksheet "%s" does not exist.', [CbWorksheets.Text]);
          AControl := CbWorksheets;
          exit;
        end;
      end;

    TAG_MAIL:
      begin
        if CbMailRecipient.Text = '' then
        begin
          AMsg := 'No mail recipient specified.';
          AControl := CbMailRecipient;
          exit;
        end;
        // Check e-mail address here also!
      end;
  end;
  Result := true;
end;

end.

