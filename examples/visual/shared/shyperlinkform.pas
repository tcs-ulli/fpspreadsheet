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
    CbFtpServer: TComboBox;
    CbFtpUsername: TComboBox;
    CbFtpPassword: TComboBox;
    CbHttpAddress: TComboBox;
    CbFileBookmark: TComboBox;
    CbWorksheets: TComboBox;
    CbCellAddress: TComboBox;
    CbFileName: TComboBox;
    CbMailRecipient: TComboBox;
    EdHttpBookmark: TEdit;
    EdTooltip: TEdit;
    EdMailSubject: TEdit;
    GroupBox2: TGroupBox;
    GbFileName: TGroupBox;
    GbInternetLinkType: TGroupBox;
    GbHttp: TGroupBox;
    GbMailRecipient: TGroupBox;
    GroupBox6: TGroupBox;
    GbFileBookmark: TGroupBox;
    GroupBox8: TGroupBox;
    GbFtp: TGroupBox;
    Images: TImageList;
    HyperlinkInfo: TLabel;
    Label1: TLabel;
    LblFtpUserName: TLabel;
    LblFtpPassword: TLabel;
    LblHttpAddress: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    LblHttpBookmark: TLabel;
    Notebook: TNotebook;
    InternetNotebook: TNotebook;
    OpenDialog: TOpenDialog;
    PgHTTP: TPage;
    PfFTP: TPage;
    PgInternal: TPage;
    PgFile: TPage;
    PgInternet: TPage;
    PgMail: TPage;
    Panel2: TPanel;
    RbFTP: TRadioButton;
    RbHTTP: TRadioButton;
    ToolBar: TToolBar;
    TbInternal: TToolButton;
    TbFile: TToolButton;
    TbInternet: TToolButton;
    TbMail: TToolButton;
    procedure BtnBrowseFileClick(Sender: TObject);
    procedure CbCellAddressEditingDone(Sender: TObject);
    procedure CbFileBookmarkDropDown(Sender: TObject);
    procedure CbFileNameEditingDone(Sender: TObject);
    procedure CbFtpServerEditingDone(Sender: TObject);
    procedure CbHttpAddressEditingDone(Sender: TObject);
    procedure CbMailRecipientEditingDone(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
    procedure HTTP_FTP_Change(Sender: TObject);
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
    procedure SetInternetLinkKind(AValue: Integer);
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

  TAG_HTTP = 0;
  TAG_FTP = 1;

{ THyperlinkForm }

procedure THyperlinkForm.BtnBrowseFileClick(Sender: TObject);
begin
  with OpenDialog do begin
    Filename := CbFileName.Text;
    if Execute then begin
      InitialDir := ExtractFileDir(FileName);
      CbFileName.Text := FileName;
      if (CbFileName.Text <> '') and (CbFileName.Items.IndexOf(FileName) = -1) then
        CbFilename.Items.Insert(0, FileName);
    end;
  end;
end;

procedure THyperlinkForm.CbCellAddressEditingDone(Sender: TObject);
begin
  CbCellAddress.Text := Uppercase(CbCellAddress.Text);
end;

procedure THyperlinkForm.CbFileBookmarkDropDown(Sender: TObject);
var
  ext: String;
  wb: TsWorkbook;
  ws: TsWorksheet;
  i: Integer;
begin
  CbFileBookmark.Items.Clear;
  if FileExists(CbFilename.Text) then begin
    ext := Lowercase(ExtractFileExt(CbFileName.Text));
    if (ext = '.xls') or (ext = '.xlsx') or (ext = '.ods') then begin
      wb := TsWorkbook.Create;
      try
        wb.ReadFromFile(CbFileName.Text);
        for i:=0 to wb.GetWorksheetCount-1 do
        begin
          ws := wb.GetWorksheetByIndex(i);
          CbFileBookmark.Items.Add(ws.Name);
        end;
      finally
        wb.Free;
      end;
    end;
  end;
end;

procedure THyperlinkForm.CbFileNameEditingDone(Sender: TObject);
begin
  if (CbFilename.Text <> '') and
     (CbFilename.Items.IndexOf(CbFilename.Text) = -1)
  then
    CbFileName.Items.Insert(0, CbFileName.Text);
end;

procedure THyperlinkForm.CbFtpServerEditingDone(Sender: TObject);
begin
  if (CbFtpServer.Text <> '') and
     (CbFtpServer.Items.IndexOf(CbFtpServer.Text) = -1)
  then
    CbFtpServer.Items.Insert(0, CbFtpServer.Text);
end;

procedure THyperlinkForm.CbHttpAddressEditingDone(Sender: TObject);
begin
  if (CbHttpAddress.Text <> '') and
     (CbHttpAddress.Items.Indexof(CbHttpAddress.Text) = -1)
  then
    CbHttpAddress.Items.Insert(0, CbHttpAddress.Text);
end;

procedure THyperlinkForm.CbMailRecipientEditingDone(Sender: TObject);
begin
  if (CbMailRecipient.Text <> '') and
     (CbMaiLRecipient.Items.IndexOf(CbMailRecipient.Text) = -1)
  then
    CbMailRecipient.Items.Insert(0, CbMailRecipient.Text);
end;

procedure THyperlinkForm.FormCreate(Sender: TObject);
begin
  HTTP_FTP_Change(nil);
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
  Result := '';
  case GetHyperlinkKind of
    TAG_INTERNAL:
      begin //internal
        if (CbWorksheets.ItemIndex > 0) and (CbCellAddress.Text <> '') then
          Result := '#' + CbWorksheets.Text + '!' + Uppercase(CbCellAddress.Text)
        else if (CbWorksheets.ItemIndex > 0) then
          Result := '#' + CbWorksheets.Text + '!'
        else if (CbCellAddress.Text <> '') then
          Result := '#' + Uppercase(CbCellAddress.Text);
      end;

    TAG_FILE:
      begin  // File
        if FileNameIsAbsolute(CbFilename.Text) then
          Result := FilenameToURI(CbFilename.Text)
        else
          Result := CbFilename.Text;
        if CbFileBookmark.Text <> '' then
          Result := Result + '#' + CbFileBookmark.Text;
      end;

    TAG_INTERNET:
      begin  // Internet link
        if RbHttp.Checked and (CbHttpAddress.Text <> '') then
        begin
          if pos('http', Lowercase(CbHttpAddress.Text)) = 1 then
            Result := CbHttpAddress.Text
          else
            Result := 'http://' + CbHttpAddress.Text;
          if EdHttpBookmark.Text <> '' then
            Result := Result + '#' + EdHttpBookmark.Text;
        end else
        if RbFtp.Checked and (CbFtpServer.Text <> '') then
        begin
          if (CbFtpUsername.Text <> '') and (CbFtpPassword.Text <> '') then
            Result := Format('ftp://%s:%s@%s', [CbFtpUsername.Text, CbFtpPassword.Text, CbFtpServer.Text])
          else
          if (CbFtpUsername.Text <> '') and (CbFtpPassword.Text = '') then
            Result := Format('ftp://%s@%s', [CbFtpUsername.Text , CbFtpServer.Text])
          else
            Result := 'ftp://anonymous@' + CbFtpServer.Text;
        end;
      end;

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
  Result := EdTooltip.Text;
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

procedure THyperlinkForm.HTTP_FTP_Change(Sender: TObject);
begin
  if RbHTTP.Checked then
    InternetNotebook.PageIndex := 0;
  if RbFTP.Checked then
    InternetNotebook.PageIndex := 1;
  UpdateHyperlinkInfo(nil);
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
  fn, bm: String;
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

  // File with absolute path
  if SameText(u.Protocol, 'file') then
  begin
    SetHyperlinkKind(TAG_FILE);
    UriToFilename(AValue, fn);
    CbFilename.Text := fn;
    CbFileBookmark.Text := u.Bookmark;
    UpdateHyperlinkInfo(nil);
    exit;
  end;

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

  // http
  if SameText(u.Protocol, 'http') or SameText(u.Protocol, 'https') then
  begin
    SetHyperlinkKind(TAG_INTERNET);
    SetInternetLinkKind(TAG_HTTP);
    CbHttpAddress.Text := u.Host;
    EdHttpBookmark.Text := u.Bookmark;
    UpdateHyperlinkInfo(nil);
    exit;
  end;

  // ftp
  if SameText(u.Protocol, 'ftp') then
  begin
    SetHyperlinkKind(TAG_INTERNET);
    SetInternetLinkKind(TAG_FTP);
    CbFtpServer.Text := u.Host;
    CbFtpUserName.text := u.UserName;
    CbFtpPassword.Text := u.Password;
    UpdateHyperlinkInfo(nil);
    exit;
  end;

  // If we get there it must be a local file with relative path
  SetHyperlinkKind(TAG_FILE);
  SplitHyperlink(AValue, fn, bm);
  CbFileName.Text := fn;
  CbFileBookmark.Text := bm;
  UpdateHyperlinkInfo(nil);
end;

procedure THyperlinkForm.SetHyperlinkTooltip(const AValue: String);
begin
  EdTooltip.Text := AValue;
end;

procedure THyperlinkForm.SetInternetLinkKind(AValue: Integer);
begin
  RbHttp.Checked := AValue = TAG_HTTP;
  RbFtp.Checked := AValue = TAG_FTP;
  InternetNotebook.PageIndex := AValue;
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
var
  s: String;
begin
  s := GetHyperlinkTarget;
  if s = '' then s := #32;
  HyperlinkInfo.Caption := s;
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

    TAG_FILE:
      begin
        if CbFilename.Text = '' then
        begin
          AMsg := 'No filename specified.';
          AControl := CbFileName;
          exit;
        end;
      end;

    TAG_INTERNET:
      if RbHttp.Checked then
      begin
        if CbHttpAddress.Text = '' then
        begin
          AMsg := 'URL of web site not specified.';
          AControl := CbHttpAddress;
          exit;
        end;
      end else
      if RbFtp.Checked then
      begin
        if CbFtpServer.Text = '' then
        begin
          AMsg := 'Ftp server not specified.';
          AControl := CbFtpServer;
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

