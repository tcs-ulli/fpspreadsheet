unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ComCtrls, ExtCtrls, db, dbf, fpspreadsheet, fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    Bevel1: TBevel;
    Bevel2: TBevel;
    BtnCreateDbf: TButton;
    BtnExport: TButton;
    EdRecordCount: TEdit;
    InfoLabel2: TLabel;
    HeaderLabel1: TLabel;
    InfoLabel1: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    HeaderLabel2: TLabel;
    PageControl: TPageControl;
    Panel1: TPanel;
    RgFileFormat: TRadioGroup;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    procedure BtnCreateDbfClick(Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
    FDataset: TDbf;
    FWorkbook: TsWorkbook;
    FHeaderTemplateCell: PCell;
    FDateTemplateCell: PCell;
    procedure WriteCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
      var AValue: variant; var AStyleCell: PCell);
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

const
  NUM_LAST_NAMES = 8;
  NUM_FIRST_NAMES = 8;
  NUM_CITIES = 10;
  LAST_NAMES: array[0..NUM_LAST_NAMES-1] of string = (
    'Chaplin', 'Washington', 'Dylan', 'Springsteen', 'Brando',
    'Monroe', 'Dean', 'Lincoln');
  FIRST_NAMES: array[0..NUM_FIRST_NAMES-1] of string = (
    'Charley', 'George', 'Bob', 'Bruce', 'Marlon',
    'Marylin', 'James', 'Abraham');
  CITIES: array[0..NUM_CITIES-1] of string = (
    'New York', 'Los Angeles', 'San Francisco', 'Chicago', 'Miami',
    'New Orleans', 'Washington', 'Boston', 'Seattle', 'Las Vegas');


{ TForm1 }

{ This procedure creates a test database table with random data for us to play with }
procedure TForm1.BtnCreateDbfClick(Sender: TObject);
var
  i: Integer;
  startDate: TDate;
  maxAge: Integer = 80 * 365;
begin
  if FDataset <> nil then
    FDataset.Free;

  ForceDirectories('data');
  startDate := EncodeDate(2010, 8, 1);

  FDataset := TDbf.Create(self);
  FDataset.FilePathFull := 'data' + DirectorySeparator;
  FDataset.TableName := 'people.dbf';
  FDataset.FieldDefs.Add('Last name', ftString);
  FDataset.FieldDefs.Add('First name', ftString);
  FDataset.FieldDefs.Add('City', ftString);
  FDataset.FieldDefs.Add('Birthday', ftDateTime);
  DeleteFile(FDataset.FilePathFull + FDataset.TableName);
  FDataset.CreateTable;

  FDataset.Open;
  for i:=1 to StrToInt(EdRecordCount.Text) do begin
    if (i mod 25) = 0 then begin
      InfoLabel1.Caption := Format('Adding record %d...', [i]);
      Application.ProcessMessages;
    end;
    FDataset.Insert;
    FDataset.FieldByName('Last name').AsString := LAST_NAMES[Random(NUM_LAST_NAMES)];
    FDataset.FieldByName('First name').AsString := FIRST_NAMES[Random(NUM_FIRST_NAMES)];
    FDataset.FieldByName('City').AsString := CITIES[Random(NUM_CITIES)];
    FDataset.FieldByName('Birthday').AsDateTime := startDate - random(maxAge);
      // creates a random date between "startDate" and "maxAge" days back
    FDataset.Post;
  end;
  FDataset.Close;

  InfoLabel1.Caption := Format('Done. Created file "%s" in folder "data".', [
    FDataset.TableName, FDataset.FilePathFull
  ]);
  InfoLabel2.Caption := '';
end;

procedure TForm1.BtnExportClick(Sender: TObject);
const
  FILE_FORMATS: array[0..4] of TsSpreadsheetFormat = (
    sfExcel2, sfExcel5, sfExcel8, sfOOXML, sfOpenDocument
  );
  EXT: array[0..4] of string = (
    '_excel2.xls', '_excel5.xls', '.xls', '.xlsx', '.ods');
var
  fn: String;
  worksheet: TsWorksheet;
begin
  InfoLabel2.Caption := '';
  Application.ProcessMessages;

  if FDataset = nil then begin
    FDataset := TDbf.Create(self);
    FDataset.FilePathFull := 'data' + DirectorySeparator;
    FDataset.TableName := 'people.dbf';
  end;

  fn := FDataset.FilePathFull + FDataset.TableName;
  if not FileExists(fn) then begin
    MessageDlg(Format('Database file "%s" not found. Please run "Create database" first.',
      [fn]), mtError, [mbOK], 0);
    exit;
  end;

  FDataset.Open;

  FWorkbook := TsWorkbook.Create;
  try
    worksheet := FWorkbook.AddWorksheet(FDataset.TableName);

    // Make header line frozen
    worksheet.Options := worksheet.Options + [soHasFrozenPanes];
    worksheet.TopPaneHeight := 1;

    // Prepare template for header line
    FHeaderTemplateCell := worksheet.GetCell(0, 0);
    worksheet.WriteFontStyle(FHeaderTemplateCell, [fssBold]);
    worksheet.WriteFontColor(FHeaderTemplateCell, scWhite);
    worksheet.WriteBackgroundColor(FHeaderTemplateCell, scGray);

    // Prepare template for date column
    FDateTemplateCell := worksheet.GetCell(0, 1);
    worksheet.WriteDateTimeFormat(FDateTemplateCell, nfShortDate);

    // Make first three columns a bit wider
    worksheet.WriteColWidth(0, 20);
    worksheet.WriteColWidth(1, 20);
    worksheet.WriteColWidth(2, 20);

    // Setup virtual mode
//    FWorkbook.Options := FWorkbook.Options + [boVirtualMode, boBufStream];
    FWorkbook.Options := FWorkbook.Options + [boVirtualMode];
    FWorkbook.OnWriteCellData := @WriteCellDataHandler;
    FWorkbook.VirtualRowCount := FDataset.RecordCount + 1;  // +1 for the header line
    FWorkbook.VirtualColCount := FDataset.FieldCount;

    // Write
    fn := ChangeFileExt(fn, EXT[RgFileFormat.ItemIndex]);
    FWorkbook.WriteToFile(fn, FILE_FORMATS[RgFileFormat.ItemIndex], true);
  finally
    FreeAndNil(FWorkbook);
  end;

  InfoLabel2.Caption := Format('Done. Database exported to file "%s" in folder "%s"',
    [ChangeFileExt(FDataset.TableName, EXT[RgFileFormat.ItemIndex]), FDataset.FilePathFull]);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  InfoLabel1.Caption := '';
  InfoLabel2.Caption := '';
  PageControl.ActivePageIndex := 0;
end;

procedure TForm1.WriteCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
  var AValue: variant; var AStyleCell: PCell);
begin
  // Header line: we want to show the field names here.
  if ARow = 0 then begin
    AValue := FDataset.Fields[ACol].FieldName;
    AStyleCell := FHeaderTemplateCell;
    FDataset.First;
  end else begin
    AValue := FDataset.Fields[ACol].Value;
    if FDataset.Fields[ACol].DataType = ftDate then
      AStyleCell := FDateTemplateCell;
    if ACol = FWorkbook.VirtualColCount-1 then begin
      FDataset.Next;
      if (ARow-1) mod 25 = 0 then begin
        InfoLabel1.Caption := Format('Writing record %d...', [ARow-1]);
        Application.ProcessMessages;
      end;
    end;
  end;
end;

end.

