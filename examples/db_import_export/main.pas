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
    Bevel3: TBevel;
    BtnCreateDbf: TButton;
    BtnExport: TButton;
    BtnImport: TButton;
    EdRecordCount: TEdit;
    HeaderLabel3: TLabel;
    InfoLabel2: TLabel;
    HeaderLabel1: TLabel;
    InfoLabel1: TLabel;
    InfoLabel3: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    HeaderLabel2: TLabel;
    FileList: TListBox;
    Label3: TLabel;
    PageControl: TPageControl;
    Panel1: TPanel;
    RgFileFormat: TRadioGroup;
    TabDataGenerator: TTabSheet;
    TabExport: TTabSheet;
    TabImport: TTabSheet;
    procedure BtnCreateDbfClick(Sender: TObject);
    procedure BtnExportClick(Sender: TObject);
    procedure BtnImportClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
  private
    { private declarations }
    FExportDataset: TDbf;
    FImportDataset: TDbf;
    FWorkbook: TsWorkbook;
    FHeaderTemplateCell: PCell;
    FDateTemplateCell: PCell;
    FImportedFieldNames: TStringList;
    FImportedFieldTypes: Array of TFieldType;
    // For reading: all data for the database is generated here out of the spreadsheet file
    procedure ReadCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
      const ADataCell: PCell);
    // For writing: all data for the cells is generated here (out of the .dbf file)
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

  TABLENAME = 'people.dbf'; //name for the dbf table
  DATADIR = 'data'; //subdirectory where .dbf is stored

  // File formats corresponding to the items of the RgFileFormat radiogroup
  FILE_FORMATS: array[0..4] of TsSpreadsheetFormat = (
    sfExcel2, sfExcel5, sfExcel8, sfOOXML, sfOpenDocument
  );
  // Spreadsheet files will get the TABLENAME and have one of these extensions.
  FILE_EXT: array[0..4] of string = (
    '_excel2.xls', '_excel5.xls', '.xls', '.xlsx', '.ods');


{ TForm1 }

{ This procedure creates a test dbf table with random data for us to play with }
procedure TForm1.BtnCreateDbfClick(Sender: TObject);
var
  i: Integer;
  startDate: TDate;
  maxAge: Integer = 80 * 365;
begin
  if FExportDataset <> nil then
    FExportDataset.Free;

  ForceDirectories(DATADIR);
  startDate := EncodeDate(2010, 8, 1);

  FExportDataset := TDbf.Create(self);
  FExportDataset.FilePathFull := DATADIR + DirectorySeparator;
  FExportDataset.TableName := TABLENAME;
  FExportDataset.TableLevel := 4; //DBase IV; most widely used.
  FExportDataset.FieldDefs.Add('Last name', ftString);
  FExportDataset.FieldDefs.Add('First name', ftString);
  FExportDataset.FieldDefs.Add('City', ftString);
  FExportDataset.FieldDefs.Add('Birthday', ftDateTime);
  DeleteFile(FExportDataset.FilePathFull + FExportDataset.TableName);
  FExportDataset.CreateTable;

  FExportDataset.Open;
  for i:=1 to StrToInt(EdRecordCount.Text) do begin
    if (i mod 25) = 0 then
    begin
      InfoLabel1.Caption := Format('Adding record %d...', [i]);
      Application.ProcessMessages;
    end;
    FExportDataset.Insert;
    FExportDataset.FieldByName('Last name').AsString := LAST_NAMES[Random(NUM_LAST_NAMES)];
    FExportDataset.FieldByName('First name').AsString := FIRST_NAMES[Random(NUM_FIRST_NAMES)];
    FExportDataset.FieldByName('City').AsString := CITIES[Random(NUM_CITIES)];
    FExportDataset.FieldByName('Birthday').AsDateTime := startDate - random(maxAge);
      // creates a random date between "startDate" and "maxAge" days back
    FExportDataset.Post;
  end;
  FExportDataset.Close;

  InfoLabel1.Caption := Format('Done. Created file "%s" in folder "data".', [
    FExportDataset.TableName, FExportDataset.FilePathFull
  ]);
  InfoLabel2.Caption := '';
  InfoLabel3.Caption := '';
end;

procedure TForm1.BtnExportClick(Sender: TObject);
var
  DataFileName: String;
  worksheet: TsWorksheet;
begin
  InfoLabel2.Caption := '';
  Application.ProcessMessages;

  if FExportDataset = nil then
  begin
    FExportDataset := TDbf.Create(self);
    FExportDataset.FilePathFull := DATADIR + DirectorySeparator;
    FExportDataset.TableName := TABLENAME;
  end;

  DataFileName := FExportDataset.FilePathFull + FExportDataset.TableName;
  if not FileExists(DataFileName) then
  begin
    MessageDlg(Format('Database file "%s" not found. Please run "Create database" first.',
      [DataFileName]), mtError, [mbOK], 0);
    exit;
  end;

  FExportDataset.Open;

  FWorkbook := TsWorkbook.Create;
  try
    worksheet := FWorkbook.AddWorksheet(FExportDataset.TableName);

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

    // Setup virtual mode to save memory
//    FWorkbook.Options := FWorkbook.Options + [boVirtualMode, boBufStream];
    FWorkbook.Options := FWorkbook.Options + [boVirtualMode];
    FWorkbook.OnWriteCellData := @WriteCellDataHandler;
    FWorkbook.VirtualRowCount := FExportDataset.RecordCount + 1;  // +1 for the header line
    FWorkbook.VirtualColCount := FExportDataset.FieldCount;

    // Write
    DataFileName := ChangeFileExt(DataFileName, FILE_EXT[RgFileFormat.ItemIndex]);
    FWorkbook.WriteToFile(DataFileName, FILE_FORMATS[RgFileFormat.ItemIndex], true);
  finally
    FreeAndNil(FWorkbook);
  end;

  InfoLabel2.Caption := Format('Done. Database exported to file "%s" in folder "%s"', [
    ChangeFileExt(FExportDataset.TableName, FILE_EXT[RgFileFormat.ItemIndex]),
    FExportDataset.FilePathFull
  ]);
end;

{ This procedure imports the contents of the selected spreadsheet file into a
  new dbf database file using virtual mode. }
procedure TForm1.BtnImportClick(Sender: TObject);
var
  DataFileName: String;
  fmt: TsSpreadsheetFormat;
  ext: String;
begin
  if FileList.ItemIndex = -1 then begin
    MessageDlg('Please select a file in the listbox.', mtInformation, [mbOK], 0);
    exit;
  end;

  // Determine the file format from the filename - just to avoid the annoying
  // exceptions that occur for Excel2 and Excel5.
  DataFileName := FileList.Items[FileList.ItemIndex];
  ext := lowercase(ExtractFileExt(DataFileName));
  if ext = '.xls' then begin
    if pos(FILE_EXT[0], DataFileName) > 0 then
      fmt := sfExcel2
    else
    if pos(FILE_EXT[1], DataFileName) > 0 then
      fmt := sfExcel5
    else
      fmt := sfExcel8;
  end else
  if ext = '.xlsx' then
    fmt := sfOOXML
  else
  if ext = '.ods' then
    fmt := sfOpenDocument
  else begin
    MessageDlg('Unknown spreadsheet file format.', mtError, [mbOK], 0);
    exit;
  end;

  DataFileName := DATADIR + DirectorySeparator + DataFileName;

  // Prepare dbf table for the spreadsheet data to be imported
  if FImportDataset <> nil then
    FImportDataset.Free;
  FImportDataset := TDbf.Create(self);
  FImportDataset.FilePathFull := DATADIR + DirectorySeparator;
  FImportDataset.TableName := 'imported_' + TABLENAME;
  FImportDataset.TableLevel := 4; //DBase IV; most widely used.
  DeleteFile(FImportDataset.FilePathFull + FImportDataset.TableName);

  // The stringlist will temporarily store the field names ...
  if FImportedFieldNames = nil then
    FImportedFieldNames := TStringList.Create;
  FImportedFieldNames.Clear;

  // ... and this array stores the field types until we have all information
  // to create the dbf table.
  SetLength(FImportedFieldTypes, 0);

  // Create the workbook and activate virtual mode
  FWorkbook := TsWorkbook.Create;
  try
    FWorkbook.Options := FWorkbook.Options + [boVirtualMode];
    FWorkbook.OnReadCellData := @ReadCellDataHandler;
    // Read the data from the spreadsheet file transparently into the dbf file
    // The data are not permanently available in the worksheet and do occupy
    // memory there - this is virtual mode.
    FWorkbook.ReadFromFile(DataFilename, fmt);
  finally
    FWorkbook.Free;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  InfoLabel1.Caption := '';
  InfoLabel2.Caption := '';
  InfoLabel3.Caption := '';
  PageControl.ActivePageIndex := 0;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FImportedFieldNames);
end;

{ When we activate the "Import" page of the pagecontrol we read the data
  folder and collect all spreadsheet files available in a list box. The user
  will have to select the one to be converted to dbf. }
procedure TForm1.PageControlChange(Sender: TObject);
var
  sr: TSearchRec;
  ext: String;
begin
  if PageControl.ActivePage = TabImport then begin
    FileList.Clear;
    if FindFirst(DATADIR + DirectorySeparator + ChangeFileExt(TABLENAME, '') + '*.*', faAnyFile, sr) = 0
    then begin
      repeat
        if (sr.Name = '.') or (sr.Name = '..') then
          Continue;
        ext := lowercase(ExtractFileExt(sr.Name));
        if (ext = '.xls') or (ext = '.xlsx') or (ext = '.ods') then
          FileList.Items.Add(sr.Name);
      until FindNext(sr) <> 0;
      FindClose(sr);
    end;
  end;
end;

{ This is the event handler for reading a spreadsheet file in virtual mode.
  The data are not stored in the worksheet and exist only temporarily.
  This event handler picks the data and posts them to the database table.
  Note that we do not make many assumptions on the data structure here. Therefore
  we have to buffer the first two rows of the spreadsheet file until the
  structure of the table is clear. }
procedure TForm1.ReadCellDataHandler(Sender: TObject; ARow, Acol: Cardinal;
  const ADataCell: PCell);
var
  i: Integer;
begin
  // The first row (index 0) holds the field names. We temporarily store the
  // field names in a string list because we don't know the data types of the
  // cell before we have not read the second row (index 1).
  if ARow = 0 then begin
    // We know that the first row contains string cells -> no further checks.
    FImportedFieldNames.Add(ADataCell^.UTF8StringValue);
  end
  else
  // We have to buffer the second row (index 1) as well. When it is fully read
  // we can put everything together and create the dfb table.
  if ARow = 1 then begin
    if Length(FImportedFieldTypes) = 0 then
      SetLength(FImportedFieldTypes, FImportedFieldNames.Count);
    case ADataCell^.ContentType of
      cctNumber    : FImportedFieldTypes[ACol] := ftFloat;
      cctUTF8String: FImportedFieldTypes[ACol] := ftString;
      cctDateTime  : FImportedFieldTypes[ACol] := ftDate;
    end;
    // All field types are known --> we create the table
    if ACol = High(FImportedFieldTypes) then begin
      for i:=0 to High(FImportedFieldTypes) do
        FImportDataset.FieldDefs.Add(FImportedFieldNames[i], FImportedFieldTypes[i]);
      DeleteFile(FImportDataset.FilePathFull + FImportDataset.TableName);
      FImportDataset.CreateTable;
      FImportDataset.Open;
    end;
  end
  else
  begin
    // Now that we know everything we can add the data to the table
    if ACol = 0 then
      FImportDataset.Insert;
    case ADataCell^.ContentType of
      cctNumber    : FImportDataSet.Fields[Acol].AsFloat :=  ADataCell^.NumberValue;
      cctUTF8String: FImportDataset.Fields[Acol].AsString := ADataCell^.UTF8StringValue;
      cctDateTime  : FImportDataset.Fields[ACol].AsDateTime := ADataCell^.DateTimeValue;
    end;
    if ACol = High(FImportedFieldTypes) then
      FImportDataset.Post;
  end;
end;

{ This is the event handler for exporting a database file to spreadsheet format
  in virtual mode. Data are not written into the worksheet, they exist only
  temporarily. }
procedure TForm1.WriteCellDataHandler(Sender: TObject; ARow, ACol: Cardinal;
  var AValue: variant; var AStyleCell: PCell);
begin
  // Header line: we want to show the field names here.
  if ARow = 0 then
  begin
    AValue := FExportDataset.Fields[ACol].FieldName;
    AStyleCell := FHeaderTemplateCell;
    FExportDataset.First;
  end
  else
  begin
    AValue := FExportDataset.Fields[ACol].Value;
    if FExportDataset.Fields[ACol].DataType = ftDate then
      AStyleCell := FDateTemplateCell;
    if ACol = FWorkbook.VirtualColCount-1 then
    begin
      FExportDataset.Next;
      if (ARow-1) mod 25 = 0 then
      begin
        InfoLabel1.Caption := Format('Writing record %d...', [ARow-1]);
        Application.ProcessMessages;
      end;
    end;
  end;
end;

end.

